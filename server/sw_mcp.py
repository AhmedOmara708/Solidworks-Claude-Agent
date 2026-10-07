"""SolidWorks MCP server.

Exposes the SolidWorks COM API as tools that Claude Code can call. Runs on the
Windows machine where SolidWorks is installed (SolidWorks is Windows-only).

Conventions for every tool: lengths in millimetres, angles in degrees. The
SolidWorks API itself uses metres/radians; conversion happens here.
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import traceback
from typing import Any

try:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP
except ImportError:  # mcp 2.x
    from mcp.server.mcpserver import MCPServer as FastMCP

mcp = FastMCP("solidworks")

MM = 0.001
_sw = None

# swDocumentTypes_e
DOC_TYPES = {"part": 1, "assembly": 2, "drawing": 3}
# swUserPreferenceStringValue_e default templates
TEMPLATE_PREF = {"part": 8, "assembly": 9, "drawing": 10}
# swEndConditions_e
END_COND = {"blind": 0, "through_all": 1, "up_to_next": 2, "mid_plane": 6}


# --------------------------------------------------------------------------- COM plumbing
def _com():
    try:
        import pythoncom
        import win32com.client as wc
    except ImportError as e:  # pragma: no cover - non-Windows
        raise RuntimeError(
            "pywin32 is not available. This server must run on Windows with SolidWorks installed "
            "(pip install pywin32)."
        ) from e
    return pythoncom, wc


def _null():
    pythoncom, wc = _com()
    return wc.VARIANT(pythoncom.VT_DISPATCH, None)


def _byref_int():
    pythoncom, wc = _com()
    return wc.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)


def app():
    """Connect to (or start) SolidWorks."""
    global _sw
    _, wc = _com()
    if _sw is not None:
        try:
            _sw.RevisionNumber  # liveness probe
            return _sw
        except Exception:
            _sw = None
    try:
        _sw = wc.GetActiveObject("SldWorks.Application")
    except Exception:
        _sw = wc.Dispatch("SldWorks.Application")
    _sw.Visible = True
    return _sw


def active_model(required_type: str | None = None):
    m = app().ActiveDoc
    if m is None:
        raise RuntimeError("No active document. Use sw_new_document or sw_open_document first.")
    if required_type and m.GetType() != DOC_TYPES[required_type]:
        raise RuntimeError(f"Active document is not a {required_type}.")
    return m


def select(model, name: str, kind: str, x=0.0, y=0.0, z=0.0, append=False, mark=0) -> bool:
    """Select by name/type; x,y,z in mm (needed for EDGE/FACE selection by point)."""
    return bool(
        model.Extension.SelectByID2(name, kind, x * MM, y * MM, z * MM, append, mark, _null(), 0)
    )


def features(model) -> list[Any]:
    out, f = [], model.FirstFeature()
    while f is not None:
        out.append(f)
        f = f.GetNextFeature()
    return out


def _last_feature(model, type_name: str):
    last = None
    for f in features(model):
        if f.GetTypeName2() == type_name:
            last = f
    return last


def _ok(**kw) -> str:
    return json.dumps({"ok": True, **kw}, default=str)


def _fail(msg: str, **kw) -> str:
    return json.dumps({"ok": False, "error": msg, **kw}, default=str)


def tool(fn):
    """Register an MCP tool and turn exceptions into readable error results."""

    def wrapper(*a, **k):
        try:
            return fn(*a, **k)
        except Exception as e:
            return _fail(f"{type(e).__name__}: {e}", trace=traceback.format_exc(limit=3))

    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    wrapper.__annotations__ = fn.__annotations__
    return mcp.tool()(wrapper)


# --------------------------------------------------------------------------- documents
@tool
def sw_status() -> str:
    """Connect to SolidWorks and report version, open documents and the active document."""
    sw = app()
    docs, d = [], sw.GetFirstDocument()
    while d is not None:
        docs.append({"title": d.GetTitle(), "path": d.GetPathName(), "type": d.GetType()})
        d = d.GetNext()
    act = sw.ActiveDoc
    return _ok(
        version=sw.RevisionNumber,
        open_documents=docs,
        active=act.GetTitle() if act else None,
    )


@tool
def sw_new_document(kind: str = "part") -> str:
    """Create a new part, assembly or drawing from the user's default template. kind: part|assembly|drawing."""
    sw = app()
    tpl = sw.GetUserPreferenceStringValue(TEMPLATE_PREF[kind])
    if not tpl:
        return _fail("No default template configured in SolidWorks (Tools > Options > Default Templates).")
    m = sw.NewDocument(tpl, 0, 0, 0)
    return _ok(title=m.GetTitle()) if m else _fail("NewDocument returned nothing")


@tool
def sw_open_document(path: str) -> str:
    """Open a .sldprt/.sldasm/.slddrw (or STEP/IGES etc.) file by absolute path."""
    ext = os.path.splitext(path)[1].lower()
    kind = {".sldprt": 1, ".sldasm": 2, ".slddrw": 3}.get(ext, 1)
    err, warn = _byref_int(), _byref_int()
    m = app().OpenDoc6(path, kind, 1, "", err, warn)
    return _ok(title=m.GetTitle(), warnings=warn.value) if m else _fail(f"OpenDoc6 failed, error code {err.value}")


@tool
def sw_save(path: str = "") -> str:
    """Save the active document. If path is given, Save As to that absolute path (.sldprt/.sldasm)."""
    m = active_model()
    if path:
        r = m.SaveAs3(path, 0, 1)
    else:
        err, warn = _byref_int(), _byref_int()
        r = 0 if m.Save3(1, err, warn) else 1
    return _ok(saved=(r == 0)) if r == 0 else _fail(f"Save failed (code {r})")


@tool
def sw_close(save: bool = False) -> str:
    """Close the active document, optionally saving first."""
    m = active_model()
    title = m.GetTitle()
    if save:
        m.Save3(1, _byref_int(), _byref_int())
    app().CloseDoc(title)
    return _ok(closed=title)


@tool
def sw_export(path: str) -> str:
    """Export the active document; format chosen by extension (.step .stp .stl .iges .x_t .pdf .dxf .png .jpg .3mf). Does not change the active doc's filename."""
    m = active_model()
    r = m.SaveAs3(path, 0, 1 | 2)
    return _ok(path=path, exists=os.path.exists(path)) if r == 0 else _fail(f"Export failed (code {r})")


@tool
def sw_screenshot(path: str, view: str = "*Isometric") -> str:
    """Save a PNG of the model window so it can be inspected (read the file to see it).
    view: *Front *Back *Left *Right *Top *Bottom *Isometric *Trimetric *Dimetric."""
    m = active_model()
    m.ShowNamedView2(view, -1)
    m.ViewZoomtofit2()
    r = m.SaveAs3(path, 0, 1 | 2)
    return _ok(path=path) if r == 0 else _fail(f"Screenshot failed (code {r})")


# --------------------------------------------------------------------------- inspection
@tool
def sw_feature_tree() -> str:
    """List all features (name, type, suppressed) of the active document in order."""
    m = active_model()
    return _ok(
        features=[
            {"name": f.Name, "type": f.GetTypeName2(), "suppressed": bool(f.IsSuppressed())}
            for f in features(m)
        ]
    )


@tool
def sw_mass_properties() -> str:
    """Mass, volume, surface area, centre of mass (mm) and density of the active part/assembly."""
    m = active_model()
    p = m.Extension.CreateMassProperty()
    com = [c / MM for c in p.CenterOfMass]
    return _ok(mass_kg=p.Mass, volume_mm3=p.Volume / MM**3, surface_area_mm2=p.SurfaceArea / MM**2,
               center_of_mass_mm=com, density_kg_m3=p.Density)


@tool
def sw_get_dimension(name: str) -> str:
    """Read a dimension, e.g. 'D1@Sketch1' or 'D1@Boss-Extrude1'. Returns mm (or degrees for angular)."""
    m = active_model()
    d = m.Parameter(name)
    if d is None:
        return _fail(f"No dimension named {name}")
    v = d.SystemValue
    return _ok(name=name, value=v / MM if d.GetType() != 2 else v * 180 / 3.141592653589793)


@tool
def sw_set_dimension(name: str, value: float) -> str:
    """Set a dimension (mm, or degrees for angular dimensions) and rebuild."""
    import math

    m = active_model()
    d = m.Parameter(name)
    if d is None:
        return _fail(f"No dimension named {name}")
    d.SystemValue = value * MM if d.GetType() != 2 else math.radians(value)
    ok = m.EditRebuild3()
    return _ok(name=name, value=value, rebuilt=bool(ok))


@tool
def sw_rebuild() -> str:
    """Force a rebuild of the active document and report whether it succeeded."""
    return _ok(rebuilt=bool(active_model().ForceRebuild3(False)))


@tool
def sw_custom_properties(set_name: str = "", set_value: str = "") -> str:
    """List file-level custom properties; if set_name is given, add/overwrite it first."""
    m = active_model()
    mgr = m.Extension.CustomPropertyManager("")
    if set_name:
        mgr.Add3(set_name, 30, set_value, 1)
    names = list(mgr.GetNames() or [])
    return _ok(properties={n: mgr.Get(n) for n in names})


@tool
def sw_equations(add: str = "") -> str:
    """List equations / global variables; optionally add one, e.g. '"Width"= 40' or '"D1@Sketch1"="Width"/2'."""
    m = active_model()
    eq = m.GetEquationMgr()
    if add:
        eq.Add2(eq.GetCount(), add, True)
        m.EditRebuild3()
    return _ok(equations=[eq.Equation(i) for i in range(eq.GetCount())])


# --------------------------------------------------------------------------- modelling
@tool
def sw_select(name: str, kind: str, x: float = 0, y: float = 0, z: float = 0, append: bool = False) -> str:
    """Select an entity. kind: PLANE, SKETCH, BODYFEATURE, EDGE, FACE, VERTEX, COMPONENT, AXIS... For EDGE/FACE give a point (mm) on it in model space; name may be ''."""
    m = active_model()
    return _ok(selected=select(m, name, kind, x, y, z, append))


@tool
def sw_clear_selection() -> str:
    """Clear the current selection."""
    active_model().ClearSelection2(True)
    return _ok()


@tool
def sw_sketch(plane: str, entities: list[dict], name: str = "") -> str:
    """Create and close a 2D sketch in one call.

    plane: 'Front Plane' | 'Top Plane' | 'Right Plane' (names as shown in the feature tree) or the name of an existing plane.
    entities: list of dicts (mm, sketch-local x/y), each one of:
      {"type":"line","x1":0,"y1":0,"x2":10,"y2":0}
      {"type":"centerline","x1":0,"y1":0,"x2":0,"y2":10}     (use as revolve axis)
      {"type":"rect","x1":-10,"y1":-5,"x2":10,"y2":5}         (corner rectangle)
      {"type":"circle","x":0,"y":0,"r":5}
      {"type":"arc","cx":0,"cy":0,"x1":5,"y1":0,"x2":0,"y2":5,"ccw":true}
      {"type":"polygon","cx":0,"cy":0,"x":5,"y":0,"sides":6}   (x,y = a vertex)
    Returns the sketch's feature name for use with sw_extrude / sw_revolve."""
    m = active_model("part")
    m.ClearSelection2(True)
    if not select(m, plane, "PLANE"):
        return _fail(f"Could not select plane '{plane}'. Check the exact name in sw_feature_tree.")
    sm = m.SketchManager
    sm.InsertSketch(True)
    made = 0
    for e in entities:
        t = e["type"]
        if t == "line":
            r = sm.CreateLine(e["x1"] * MM, e["y1"] * MM, 0, e["x2"] * MM, e["y2"] * MM, 0)
        elif t == "centerline":
            r = sm.CreateCenterLine(e["x1"] * MM, e["y1"] * MM, 0, e["x2"] * MM, e["y2"] * MM, 0)
        elif t == "rect":
            r = sm.CreateCornerRectangle(e["x1"] * MM, e["y1"] * MM, 0, e["x2"] * MM, e["y2"] * MM, 0)
        elif t == "circle":
            r = sm.CreateCircleByRadius(e["x"] * MM, e["y"] * MM, 0, e["r"] * MM)
        elif t == "arc":
            r = sm.CreateArc(e["cx"] * MM, e["cy"] * MM, 0, e["x1"] * MM, e["y1"] * MM, 0,
                             e["x2"] * MM, e["y2"] * MM, 0, 1 if e.get("ccw", True) else -1)
        elif t == "polygon":
            r = sm.CreatePolygon(e["cx"] * MM, e["cy"] * MM, 0, e["x"] * MM, e["y"] * MM, 0,
                                 int(e["sides"]), True)
        else:
            sm.InsertSketch(True)
            return _fail(f"Unknown entity type {t}")
        made += 1 if r is not None else 0
    sm.InsertSketch(True)  # exit
    feat = _last_feature(m, "ProfileFeature")
    if feat is not None and name:
        feat.Name = name
    return _ok(sketch=feat.Name if feat else None, entities_created=made, entities_requested=len(entities))


def _select_sketch(m, sketch: str):
    m.ClearSelection2(True)
    if not select(m, sketch, "SKETCH"):
        raise RuntimeError(f"Could not select sketch '{sketch}'.")


@tool
def sw_extrude(sketch: str, depth: float, cut: bool = False, end: str = "blind",
               both_directions: bool = False, depth2: float = 0, flip: bool = False,
               draft_deg: float = 0, merge: bool = True) -> str:
    """Extrude a sketch as boss (cut=False) or cut (cut=True). depth in mm.
    end: blind | through_all | up_to_next | mid_plane."""
    import math

    m = active_model("part")
    _select_sketch(m, sketch)
    fm = m.FeatureManager
    t1 = END_COND[end]
    n = _null()
    if cut:
        f = fm.FeatureCut4(True, flip, False, t1, 0, depth * MM, depth2 * MM, False, False, False, False,
                           math.radians(draft_deg), 0, False, False, False, False, True, True, True,
                           True, False, False, 0, 0, False, False)
    else:
        f = fm.FeatureExtrusion3(True, flip, both_directions, t1, 0 if not both_directions else t1,
                                 depth * MM, depth2 * MM, False, False, False, False,
                                 math.radians(draft_deg), 0, False, False, False, False, merge, True, True,
                                 0, 0, False)
    m.ClearSelection2(True)
    return _ok(feature=f.Name) if f else _fail("Extrude failed (check sketch is closed and not self-intersecting).")


@tool
def sw_revolve(sketch: str, angle_deg: float = 360, cut: bool = False) -> str:
    """Revolve a sketch about a centerline contained in that same sketch (add one with type 'centerline')."""
    import math

    m = active_model("part")
    _select_sketch(m, sketch)
    # Select the sketch's centerline as the axis (mark 4).
    feat = [f for f in features(m) if f.Name == sketch][0]
    sk = feat.GetSpecificFeature2()
    axis = None
    for seg in (sk.GetSketchSegments() or []):
        if seg.ConstructionGeometry and seg.GetType() == 0:
            axis = seg
    if axis is None:
        return _fail("Sketch has no centerline to use as the axis.")
    data = m.SelectionManager.CreateSelectData()
    data.Mark = 4
    axis.Select4(True, data)
    f = m.FeatureManager.FeatureRevolve2(True, True, False, cut, False, False, 0, 0, math.radians(angle_deg), 0,
                                         0, 0, 0, 0, 0, 0, 0, True, True, True)
    m.ClearSelection2(True)
    return _ok(feature=f.Name) if f else _fail("Revolve failed.")


@tool
def sw_fillet(radius: float, edges: list[dict]) -> str:
    """Constant-radius fillet. edges: list of points (mm, model space) lying on each edge: [{"x":10,"y":5,"z":0}, ...]."""
    m = active_model("part")
    m.ClearSelection2(True)
    for i, p in enumerate(edges):
        if not select(m, "", "EDGE", p["x"], p["y"], p["z"], append=i > 0):
            return _fail(f"Could not select an edge at {p}")
    n = _null()
    f = m.FeatureManager.FeatureFillet3(195, radius * MM, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, n, n, n, n, n, n, n)
    m.ClearSelection2(True)
    return _ok(feature=f.Name) if f else _fail("Fillet failed.")


@tool
def sw_chamfer(distance: float, edges: list[dict], angle_deg: float = 45) -> str:
    """Angle-distance chamfer on the edges given by points on them (mm, model space)."""
    import math

    m = active_model("part")
    m.ClearSelection2(True)
    for i, p in enumerate(edges):
        if not select(m, "", "EDGE", p["x"], p["y"], p["z"], append=i > 0):
            return _fail(f"Could not select an edge at {p}")
    f = m.FeatureManager.InsertFeatureChamfer(4, 1, distance * MM, math.radians(angle_deg), 0, 0, 0, 0)
    m.ClearSelection2(True)
    return _ok(feature=f.Name) if f else _fail("Chamfer failed.")


@tool
def sw_shell(thickness: float, faces_to_remove: list[dict], outward: bool = False) -> str:
    """Shell the part, removing the faces that contain the given points (mm, model space)."""
    m = active_model("part")
    m.ClearSelection2(True)
    for i, p in enumerate(faces_to_remove):
        if not select(m, "", "FACE", p["x"], p["y"], p["z"], append=i > 0):
            return _fail(f"Could not select a face at {p}")
    m.InsertFeatureShell(thickness * MM, outward)
    m.ClearSelection2(True)
    return _ok()


@tool
def sw_linear_pattern(feature: str, count: int, spacing: float, direction_edge: dict) -> str:
    """Linear pattern of a feature along an edge (point on edge given in mm). Direction 1 only."""
    m = active_model("part")
    m.ClearSelection2(True)
    select(m, "", "EDGE", direction_edge["x"], direction_edge["y"], direction_edge["z"], mark=1)
    select(m, feature, "BODYFEATURE", append=True, mark=4)
    n = _null()
    f = m.FeatureManager.FeatureLinearPattern4(count, spacing * MM, 1, 0, False, False, "NULL", "NULL",
                                               False, False, False, False, False, False, True, True, False,
                                               False, 0, 0, False)
    m.ClearSelection2(True)
    return _ok(feature=f.Name) if f else _fail("Pattern failed.")


@tool
def sw_suppress(feature: str, suppress: bool = True) -> str:
    """Suppress or unsuppress a feature by name."""
    m = active_model()
    m.ClearSelection2(True)
    if not select(m, feature, "BODYFEATURE"):
        return _fail(f"Feature '{feature}' not found.")
    r = m.EditSuppress2() if suppress else m.EditUnsuppress2()
    m.ClearSelection2(True)
    return _ok(result=bool(r))


@tool
def sw_delete_feature(feature: str) -> str:
    """Delete a feature (or sketch) by name."""
    m = active_model()
    m.ClearSelection2(True)
    if not (select(m, feature, "BODYFEATURE") or select(m, feature, "SKETCH")):
        return _fail(f"'{feature}' not found.")
    return _ok(deleted=bool(m.Extension.DeleteSelection2(0)))


@tool
def sw_set_material(material: str, database: str = "SOLIDWORKS Materials") -> str:
    """Assign a material, e.g. material='1060 Alloy' (database name is the .sldmat library, default SOLIDWORKS Materials)."""
    m = active_model("part")
    m.SetMaterialPropertyName2("", database, material)
    return _ok(material=material)


# --------------------------------------------------------------------------- assemblies
@tool
def sw_add_component(path: str, x: float = 0, y: float = 0, z: float = 0) -> str:
    """Insert a part/sub-assembly into the active assembly at (x,y,z) mm. Mates: use sw_exec_python (AddMate5)."""
    m = active_model("assembly")
    c = m.AddComponent5(path, 0, "", False, "", x * MM, y * MM, z * MM)
    return _ok(component=c.Name2) if c else _fail("AddComponent5 failed.")


# --------------------------------------------------------------------------- Simulation (add-in)
def _sim_mgr():
    sw = app()
    cw = sw.GetAddInObject("SldWorks.Simulation")
    if cw is None:
        raise RuntimeError("SOLIDWORKS Simulation add-in is not loaded. Enable it in Tools > Add-Ins.")
    return cw.ActiveDoc.StudyManager


@tool
def sw_sim_list_studies() -> str:
    """List Simulation studies in the active document (requires the Simulation add-in)."""
    sm = _sim_mgr()
    return _ok(studies=[sm.GetStudy(i).Name for i in range(sm.StudyCount)])


@tool
def sw_sim_create_study(name: str, study_type: int = 0) -> str:
    """Create a Simulation study. study_type: 0 static, 1 frequency, 2 buckling, 3 thermal, 4 fatigue... (swsAnalysisStudyType_e)."""
    sm = _sim_mgr()
    err = _byref_int()
    s = sm.CreateNewStudy3(name, study_type, 0, err)
    return _ok(study=name) if s else _fail(f"CreateNewStudy3 failed, code {err.value}")


@tool
def sw_sim_mesh_and_run(study: str) -> str:
    """Mesh (default settings) and run the named study. Loads/fixtures must already exist (set up via sw_exec_python or the GUI)."""
    sm = _sim_mgr()
    for i in range(sm.StudyCount):
        s = sm.GetStudy(i)
        if s.Name == study:
            merr = s.CreateMesh(0, 0, 0)
            rerr = s.RunAnalysis()
            return _ok(mesh_error=merr, run_error=rerr)
    return _fail(f"Study '{study}' not found.")


# --------------------------------------------------------------------------- escape hatches
@tool
def sw_exec_python(code: str) -> str:
    """Run arbitrary Python against the live SolidWorks COM API for anything the dedicated tools don't cover
    (mates, drawings, Simulation loads/fixtures, Flow Simulation, ...). In scope: swApp, model (active doc or None),
    MM (0.001), NULL (null COM dispatch for optional object args), byref_int(), win32com, pythoncom.
    Assign to `result` to return a value; print() output is captured. Lengths in the SolidWorks API are metres."""
    sw = app()
    g = {"swApp": sw, "model": sw.ActiveDoc, "MM": MM, "NULL": _null(), "byref_int": _byref_int,
         "result": None}
    pythoncom, wc = _com()
    g.update(win32com=wc, pythoncom=pythoncom)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        exec(code, g)
    return _ok(stdout=buf.getvalue(), result=g["result"])


@tool
def sw_run_macro(path: str, module: str, procedure: str = "main") -> str:
    """Run a recorded .swp macro: path, module name, procedure name."""
    err = _byref_int()
    r = app().RunMacro2(path, module, procedure, 1, err)
    return _ok(ran=bool(r), error_code=err.value)


if __name__ == "__main__":
    mcp.run()
