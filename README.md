# Claude x SolidWorks

Lets Claude drive SolidWorks (new models, edits, exports, Simulation studies, anything reachable through the
COM API) from **Claude Code**, using your own Claude login and usage allowance.

## Why Claude Code is the front end
You asked for history, model selection, internet access and image uploads. Claude Code already has all of them
(`/resume` for history, `/model`, WebSearch/WebFetch, paste or drag images), and it signs in with your Claude
subscription. A custom chat app that reused your subscription login would not be permitted by Anthropic's terms
(third-party apps must use API keys), so the SolidWorks-specific part is an MCP server that plugs into Claude Code.
Check Anthropic's current terms if in doubt.

## Setup (on the Windows PC that has SolidWorks)
1. Install Python 3.10+ and Claude Code, then `claude login` with your Claude account.
2. `git clone` this repo, `cd solidworks-claude`, `pip install -r requirements.txt`
3. Start SolidWorks (any licence; Simulation/Flow need those add-ins enabled in Tools > Add-Ins).
4. In this folder run `claude`. It picks up `.mcp.json` and `CLAUDE.md`; approve the `solidworks` server.
5. Ask for things, e.g. "Make a 60x40x20 mm block with a 10 mm through hole and 2 mm fillets, then show me a screenshot".

Tips: attach reference sketches/photos by pasting an image; `/model` to switch models; `claude --resume` for history.
Don't pre-approve `sw_exec_python` - it runs arbitrary code on your machine.

## Tools (32)
Documents (new/open/save/close/export/screenshot), inspection (feature tree, mass properties, dimensions,
custom properties, equations), modelling (sketch, extrude/cut, revolve, fillet, chamfer, shell, linear pattern,
suppress, delete, material), assembly (add component), Simulation (list/create/mesh+run study), and two escape
hatches: `sw_exec_python` (full COM access) and `sw_run_macro`.

## Honest status
- Written against the SolidWorks API from documentation, **not yet run against a real SolidWorks install**
  (the build environment is Linux). The server loads and registers all tools; COM calls are untested. Several calls
  (FeatureCut4, FeatureFillet3, FeatureLinearPattern4, FeatureRevolve2, Simulation) have long positional signatures
  that vary by SolidWorks version; if one errors, Claude reports it and can retry via `sw_exec_python`.
  Please report failures so the signatures can be fixed.
- SolidWorks Flow Simulation has almost no public automation API; expect partial/manual workflows there.
- Complex mates, drawings and Simulation loads/fixtures go through `sw_exec_python` rather than dedicated tools.
