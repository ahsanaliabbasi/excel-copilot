# Excel Copilot — UI Redesign & Upload Fixes (Session Context)

Companion to [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md). That file describes the product and engine; this one records the **visual redesign** and the **run / upload troubleshooting** work done afterwards: what was asked, what changed, what was decided, and what is still open.

---

## 1. Request

1. Re-skin the whole app with the branding, colours, styling and premium feel of **Insider One** (insiderone.com), using nine reference screenshots (dark hero, cream sections, frosted cards, orange gradient pill buttons, charcoal nav).
2. Add a proper **light mode** that follows the light-section look: gradient cards, soft effects — then refine it to the cleaner warm-neutral palette of the "Loved by brands" section so text stays clearly visible.
3. Fix problems running the project and uploading a workbook (Upload button, drag-and-drop).
4. Keep the disabled **Download Edited File** button readable in both themes.

Only colours, type and styling were taken from the references. No Insider One logo, photos or copy were used.

## 2. Design system

### Typography
Red Hat Display (Google Fonts, weights 300–700) replaces Inter. Headings are light weight (400–500) and large; brand name and section labels use wide letter-spacing and uppercase.

### Tokens (`:root` in `frontend/style.css`)
| Token | Light | Dark |
|---|---|---|
| Canvas `--bg-app` | `#f3f0ec` (warm grey-cream) | `#121214` |
| Surface `--bg-surface` | `#fbf9f7` | `#1c1c1f` |
| Sidebar `--bg-sidebar` | `#f7f4f0` (light override) | `#0e0e10` |
| Text `--text-primary` / secondary | `#1f1d23` / `#4f4c55` | `#f5f2ef` / `#b1acb0` |
| Border | `#e6e0da` | `#2e2e33` |
| Accent `--primary` | `#ee4b2b` | `#ff5a36` |
| Accent gradient `--accent-grad` | `#ff6a3d → #ee3b1f` (135°) | same |
| Radius | 24px cards, 28px modals, pill (999px) buttons/inputs | same |

Semantic colours (success green, danger red) were re-tuned to sit with the orange accent. Formula text is green, formula cells ember.

### Components
- **Top bar** — charcoal (dark) / cream (light), uppercase spaced brand, orange gradient logo mark, white "Upload" pill, orange gradient "Download" pill, outlined theme toggle.
- **Buttons** — pill-shaped; primary and success use the orange gradient with a soft orange glow.
- **Upload screen (hero)** — dark: charcoal with red/orange radial glows and a dark curved "ridge" at the bottom, big light-weight headline with "clicking" in orange, frosted-glass drop zone. Light: cream background, dark text, white drop zone.
- **Sidebar** — charcoal with a faint orange glow at the bottom (dark); plain off-white panel (light). Rounded items, orange icons, active sheet in the orange gradient.
- **Cards / modals / wizard cards** — very rounded, thin border, soft shadow. Dark: charcoal surfaces. Light: near-white gradient (`#fdfcfb → #f8f1ec`).
- **Chips / toolbar buttons** — white pills with a faint shadow in light mode (like the "3.96/5.0" score pills in the reference).
- **Grid** — new default row-stripe theme **"Ember & cream"** (`data-band="ember"`, default in `grid.js`); the other stripe colours remain in the dropdown. Hover, selection and duplicate highlights use the accent.
- **Toasts** — dark pill with an orange edge.
- **Scrollbars** — thin, dark in the sidebar, accent-tinted in the grid and modals.

### Theme switching
`data-theme="dark"` on `<html>` (persisted in `localStorage` as `theme`). **No attribute = light.** Light-mode overrides are written as `:root:not([data-theme='dark']) …`.

## 3. How the CSS is organised (important before editing)

`style.css` was rebuilt in place, then extended with **appended override blocks** as the light theme was iterated:

1. Tokens + chrome (top bar, buttons, hero, sidebar, cards) — rewritten near the top.
2. Component rules from the original file — updated in place (indigo → ember, Inter → Red Hat Display).
3. Appended blocks, in order: scrollbars → **Light mode (round 1: frosted glass + mountain glow)** → **round 2 (light hero, blurred SVG mountain)** → **round 3 (clean warm-neutral palette; mountain glow disabled with `body::before { display: none }`)** → drop-zone click target → dark and light "disabled Download" rules.

Later blocks intentionally win over earlier ones. **Round 3 is the current light look.** Rounds 1–2 rules are still in the file (the `body::before` mountain and the glass gradients are overridden, not deleted). A future clean-up should merge these into a single light-mode section.

The last two rules in the file:
```css
[data-theme='dark'] .btn.success:disabled { opacity: 1; color: #d9d3ce; background: #3a3a41; }
:root:not([data-theme='dark']) .btn.success:disabled { opacity: 1; color: #5b555d; background: #e4dcd4; }
```

## 4. Files changed

| File | Change |
|---|---|
| `frontend/style.css` | New tokens, chrome, hero, light-mode layers, drop-zone click target, disabled-button rules |
| `frontend/index.html` | Red Hat Display font link; "Ember & cream" stripe option; hero headline with `<em>clicking</em>`; trust line under the drop zone; rounder logo mark |
| `frontend/grid.js` | Default stripe theme `ember` |
| `frontend/app.js` | Upload picker helper, drop-zone click, robust drag-and-drop, `loadWorkbook()` shared by file handling (see §5) |
| `backend/main.py` | Upload set-up moved into `_register_upload()` (behaviour unchanged) |
| `backend/requirements.txt` | Pins loosened to minimum versions (see §6) |
| `start.bat` (repo root) | Starts the server with the working venv and opens the default browser |

`FRONTEND_VERSION` (17) and `API_VERSION` (6) were **not** bumped — the server serves the frontend with `no-store`, so Ctrl+F5 is enough.

## 5. Upload and drag-and-drop behaviour (final state)

- **Any of these opens the file picker:** top-bar "Upload Excel File", "Browse for a file", the big upload icon, or anywhere in the drop zone (icon, text, empty space; also Enter/Space when focused). All go through `openFilePicker()`, which clears the hidden `#fileInput` value and clicks it (the original mechanism; picking the same file twice still works).
- **Drag and drop** works anywhere on the page via `window` listeners. A dragenter/leave counter stops the drop zone flickering; dropping prevents the browser from opening the file itself. Files are read from `dataTransfer.files`, falling back to `dataTransfer.items`. A non-`.xlsx/.xlsm` file shows "Please drop a .xlsx file". A drop that carries no file shows a hint to drag from File Explorer, not from inside the editor.
- **Load flow:** `handleFile()` → `loadWorkbook(info, url, init, name)` (upload → show app → select first sheet → toast → background sheet polling). A 404/405 from an outdated server shows "the server is out of date — stop it and start it again".

### Tried and removed
- A **"paste the file path → Open"** box with a `/api/open-path` endpoint. It worked, but was removed at the user's request; both the UI and the endpoint are gone.
- A per-click freshly-created hidden `<input type=file>` and an icon-as-`<button>`. Both worked in testing but were reverted in favour of the original hidden-input mechanism.

## 6. Running the project — problems found

| Symptom | Cause | Fix |
|---|---|---|
| `pip install -r requirements.txt` fails building **pandas** (`metadata-generation-failed`) | Pinned `pandas==2.2.3` (and other `==` pins) have no wheels for **Python 3.14**, so pip tries to compile them | `requirements.txt` now uses `>=` minimums; the working venv runs pandas 3.0.6, fastapi 0.141, uvicorn 0.53, numpy 2.5, sqlglot 30 |
| `uvicorn` / `fastapi` "not recognized / not installed" | Two virtual environments exist. The **outer** `vv/excel-copilot/venv` is empty (and was created at a different path). The working one is `backend/venv` | Run with the backend venv's Python (below) and select `backend\venv` as the VS Code interpreter |
| `Could not import module "main"` | uvicorn started from the repo root | Start it from `backend/` |

**Run it:**
```powershell
cd "C:\Users\aa255214\Downloads\DataEngineering\MS Excel\excel-copilot\vv\excel-copilot\backend"
.\venv\Scripts\python.exe -m uvicorn main:app --reload --port 8000
```
Open http://localhost:8000 and press Ctrl+F5. Or double-click `start.bat`.

The backend serves the frontend, so there is no separate frontend server. For a frontend-only static server (`python -m http.server 5500` in `frontend/`), uploads and operations still need the backend on port 8000.

### Upload not opening a file dialog
The click handlers are verified working (§7). The likely cause on the user's machine is an **embedded browser (the VS Code Simple Browser) that blocks the OS file dialog and does not receive files dragged from Explorer**. This was inferred from screenshots, not confirmed. Use Edge or Chrome directly (`start.bat` does this). If it still fails in a normal browser, gather: is anything shown (toast, dialog, nothing)?

## 7. How it was tested

- Headless Edge screenshots of the upload screen and a workspace view (with injected fake grid rows), in both themes.
- Chrome DevTools Protocol driven from Python (`websocket-client`): real mouse events on each upload target with `HTMLInputElement.click` counted; `DOM.setFileInputFiles` to simulate choosing a workbook; `Input.dispatchDragEvent` with a real file to simulate an OS drop (on the drop zone, on empty page area, and with a non-xlsx file).
- Results: every upload target opened the picker exactly once in light and dark; the sample workbook (`Data Exercise - Ahsan Ali.xlsx`, 5 sheets) loaded via picker and via drop; a `.ini` drop was rejected.
- Temporary servers ran on port 8001 and were stopped afterwards. Test scripts were kept in the session scratchpad, not the repo.

**Not verified visually:** dark mode after the light-mode rounds, and modals, the wizard, SQL and Duplicates screens in either theme (they use the shared tokens, so they should follow). The real Windows file dialog cannot be shown in headless Edge.

## 8. Decisions worth remembering

- The dark hero stays for dark mode only; light mode has a fully light hero (matching "change the theme entirely to light").
- Light mode first used frosted blur + a mountain-shaped glow, then was simplified (round 3) because text visibility mattered more than the effect. The mountain SVG glow is still in the CSS but hidden.
- The disabled Download button must remain readable in both themes (explicit rules; opacity forced to 1).
- Any change that needs the server to restart (Python) should be called out to the user; frontend-only changes need just Ctrl+F5.

## 9. Open items / next steps

- Consolidate the appended light-mode override blocks into one clean section in `style.css`.
- Visually review modals, wizard, SQL and Duplicates screens in both themes; check the "Ember & cream" stripes against dark mode.
- Delete or repair the empty outer `venv/` folder so it stops being picked up by terminals and the IDE.
- Consider a `README` note about Python 3.14 requirements and `start.bat`.
- Optionally bump `FRONTEND_VERSION` (and the four `?v=17` query strings) if a stale-cache banner is wanted for this redesign.
- Earlier suggestions from `PROJECT_CONTEXT.md` still stand (wizard for remaining tools, persisted sessions, pivot tables/charts, undo for row deletion, auth before sharing).
