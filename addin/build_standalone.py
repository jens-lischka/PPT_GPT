#!/usr/bin/env python3
"""Build a self-contained spike task pane.

Inlines the pre-rendered spike deck (base64) and the golden brief straight
into a single HTML file so the pane can be hosted anywhere (or opened from
any static host) with NO sibling asset fetches, NO CORS, and NO relative
paths. Output: dist/taskpane.html.

The only external dependency left is office.js from Microsoft's CDN, which
Office itself requires and always allows.

Writes two copies of the same file:
  - addin/dist/taskpane.html  (generic hostable artifact)
  - docs/index.html           (served by GitHub Pages at the repo root URL)

    python3 build_standalone.py
"""
from __future__ import annotations

import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "src" / "assets"
DIST = HERE / "dist"
DOCS = HERE.parent / "docs"

SPIKE_B64 = (ASSETS / "spike_deck.b64").read_text().strip()
BRIEF = json.loads((ASSETS / "board_recommendation.json").read_text())

HTML = """<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>OW Deck Builder</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <style>
    body {{ font-family: "Segoe UI", sans-serif; margin: 12px; font-size: 13px; }}
    #log {{ white-space: pre-wrap; background: #f5f4f0; padding: 8px;
           border-radius: 4px; min-height: 120px; margin-top: 8px; }}
    button {{ margin: 4px 4px 4px 0; padding: 6px 10px; }}
    .gate-pass {{ color: #2a7a2a; }} .gate-fail {{ color: #b02020; }}
  </style>
</head>
<body>
  <h3>OW Deck Builder — Phase 0 spike</h3>

  <p><b>Q1 (fidelity):</b> does the full feature set (chartEx treemap,
  icon freeforms, native tables, waterfall, de-DE numbers) survive
  <code>insertSlidesFromBase64</code> visually intact?</p>
  <button id="btnSpike">Insert test deck at end</button>

  <p><b>Q2 (master hygiene):</b> click the button 3× and inspect
  View &rarr; Slide Master for duplicated masters/layouts.</p>

  <hr/>
  <button id="btnSelection">Log selected slide ids + tags</button>

  <div id="log">ready.</div>

  <script>
    // ---- Deck payload inlined at build time (self-contained; no fetch). ----
    const SPIKE_B64 = "{spike_b64}";

    const log = (m) => {{
      document.getElementById("log").textContent += "\\n" + m;
    }};

    Office.onReady(() => {{
      const host = Office.context.host;
      const set15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
      log("Office.js ready — host: " + host + " | PowerPointApi 1.5: " + set15);
    }});

    // ---- Core: insert the base64 pptx, mapping layouts onto the open master.
    async function insertBase64(pptxBase64) {{
      await PowerPoint.run(async (ctx) => {{
        ctx.presentation.insertSlidesFromBase64(pptxBase64, {{
          formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
        }});
        await ctx.sync();
      }});
    }}

    document.getElementById("btnSpike").onclick = async () => {{
      try {{
        await insertBase64(SPIKE_B64);
        log("inserted spike deck — VISUALLY inspect: treemap? icons? table? waterfall? de-DE numbers?");
      }} catch (e) {{ log("ERROR " + e); }}
    }};

    document.getElementById("btnSelection").onclick = async () => {{
      try {{
        await PowerPoint.run(async (ctx) => {{
          const sel = ctx.presentation.getSelectedSlides();
          sel.load("items/id,items/index,items/tags/key,items/tags/value");
          await ctx.sync();
          for (const s of sel.items) {{
            const tags = s.tags.items.map(t => t.key + "=" + t.value).join(", ") || "(no tags)";
            log("slide index=" + s.index + " id=" + s.id + " tags: " + tags);
          }}
          if (!sel.items.length) log("(nothing selected)");
        }});
      }} catch (e) {{ log("ERROR " + e); }}
    }};
  </script>
</body>
</html>
"""

rendered = HTML.format(spike_b64=SPIKE_B64)

DIST.mkdir(exist_ok=True)
DOCS.mkdir(exist_ok=True)
for out in (DIST / "taskpane.html", DOCS / "index.html"):
    out.write_text(rendered)
    print(f"wrote {out} ({out.stat().st_size} bytes)")

# The Phase 1+2 app pane is hand-maintained (no inlined assets); copy it to
# docs/ so GitHub Pages serves it alongside the spike pane.
app = (HERE / "src" / "app.html").read_text()
(DOCS / "app.html").write_text(app)
print(f"wrote {DOCS / 'app.html'} ({len(app)} bytes)")

# .nojekyll stops GitHub Pages' Jekyll pass from touching the static files.
(DOCS / ".nojekyll").write_text("")
print(f"wrote {DOCS / '.nojekyll'}")
