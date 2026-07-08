# Phase 0 — Spike notes

The spike asks two questions that decide whether the whole architecture holds
(see `CLAUDE.md` → "PHASE 0"). Both ultimately need **real PowerPoint (Windows
or Mac, current M365)** because they are about how PowerPoint *renders* an
inserted deck and how it manages masters on `insertSlidesFromBase64`.

This file records what was verified automatically (the server-side half — the
payload and the renderer that produces it) and the exact manual steps + pass
criteria for the PowerPoint half.

## A. Server-side verification — DONE (automated, reproducible on Linux)

| Check | Result |
|---|---|
| Runtime version pin | **v2.6.18** (matches `CLAUDE.md`) |
| Runtime test suite (`cd renderer && pytest -q`) | **174 passed, 1 skipped** |
| Golden brief `board_recommendation.json` via `/render` | 4 slides, gate **PASS**, score **96** |
| Golden brief `data_review.json` via `/render` | 3 slides, gate **PASS**, score **98** |
| Golden brief `narrative_briefing.json` via `/render` | 4 slides, gate **PASS**, score **92** |
| `spike_deck.b64` decodes and equals `spike_deck.pptx` (byte-identical) | **OK** (813,831 bytes) |
| Spike deck slide count / size | **5 slides**, 16:9 (12192000 × 6858000 EMU) |
| Template `ow_default.pptx` | 1 master, **18 layouts** (matches `CLAUDE.md`) |

The spike deck payload provably **contains** every element Q1 asks about
(inspected in the OOXML — this is the server-side guarantee; whether PowerPoint
draws them correctly on insert is exactly what the manual step confirms):

- **Treemap** → `ppt/charts/chartEx1.xml` (chartEx) + EMF fallback (`image1.emf`)
- **Waterfall** → `ppt/charts/chart2.xml` (barChart) with data-label number
  format **`[$-407]#,##0`** → the de-DE `1.000` styling
- **Icon freeforms** → 8 × `custGeom` + an SVG icon in media
- **Native tables** → 3 × `<a:tbl>`
- 17 `graphicFrame` anchors total

### Reproduce

```bash
cd renderer
pip install -r requirements.txt -r runtime_requirements.txt
pytest -q                                  # 174 passed, 1 skipped
uvicorn main:app --port 8123 &             # then POST testdata/*.json to /render
curl localhost:8123/healthz               # {"ok":true,"runtime_version":"2.6.18"}
```

## B. PowerPoint verification — PASSED ✅ (2026-07-08, PowerPoint on the web, M365)

Verified in real PowerPoint. **Verdict: green light for the architecture.**

- **Q1 — Fidelity: PASS.** All 5 slides render intact after
  `insertSlidesFromBase64` — treemap, icon freeforms, native table, waterfall
  (bars + connectors), and de-DE `1.000`-style numbers all correct.
- **Q2 — Master hygiene: PASS.** After 3× insert there is still only **one**
  master/layout set — no duplication. `useDestinationTheme` maps the inserted
  layouts cleanly onto the open presentation's master.

Sideload path that worked on a locked-down (non-admin) corporate Mac: the
self-contained pane hosted on **GitHub Pages** (`docs/index.html`,
`https://<user>.github.io/<repo>/`), manifest `manifest.hosted.xml` uploaded
via **PowerPoint on the web → Insert → Add-ins → Upload My Add-in**. No
localhost, no dev certificate, no admin rights. The localhost dev flow
(`npm run sideload`) is a non-starter without admin to trust the dev cert.

The steps below are retained for reference / re-running the spike.

### Reproduce (steps)

This half cannot run in the headless build environment. Steps:

```bash
cd addin
npm install
npm run sideload         # office-addin-debugging; office-addin-dev-certs
                         # provisions the https://localhost:3000 cert
```

If `npm run sideload` misbehaves, sideload `manifest.xml` manually and serve
`src/` over **https://localhost:3000** (Office requires https even for
localhost). `SourceLocation` is `https://localhost:3000/src/taskpane.html`.

> Note: `npm run validate` calls Microsoft's online validator and will fail
> with `403` behind a restricted network — that is the network policy, not a
> manifest defect. The manifest is well-formed and has all required elements
> (checked locally).

### Q1 — Fidelity

Click **"Insert test deck at end"**, then visually inspect all 5 slides:

- [ ] Treemap renders (not a broken/empty box)
- [ ] Icons crisp (freeforms, not rasterized/missing)
- [ ] Native table styled correctly
- [ ] Waterfall: bars **and** connector lines present
- [ ] Numbers show `1.000`-style (de-DE), e.g. thousands with `.`

If any element fails → **STOP**, note exactly what broke visually, reassess
(fallbacks in `CLAUDE.md`: image-based export for broken elements, or drop the
affected intent from the pane).

### Q2 — Master hygiene

Click **"Insert test deck"** **3×**, then **View → Slide Master**:

- [ ] Count masters/layouts. Does each insert **duplicate** the master set?

If inserts duplicate masters → we need a mitigation (start-from-template
discipline / post-insert cleanup / `useDestinationTheme` tuning) **before**
building further. Starting your presentation from `renderer/ow_default.pptx`
is the recommended baseline so `useDestinationTheme` maps layouts cleanly.

### Report back

Paste the Q1 checklist result + the Q2 master count. That verdict decides
whether we proceed to Phase 1 (MVP generation) as planned.
