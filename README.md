# Efes Nexus

Business case for **Efes Nexus** — a promo management platform for Efes Kazakhstan.
This repository holds the client-facing deliverables (presentation deck, commercial
proposal, interactive system demo) together with everything needed to rebuild them.

## Layout

| Path | What it holds |
| --- | --- |
| `deck/` | The presentation. `Efes_Nexus_RU.pptx` and `Efes_Nexus_EN.pptx` are the deliverables. |
| `deck/base/` | Pristine generator output — the input to the transform step. Do not edit by hand. |
| `deck/build/` | Everything that produces the deck: `build.js`, `transform.py`, icon generator, render script. |
| `proposal/` | Commercial proposal (КП) in DOCX + PDF, and the generator that produces it. |
| `system/` | Self-contained interactive demo of the platform (`index.html`), plus a dev server. |
| `assets/` | Source media: logos, icons, product screenshots, screen recording. |
| `data/` | Source data: BPMN process models and the financial / promo-status spreadsheets. |

## Rebuilding the deck

The deck is produced in two stages.

**Stage 1 — base deck (`deck/build/build.js`).** A pptxgenjs generator that lays out
the slides from scratch and writes `deck/base/Efes_Nexus_RU_base.pptx`. This stage is
slow-moving; it only needs re-running when the underlying slide set changes.

```sh
cd deck/build && npm install && node build.js
```

The English base deck is produced from the Russian one by the string-replacement
pipeline in `deck/build/i18n/` (`extract.py` pulls the strings, `translations.json`
holds the mapping, `apply.py` writes them back).

**Stage 2 — transform (`deck/build/transform.py`).** Takes the two base decks and
applies every content and design change agreed with the client: duplicating and
retitling the BPMN slides, building the KZPromotion, statistics, dashboard and
future-development slides, embedding the demo video and swapping in the current
logo. This is the stage that is actually iterated on.

```sh
pip install python-pptx
python3 deck/build/transform.py     # writes deck/Efes_Nexus_{RU,EN}.pptx
```

To review the result visually:

```sh
deck/build/render.sh                # PDF + per-slide PNGs into deck/render/
```

Icons are pre-rendered into `assets/icons/`; regenerate them only when a new one is
needed (`cd deck/build/icons && npm install && node gen.js`).

## Rebuilding the proposal

```sh
cd proposal/build && npm install && node generate.js
```

Writes the DOCX into `proposal/`. Export to PDF with LibreOffice:
`soffice --headless --convert-to pdf --outdir proposal proposal/*.docx`.

## The system demo

`system/index.html` is a single self-contained file — open it directly in a browser.
`system/index.dev.html` is the unbundled version used while developing. To serve the
folder over HTTP (needed for a few browser APIs):

```sh
node system/dev/server.js     # http://localhost:8123
```

## Conventions

- Generated output (`deck/render/`, `proposal/render/`, `node_modules/`, Python
  caches) is not tracked — see `.gitignore`.
- Deliverables *are* tracked, so anyone cloning the repo gets the current decks and
  proposal without running a build.
- Brand palette and typography live at the top of `deck/build/build.js` and are
  mirrored in `transform.py`; change both together.
