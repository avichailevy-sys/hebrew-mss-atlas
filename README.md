# Hebrew Manuscripts Atlas — 14th Century Demo

## Files

- **`index.html`** — the Leaflet map. Single HTML file with embedded CSS/JS.
- **`atlas_data.json`** — pre-computed data layer (604 manuscripts, 218 places).
- **`build_atlas.py`** — Python script that produced `atlas_data.json` from `catalogue_14c.xlsx`. Keep for future re-runs / extending the geocoding table.

## Deploy

Drop `index.html` and `atlas_data.json` together into a folder served over HTTP. For GitHub Pages, push them to the repo root and enable Pages on the `main` branch. They sit side-by-side in the same directory; `index.html` fetches `atlas_data.json` at runtime.

Won't work via `file://` because of browser CORS rules on local fetch. Test locally with:

    python -m http.server 8000

then open `http://localhost:8000/`.

## Coverage

- 604 of 630 14th-century records rendered (95.9%)
- 26 excluded — region-only attribution like "Spain", "Italy, Northern", "Yemen (Republic)" that can't be placed at a point
- 0 unmatched cities (full geocoding coverage of the city-level subset)

## Re-running the build

If you fix or extend the geocoding lookup in `build_atlas.py`, regenerate with:

    python build_atlas.py

It expects `catalogue_14c.xlsx` in the working directory and writes `atlas_data.json` next to itself.

## Footer note for the demo

The interface acknowledges the simplifications: it shows the rendered/excluded counts, links to the NLI Ktiv catalog as the data source, and names the underlying MARC fields (008 for date, 751a for place of writing). That's the minimum methodological honesty for a research-oriented audience while keeping the visual clean.
