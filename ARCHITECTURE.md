# koulutus Architecture

## Boundaries

- Static HTML files are the source material for browser-facing training pages.
- `generate_*_pptx.js` scripts generate PowerPoint decks with `pptxgenjs`.
- `generate_pdfs.js` uses browser rendering to produce PDF artifacts.
- Generated `.pptx` and `.pdf` files are committed deliverables, not a hidden
  build cache.

## Verification

- JavaScript generator syntax: `node --check`.
- Required source and generated artifacts: checked by `scripts/check-repo.sh`.
- Visual quality of generated decks/PDFs still needs manual or render-based
  review when content changes.

## Ownership Rule

Training content decisions should live in this repo when they affect the
material structure or regeneration flow. Agent memory should not be the only
record of why a deck or page is shaped a certain way.
