# koulutus Decisions

## 2026-06-12: Keep project ownership docs in the repo

- Decision: root `README.md`, `ARCHITECTURE.md`, `DECISIONS.md`,
  `AGENTS.md`, and `scripts/check-repo.sh` are the minimum local ownership
  surface.
- Why: this is a maintenance repo, but generated deliverables still need a
  local source-of-truth path.
- Verification: `scripts/check-repo.sh` is the local check entrypoint, and the
  shared infra audit checks that this ownership surface exists.

## Settled Working Decisions

- Keep training pages as static HTML.
- Keep generated PPTX/PDF artifacts committed because they are deliverables.
- Use Node-based generators for decks and PDFs.
- A syntax/file-presence gate is acceptable until content-quality checks are
  worth automating.
