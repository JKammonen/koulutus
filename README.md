# koulutus

`koulutus` contains fiber-network training material as static HTML pages,
PDFs, and generated PowerPoint decks.

## Main Files

- `index.html` - entry page
- `perusteet.html` - fiber-network basics material
- `cwdm.html` and `wdm-keycom.html` - CWDM/DWDM and KeyCom-oriented material
- `generate_perusteet_pptx.js`, `generate_cwdm_pptx.js`, `generate_pdfs.js`
  - generation scripts
- `Kuituverkon_perusteet.pptx`, `CWDM_DWDM.pptx`, and matching PDFs -
  generated outputs

## Install

```bash
npm install
```

## Check

```bash
bash scripts/check-repo.sh
```

This repository has no automated content-quality test yet. The local check
validates generator syntax and required source/output files.
