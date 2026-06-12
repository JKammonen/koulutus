#!/bin/bash

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "${ROOT}"

for file in \
  index.html \
  perusteet.html \
  cwdm.html \
  wdm-keycom.html \
  generate_perusteet_pptx.js \
  generate_cwdm_pptx.js \
  generate_pdfs.js \
  Kuituverkon_perusteet.pptx \
  Kuituverkon_perusteet.pdf \
  CWDM_DWDM.pptx \
  CWDM_DWDM.pdf
do
  test -f "${file}"
done

node --check generate_perusteet_pptx.js
node --check generate_cwdm_pptx.js
node --check generate_pdfs.js
