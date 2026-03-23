const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function generatePDF(htmlFile, outputFile) {
  console.log(`⏳ ${htmlFile} → ${outputFile}...`);
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();

  // Wide viewport for proper layout
  await page.setViewport({ width: 1280, height: 900 });

  const filePath = 'file://' + path.resolve(htmlFile);
  await page.goto(filePath, { waitUntil: 'networkidle0', timeout: 30000 });
  await sleep(2000);

  // Prepare page for PDF: expand details, hide nav, fix layout
  await page.evaluate(() => {
    // Expand all <details> elements
    document.querySelectorAll('details').forEach(d => d.open = true);

    // Hide fixed/sticky nav and print button
    document.querySelectorAll('*').forEach(el => {
      const style = getComputedStyle(el);
      if (style.position === 'fixed' || style.position === 'sticky') {
        el.style.display = 'none';
      }
    });

    // Remove min-height: 100vh from sections — prevents huge empty spaces
    document.querySelectorAll('.section').forEach(s => {
      s.style.minHeight = 'auto';
      s.style.pageBreakInside = 'avoid';
    });

    // Page break control: avoid breaking inside key elements
    document.querySelectorAll('.card, .diagram-box, .data-table, table, .insight-box, .summary-box, .device-frame, .oadm-frame, .circuit-diagram, .hierarchy, .step-list, .spectrum-box, .calc-container').forEach(el => {
      el.style.pageBreakInside = 'avoid';
    });

    // Each major section starts on a new page (except first)
    const sections = document.querySelectorAll('.section');
    sections.forEach((s, i) => {
      if (i > 0) {
        s.style.pageBreakBefore = 'always';
      }
    });

    // Part headers also start new pages
    document.querySelectorAll('.part-header').forEach(el => {
      el.style.pageBreakBefore = 'always';
    });

    // Remove animations
    document.querySelectorAll('*').forEach(el => {
      el.style.animation = 'none';
      el.style.transition = 'none';
    });

    // Remove hover effects on cards
    const style = document.createElement('style');
    style.textContent = `
      .card:hover { transform: none !important; box-shadow: none !important; }
      .data-table tr:hover td { background: transparent !important; }
      footer { page-break-before: always; }
    `;
    document.head.appendChild(style);
  });
  await sleep(500);

  await page.pdf({
    path: outputFile,
    format: 'A4',
    printBackground: true,
    margin: { top: '15mm', right: '15mm', bottom: '20mm', left: '15mm' },
    displayHeaderFooter: true,
    headerTemplate: '<div></div>',
    footerTemplate: `
      <div style="width: 100%; font-size: 9px; color: #94a3b8; padding: 0 15mm; display: flex; justify-content: space-between;">
        <span>${path.basename(htmlFile, '.html')}</span>
        <span><span class="pageNumber"></span> / <span class="totalPages"></span></span>
      </div>
    `,
  });

  await browser.close();
  const size = (fs.statSync(outputFile).size / 1024).toFixed(0);
  console.log(`✅ ${outputFile} (${size} KB)`);
}

// Generate all or specific PDF
const args = process.argv.slice(2);

const courses = [
  { html: 'perusteet.html', pdf: 'Kuituverkon_perusteet.pdf' },
  { html: 'cwdm.html', pdf: 'CWDM_DWDM.pdf' },
  { html: 'wdm-keycom.html', pdf: 'WDM_KeyComissa.pdf' },
];

(async () => {
  const targets = args.length > 0
    ? courses.filter(c => args.some(a => c.html.includes(a) || c.pdf.includes(a)))
    : courses;

  if (targets.length === 0) {
    console.log('Käyttö: node generate_pdfs.js [tiedosto]');
    console.log('Ilman argumentteja generoidaan kaikki PDF:t.');
    console.log('Kurssit:', courses.map(c => c.html).join(', '));
    process.exit(1);
  }

  for (const { html, pdf } of targets) {
    await generatePDF(html, pdf);
  }
  console.log(`\n✅ Valmis! ${targets.length} PDF generoitu.`);
})();
