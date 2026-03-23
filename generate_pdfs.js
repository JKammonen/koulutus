const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function generatePDF(htmlFile, outputFile) {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  
  const filePath = 'file://' + path.resolve(htmlFile);
  await page.goto(filePath, { waitUntil: 'networkidle0', timeout: 30000 });
  await sleep(2000);
  
  // Expand all <details> elements
  await page.evaluate(() => {
    document.querySelectorAll('details').forEach(d => d.open = true);
  });
  await sleep(500);
  
  // Hide fixed/sticky nav elements
  await page.evaluate(() => {
    document.querySelectorAll('*').forEach(el => {
      const style = getComputedStyle(el);
      if (style.position === 'fixed' || style.position === 'sticky') {
        el.style.display = 'none';
      }
    });
  });
  
  await page.pdf({
    path: outputFile,
    format: 'A4',
    printBackground: true,
    margin: { top: '15mm', right: '15mm', bottom: '15mm', left: '15mm' },
  });
  
  await browser.close();
  const size = (fs.statSync(outputFile).size / 1024).toFixed(0);
  console.log(`✅ ${outputFile} (${size} KB)`);
}

(async () => {
  await generatePDF('perusteet.html', 'Kuituverkon_perusteet.pdf');
  await generatePDF('cwdm.html', 'CWDM_DWDM.pdf');
})();
