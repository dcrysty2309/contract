import { chromium } from 'playwright';

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto('file:///D:/proiecte/contracte/.tmp/pv_line_preview.html', { waitUntil: 'load' });
await page.screenshot({ path: 'D:/proiecte/contracte/.tmp/pv_line_preview.png', fullPage: true });
await browser.close();
console.log('captured');
