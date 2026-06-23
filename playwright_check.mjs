import { chromium } from 'playwright';
import fs from 'node:fs/promises';
import path from 'node:path';

const root = process.cwd();
const imagePath = path.join(root, 'WhatsApp Image 2026-06-15 at 08.23.57.jpeg');

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 1800 } });

await page.goto('http://127.0.0.1:8000/', { waitUntil: 'networkidle' });

const productChoice = page.locator('#screen-1 .choice').filter({ hasText: 'Produse' });
if (await productChoice.count() !== 1) {
  throw new Error(`Expected one product choice card, got ${await productChoice.count()}`);
}
await productChoice.click();

const step1Continue = page.locator('#screen-1 button').filter({ hasText: 'Continua' });
if (await step1Continue.count() !== 1) {
  throw new Error(`Expected one step-1 continue button, got ${await step1Continue.count()}`);
}
await step1Continue.click();

await page.locator('#contract-no').fill('183/09.03.2026');
await page.locator('#offer-no').fill('PROF26 0142');

const step2Continue = page.locator('#screen-2 button').filter({ hasText: 'Continua' });
if (await step2Continue.count() !== 1) {
  throw new Error(`Expected one step-2 continue button, got ${await step2Continue.count()}`);
}
await step2Continue.waitFor({ state: 'visible', timeout: 5000 });
await step2Continue.click();

await page.setInputFiles('#screen-3 input[type="file"]', imagePath);
await page.waitForFunction(
  () => {
    const result = document.querySelector('#result-box');
    const json = document.querySelector('#offer-json');
    return Boolean(result && result.innerText.includes('Oferta procesata') && json && json.value.includes('"items"'));
  },
  { timeout: 30000 }
);

const rows = await page.locator('.table-row').evaluateAll((nodes) =>
  nodes.map((row) =>
    Array.from(row.querySelectorAll('input')).map((input) => input.value)
  )
);

const totals = await page.locator('#totals-panel').innerText().catch(() => '');
const jsonText = await page.locator('#offer-json').inputValue().catch(() => '');
const resultText = await page.locator('#result-box').innerText().catch(() => '');

const screenshotPath = path.join(root, 'playwright_check.png');
await page.screenshot({ path: screenshotPath, fullPage: true });

await browser.close();

await fs.writeFile(
  path.join(root, 'playwright_check.json'),
  JSON.stringify({ rows, totals, jsonText, resultText, screenshotPath }, null, 2),
  'utf8'
);

console.log(JSON.stringify({ rows, totals, resultText, screenshotPath }, null, 2));
