import { chromium } from 'playwright';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';

export type ReportSection = 'inventory' | 'not-reported' | 'incomplete' | 'chart' | 'full';

const sectionConfig: Record<ReportSection, { hash: string; selector: string; label: string }> = {
  inventory: { hash: '#capture-inventory', selector: '#capture-inventory', label: 'Inventario por entidad federativa' },
  'not-reported': { hash: '#section-no-reportaron', selector: '#section-no-reportaron', label: 'CLUES que no reportaron' },
  incomplete: { hash: '#section-clues-incompletos', selector: '#section-clues-incompletos', label: 'CLUES incompletos' },
  chart: { hash: '#capture-chart', selector: '#capture-chart', label: 'Grafica de avance' },
  full: { hash: '', selector: '#capture-full-report', label: 'Reporte completo' },
};

function buildUrl(reportUrl: string, hash: string): string {
  const url = new URL(reportUrl);
  url.hash = hash.replace(/^#/, '');
  return url.toString();
}

export async function captureReport(
  reportUrl: string,
  section: ReportSection,
  outputPath: string,
): Promise<{ outputPath: string; label: string }> {
  const config = sectionConfig[section];
  if (!config) throw new Error(`Seccion no soportada: ${section}`);

  await mkdir(path.dirname(outputPath), { recursive: true });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ deviceScaleFactor: 1.5, viewport: { width: 1440, height: 1000 } });
    await page.goto(buildUrl(reportUrl, config.hash), { waitUntil: 'networkidle' });
    await page.waitForTimeout(1500);

    const target = page.locator(config.selector);
    await target.waitFor({ state: 'visible', timeout: 15000 });
    await target.screenshot({ path: outputPath });

    return { outputPath, label: config.label };
  } finally {
    await browser.close();
  }
}
