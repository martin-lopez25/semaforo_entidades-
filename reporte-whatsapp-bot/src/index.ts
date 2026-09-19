import 'dotenv/config';
import cron from 'node-cron';
import path from 'node:path';
import { mkdir } from 'node:fs/promises';
import { captureReport, type ReportSection } from './capture.js';
import { connectWhatsApp, phoneJid, resolveGroupJid, sendImage } from './whatsapp.js';

const rootDirectory = path.resolve(import.meta.dirname, '..');
const authDirectory = path.join(rootDirectory, 'auth_reporte');
const capturesDirectory = path.join(rootDirectory, 'capturas');
const reportUrl = requiredEnv('REPORT_URL');
const groupName = requiredEnv('WHATSAPP_GRUPO_NOMBRE');
const additionalNumber = requiredEnv('WHATSAPP_NUMERO_ADICIONAL');
const section = (process.env.REPORT_SECTION ?? 'inventory') as ReportSection;
const schedule = process.env.CRON_SCHEDULE ?? '30 * * * *';
const timezone = process.env.TIMEZONE ?? 'America/Mexico_City';

function requiredEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Falta configurar ${name} en .env`);
  return value;
}

async function main(): Promise<void> {
  await mkdir(capturesDirectory, { recursive: true });
  console.log(`Seccion configurada: ${section}`);
  console.log(`Programacion: ${schedule} (${timezone})`);

  const socket = await connectWhatsApp(authDirectory);
  const groupJid = await resolveGroupJid(socket, groupName);
  const additionalJid = phoneJid(additionalNumber);
  console.log(`Grupo seleccionado: ${groupName} (${groupJid})`);
  let running = false;

  const sendReport = async (): Promise<void> => {
    if (running) {
      console.warn('El envio anterior sigue ejecutandose; se omite este ciclo.');
      return;
    }
    running = true;
    const capturePath = path.join(capturesDirectory, `reporte-${section}.png`);
    try {
      const capture = await captureReport(reportUrl, section, capturePath);
      await sendImage(socket, groupJid, capture.outputPath, `Reporte de inventario: ${capture.label}`);
      await sendImage(socket, additionalJid, capture.outputPath, `Reporte de inventario: ${capture.label}`);
      console.log(`PNG enviado correctamente: ${capture.outputPath}`);
    } catch (error) {
      console.error('Error al generar o enviar el reporte:', error);
    } finally {
      running = false;
    }
  };

  if (process.env.RUN_ON_START === 'true') {
    await sendReport();
  }

  cron.schedule(schedule, () => void sendReport(), { timezone });
  console.log('Bot activo. Presiona Ctrl+C para detenerlo.');
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
