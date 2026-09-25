import 'dotenv/config';
import cron from 'node-cron';
import path from 'node:path';
import { mkdir } from 'node:fs/promises';
import { captureReport, type ReportSection } from './capture.js';
import { connectWhatsApp, resolveGroupJid, sendImage } from './whatsapp.js';

const rootDirectory = path.resolve(import.meta.dirname, '..');
const authDirectory = path.join(rootDirectory, 'auth_reporte');
const capturesDirectory = path.join(rootDirectory, 'capturas');
const reportUrl = requiredEnv('REPORT_URL');
const groupName = requiredEnv('WHATSAPP_GRUPO_NOMBRE');
const section = (process.env.REPORT_SECTION ?? 'inventory') as ReportSection;
const schedule = process.env.CRON_SCHEDULE ?? '30 * * * *';
const shutdownSchedule = process.env.AUTO_SHUTDOWN_SCHEDULE ?? '0 22 * * *';
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
      const caption = 'holis mando el Reporte de inventario: Inventario por entidad federativa';
      await sendImage(socket, groupJid, capture.outputPath, caption);
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
  cron.schedule(shutdownSchedule, () => {
    console.log('Apagado programado del bot. La sesion de WhatsApp se conserva.');
    process.exit(0);
  }, { timezone });
  console.log(`Apagado programado: ${shutdownSchedule} (${timezone})`);
  console.log('Bot activo. Presiona Ctrl+C para detenerlo.');
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
