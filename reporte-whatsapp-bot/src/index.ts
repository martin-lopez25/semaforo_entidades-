import 'dotenv/config';
import { createHash } from 'node:crypto';
import cron from 'node-cron';
import path from 'node:path';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { captureReport, type ReportSection } from './capture.js';
import { connectWhatsApp, resolveGroupJid, sendImage } from './whatsapp.js';

const rootDirectory = path.resolve(import.meta.dirname, '..');
const authDirectory = path.join(rootDirectory, 'auth_reporte');
const capturesDirectory = path.join(rootDirectory, 'capturas');
const fingerprintPath = path.join(capturesDirectory, 'reporte-data.sha256');
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

async function getReportDataFingerprint(): Promise<string> {
  const dataUrl = new URL('./reporte-inventario-data.json', reportUrl);
  dataUrl.searchParams.set('_botCheck', Date.now().toString());
  const response = await fetch(dataUrl, { headers: { 'Cache-Control': 'no-cache' } });
  if (!response.ok) {
    throw new Error(`No se pudieron consultar los datos del reporte: ${response.status}`);
  }

  const payload: unknown = await response.json();
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('El archivo de datos del reporte no tiene un formato valido.');
  }

  const data = payload as Record<string, unknown>;
  if (!Array.isArray(data.entities) || !Array.isArray(data.notReported) || !Array.isArray(data.incomplete)) {
    throw new Error('El archivo de datos no contiene todas las listas esperadas.');
  }

  const reportData = JSON.stringify({
    entities: data.entities,
    notReported: data.notReported,
    incomplete: data.incomplete,
  });
  return createHash('sha256').update(reportData).digest('hex');
}

async function readSavedFingerprint(): Promise<string | undefined> {
  try {
    return (await readFile(fingerprintPath, 'utf8')).trim() || undefined;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    throw error;
  }
}

function isConnectionClosedError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return error.message.includes('Connection Closed') || error.message.includes('428');
}

async function main(): Promise<void> {
  await mkdir(capturesDirectory, { recursive: true });
  console.log(`Seccion configurada: ${section}`);
  console.log(`Programacion: ${schedule} (${timezone})`);

  let socket = await connectWhatsApp(authDirectory);
  let groupJid = await resolveGroupJid(socket, groupName);
  console.log(`Grupo seleccionado: ${groupName} (${groupJid})`);
  let running = false;

  const reconnect = async (): Promise<void> => {
    console.warn('Conexion cerrada. Reconectando WhatsApp con la sesion existente...');
    socket = await connectWhatsApp(authDirectory);
    groupJid = await resolveGroupJid(socket, groupName);
    console.log(`WhatsApp reconectado. Grupo seleccionado: ${groupName} (${groupJid})`);
  };

  const sendReport = async (): Promise<void> => {
    if (running) {
      console.warn('El envio anterior sigue ejecutandose; se omite este ciclo.');
      return;
    }
    running = true;
    const capturePath = path.join(capturesDirectory, `reporte-${section}.png`);
    try {
      const fingerprint = await getReportDataFingerprint();
      const previousFingerprint = await readSavedFingerprint();
      if (fingerprint === previousFingerprint) {
        console.log('Los datos no cambiaron; se omiten la captura y el envio.');
        return;
      }

      const capture = await captureReport(reportUrl, section, capturePath);
      const caption = 'holis mando el Reporte de inventario: Inventario por entidad federativa';
      try {
        await sendImage(socket, groupJid, capture.outputPath, caption);
      } catch (error) {
        if (!isConnectionClosedError(error)) throw error;
        await reconnect();
        await sendImage(socket, groupJid, capture.outputPath, caption);
      }
      await writeFile(fingerprintPath, `${fingerprint}\n`, 'utf8');
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
