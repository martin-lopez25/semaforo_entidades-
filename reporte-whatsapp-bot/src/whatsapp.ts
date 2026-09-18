import makeWASocket, {
  DisconnectReason,
  useMultiFileAuthState,
  type WASocket,
} from '@whiskeysockets/baileys';
import { Boom } from '@hapi/boom';
import pino from 'pino';
import qrcode from 'qrcode-terminal';
import path from 'node:path';

const logger = pino({ level: 'silent' });

export async function connectWhatsApp(authDirectory: string): Promise<WASocket> {
  const { state, saveCreds } = await useMultiFileAuthState(authDirectory);
  const socket = makeWASocket({ auth: state, logger });
  socket.ev.on('creds.update', saveCreds);

  await new Promise<void>((resolve, reject) => {
    const onConnectionUpdate = ({ connection, lastDisconnect, qr }: { connection?: string; lastDisconnect?: { error?: unknown }; qr?: string }) => {
      if (qr) {
        console.log('\nEscanea este QR con el numero exclusivo del reporte:\n');
        qrcode.generate(qr, { small: true });
      }
      if (connection === 'open') {
        console.log('WhatsApp del reporte conectado.');
        resolve();
      }
      if (connection === 'close') {
        const statusCode = (lastDisconnect?.error as Boom | undefined)?.output?.statusCode;
        if (statusCode !== DisconnectReason.loggedOut) {
          reject(new Error('La conexion de WhatsApp se cerro; reinicia el proceso para reconectar.'));
        } else {
          reject(new Error('La sesion de WhatsApp fue cerrada. Elimina auth_reporte y escanea un QR nuevo.'));
        }
      }
    };
    socket.ev.on('connection.update', onConnectionUpdate);
  });

  return socket;
}

export function groupJid(value: string): string {
  const normalized = value.trim();
  if (!normalized.endsWith('@g.us')) {
    throw new Error('WHATSAPP_GRUPO debe ser un JID de grupo terminado en @g.us.');
  }
  return normalized;
}

export async function resolveGroupJid(socket: WASocket, groupName: string): Promise<string> {
  const groups = await socket.groupFetchAllParticipating();
  const normalizedName = groupName.trim().toLocaleLowerCase();
  const match = Object.entries(groups).find(([, metadata]) =>
    metadata.subject?.trim().toLocaleLowerCase() === normalizedName
  );

  if (!match) {
    const available = Object.values(groups)
      .map((metadata) => metadata.subject)
      .filter(Boolean)
      .sort()
      .join(', ');
    throw new Error(`No se encontro el grupo "${groupName}". Grupos disponibles: ${available}`);
  }

  return groupJid(match[0]);
}

export async function sendImage(socket: WASocket, group: string, imagePath: string, caption: string): Promise<void> {
  await socket.sendMessage(groupJid(group), {
    image: { url: path.resolve(imagePath) },
    caption,
  });
}
