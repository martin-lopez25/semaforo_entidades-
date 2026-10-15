# Reporte WhatsApp Bot

Proceso independiente del chatbot principal. Abre la pagina publicada, captura inventario, CLUES que no reportaron y CLUES incompletos, y envia los tres PNG por WhatsApp cada hora.

## Configuracion

1. Copia `.env.example` como `.env`.
2. Define la URL de GitHub Pages y el nombre exacto del grupo en `WHATSAPP_GRUPO_NOMBRE`. El bot buscará automáticamente su JID.
3. Instala dependencias: `npm install`.
4. Instala Chromium: `npx playwright install chromium`.
5. Ejecuta `npm run dev` y escanea el QR con el numero origen del reporte.
6. La sesion se guarda en `auth_reporte/`, separada del chatbot.

El proceso envia las tres imagenes segun `CRON_SCHEDULE`. Para enviarlas inmediatamente al arrancar, usa `RUN_ON_START=true`.

Para construir y ejecutar con Node:

```powershell
npm run build
npm start
```
