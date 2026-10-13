/**
 * ============================================================================
 * PUERCO & PUERCA — SISTEMA INTEGRADO DE PEDIDOS AL MAYOR (B2B)
 * Google Sheets + Google Calendar + Notificación Instantánea por Correo
 * ============================================================================
 * 
 * ¿DÓNDE ES MEJOR CONECTARLA? (RESPUESTA: ¡EN AMBOS A LA VEZ!)
 * 
 * 1. GOOGLE SHEETS: Funciona como tu BASE DE DATOS MAESTRA y libro contable.
 *    - Lleva el control de todas las ventas, total de dólares recaudados,
 *      inventario de cajas vendidas, datos fiscales y lista de clientes.
 * 
 * 2. GOOGLE CALENDAR: Funciona como tu AGENDA OPERATIVA DE DESPACHOS.
 *    - Te crea automáticamente un evento en tu calendario de Google en la fecha
 *      y turno (Mañana o Tarde), recordándote en tu teléfono celular qué despachos
 *      debes entregar cada día con la ubicación exacta en Google Maps.
 * 
 * Este script realiza AMBAS cosas de forma 100% automática y simultánea con un solo clic.
 * 
 * ============================================================================
 * GUÍA DE INSTALACIÓN RÁPIDA (2 minutos):
 * ============================================================================
 * 1. Entra a Google Sheets (https://sheets.new) y crea una hoja nueva.
 *    Nómbrala: "Pedidos al Mayor - Puerco & Puerca".
 * 
 * 2. En el menú superior de la hoja, haz clic en:
 *    Extensiones -> Apps Script
 * 
 * 3. Borra todo el código que aparezca y PEGA TODO EL CONTENIDO DE ESTE ARCHIVO.
 * 
 * 4. Configura tu correo electrónico en la variable NOTIFICACION_EMAIL (abajo).
 * 
 * 5. Haz clic en "Guardar" (icono de disquete 💾).
 * 
 * 6. Haz clic en el botón azul superior "Implementar" -> "Nueva implementación".
 *    - Tipo de implementación: Selecciona "Aplicación web" (icono de engranaje).
 *    - Descripción: "Webhook Pedidos al Mayor v2".
 *    - Ejecutar como: "Yo" (tu cuenta de Google).
 *    - Quién tiene acceso: "Cualquier persona" (Anyone). [¡MUY IMPORTANTE!]
 *    - Haz clic en "Implementar" y autoriza los permisos de Google.
 * 
 * 7. Copia la "URL de la aplicación web" (termina en /exec) y pégala en index.html:
 *    Busca: const GOOGLE_SHEETS_WEBHOOK_URL = '...';
 * ============================================================================
 */

// ── CONFIGURACIÓN COMERCIAL ──
const NOTIFICACION_EMAIL = "yonathanaguilar1000@gmail.com"; // Tu correo para recibir las alertas
const NOMBRE_HOJA = "Pedidos al Mayor"; // Nombre de la pestaña en la hoja de cálculo
const CREAR_EVENTO_EN_CALENDAR = true; // Activar creación automática en Google Calendar

/**
 * Maneja las peticiones POST enviadas desde el formulario web
 */
function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // Evitar colisiones si entran dos pedidos a la vez

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(NOMBRE_HOJA);
    
    // Si la pestaña no existe, la crea con encabezados profesionales
    if (!sheet) {
      sheet = ss.insertSheet(NOMBRE_HOJA);
      crearEncabezados(sheet);
    } else if (sheet.getLastRow() === 0) {
      crearEncabezados(sheet);
    }

    // Parsear datos recibidos
    let data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    const timestamp = new Date();
    const fechaDespacho = data.displayDate || data.deliveryDate || 'Por coordinar';
    const fechaIso = data.deliveryDate || '';
    const turno = data.slot || 'Mañana (8:00 AM - 12:00 PM)';
    const cajas = data.boxes || 1;
    const totalFrascos = data.totalUnits || (cajas * 12);
    const totalDolares = data.totalPrice || (cajas * 45.60).toFixed(2);
    const negocio = data.bizName || 'Sin especificar';
    const contacto = data.contact || 'Sin especificar';
    const telefono = data.phone || 'Sin especificar';
    const email = data.email || 'Sin especificar';
    const direccion = data.address || 'Sin especificar';
    const mapLink = data.mapLink || `https://maps.google.com/?q=${direccion}`;
    const notas = data.notes || 'Ninguna';
    const estado = 'Nuevo Pedido (Por despachar)';

    // 1. REGISTRAR FILA EN GOOGLE SHEETS
    sheet.appendRow([
      timestamp,
      fechaDespacho,
      turno,
      cajas,
      totalFrascos,
      3.80,
      parseFloat(totalDolares),
      negocio,
      contacto,
      telefono,
      email,
      direccion,
      mapLink,
      notas,
      estado
    ]);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(lastRow, 6).setNumberFormat("$#,##0.00");
    sheet.getRange(lastRow, 7).setNumberFormat("$#,##0.00");

    lock.releaseLock();

    // 2. CREAR EVENTO AUTOMÁTICO EN GOOGLE CALENDAR
    let calendarEventId = "";
    if (CREAR_EVENTO_EN_CALENDAR && fechaIso) {
      calendarEventId = agendarEnGoogleCalendar({
        fechaIso: fechaIso,
        fechaDespacho: fechaDespacho,
        turno: turno,
        cajas: cajas,
        totalFrascos: totalFrascos,
        totalDolares: totalDolares,
        negocio: negocio,
        contacto: contacto,
        telefono: telefono,
        direccion: direccion,
        mapLink: mapLink,
        notas: notas
      });
    }

    // 3. ENVIAR NOTIFICACIÓN POR CORREO ELECTRÓNICO
    enviarAlertaCorreo({
      fechaDespacho: fechaDespacho,
      turno: turno,
      cajas: cajas,
      totalFrascos: totalFrascos,
      totalDolares: totalDolares,
      negocio: negocio,
      contacto: contacto,
      telefono: telefono,
      email: email,
      direccion: direccion,
      mapLink: mapLink,
      notas: notas,
      timestamp: timestamp,
      calendarCreated: !!calendarEventId
    });

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Pedido guardado en Google Sheets y agendado en Google Calendar", 
        row: lastRow,
        calendarEventId: calendarEventId 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Crea automáticamente el evento en Google Calendar para el equipo de despacho
 */
function agendarEnGoogleCalendar(order) {
  try {
    const parts = order.fechaIso.split('-');
    if (parts.length < 3) return "";

    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);

    // Ajustar horario según el turno solicitado
    let startHour = 8;
    let endHour = 12;
    if (order.turno && order.turno.toLowerCase().includes('tarde')) {
      startHour = 13;
      endHour = 17;
    }

    const startTime = new Date(year, month, day, startHour, 0, 0);
    const endTime = new Date(year, month, day, endHour, 0, 0);

    const titulo = `🚚 Despacho Manteca: ${order.negocio} (${order.cajas} Cajas - $${order.totalDolares})`;

    const descripcion = 
      `📦 ORDEN AL MAYOR — PUERCO & PUERCA\n\n` +
      `🏢 Negocio: ${order.negocio}\n` +
      `👤 Contacto: ${order.contacto}\n` +
      `📱 Teléfono: ${order.telefono}\n` +
      `📦 Cajas: ${order.cajas} Cajas (${order.totalFrascos} frascos de 500g)\n` +
      `💵 Total Facturado: $${order.totalDolares}\n` +
      `⏰ Turno: ${order.turno}\n` +
      `📍 Dirección: ${order.direccion}\n` +
      (order.mapLink ? `🗺️ Ver en Google Maps: ${order.mapLink}\n` : '') +
      `📝 Notas: ${order.notas}\n\n` +
      `Chat WhatsApp con cliente: https://wa.me/${order.telefono.replace(/[^0-9]/g, '')}`;

    const calendar = CalendarApp.getDefaultCalendar();
    if (!calendar) return "";

    const event = calendar.createEvent(titulo, startTime, endTime, {
      description: descripcion,
      location: order.direccion
    });

    // Recordatorios en el teléfono
    event.addPopupReminder(1440); // 1 día antes
    event.addPopupReminder(60);   // 1 hora antes

    return event.getId();
  } catch (err) {
    Logger.log("Error creando evento en Calendar: " + err);
    return "";
  }
}

/**
 * Endpoint GET para probar que la aplicación web esté activa
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: "online", 
      sistema: "Pedidos al Mayor Puerco & Puerca (Sheets + Calendar)",
      timestamp: new Date().toISOString() 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Crea los encabezados estilizados de la hoja de cálculo
 */
function crearEncabezados(sheet) {
  const headers = [
    "Marca Temporal",
    "Fecha Despacho",
    "Turno Solicitado",
    "Cajas (×12)",
    "Total Frascos (500g)",
    "Precio / Frasco ($)",
    "Total a Facturar ($)",
    "Nombre del Negocio",
    "Persona de Contacto",
    "Teléfono / WhatsApp",
    "Correo Electrónico",
    "Dirección del Local",
    "Enlace Google Maps",
    "Notas Especiales",
    "Estado de la Orden"
  ];
  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#150d06");
  headerRange.setFontColor("#d4a24c");
  headerRange.setFontWeight("bold");
  headerRange.setFontFamily("Arial");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Envía un correo con diseño corporativo al dueño con acceso a Maps y WhatsApp
 */
function enviarAlertaCorreo(order) {
  if (!NOTIFICACION_EMAIL || NOTIFICACION_EMAIL.includes("ejemplo")) return;

  const asunto = `🚨 NUEVO PEDIDO AL MAYOR: ${order.negocio} (${order.cajas} Cajas - $${order.totalDolares})`;
  const waLink = `https://wa.me/${order.telefono.replace(/[^0-9]/g, '')}`;

  const htmlBody = `
    <div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;max-width:620px;margin:0 auto;background:#ffffff;border:1px solid #e0d0b8;border-radius:12px;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,0.08);">
      
      <!-- Cabecera -->
      <div style="background:#150d06;padding:26px 24px;text-align:center;border-bottom:3px solid #d4a24c;">
        <h1 style="color:#d4a24c;margin:0;font-size:22px;letter-spacing:1px;">🐷 PUERCO & PUERCA</h1>
        <p style="color:#fdf8ef;margin:6px 0 0;font-size:13px;opacity:0.85;">NOTIFICACIÓN DE NUEVO PEDIDO AL MAYOR (B2B)</p>
      </div>

      <!-- Resumen Destacado -->
      <div style="padding:24px;">
        <div style="background:#fcf9f2;border:1.5px solid #d4a24c;border-radius:10px;padding:16px 20px;margin-bottom:20px;text-align:center;">
          <span style="font-size:12px;color:#856427;text-transform:uppercase;font-weight:bold;letter-spacing:0.5px;">Total de la Orden</span>
          <div style="font-size:32px;font-weight:800;color:#150d06;margin:4px 0;">$${order.totalDolares}</div>
          <div style="font-size:14px;color:#4a3520;font-weight:600;">
            📦 ${order.cajas} Cajas &nbsp;•&nbsp; 🫙 ${order.totalFrascos} Frascos (500 g)
          </div>
          ${order.calendarCreated ? '<div style="margin-top:6px;font-size:12px;color:#2e7d32;font-weight:bold;">📅 Agendado en tu Google Calendar con recordatorio</div>' : ''}
        </div>

        <!-- Tabla de Datos -->
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:14px;">
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;width:40%;">📅 Fecha Despacho:</td>
            <td style="padding:10px 6px;color:#111;font-weight:bold;">${order.fechaDespacho}</td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">⏰ Turno de Entrega:</td>
            <td style="padding:10px 6px;color:#111;font-weight:600;">${order.turno}</td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">🏬 Negocio / Razón:</td>
            <td style="padding:10px 6px;color:#111;font-weight:bold;">${order.negocio}</td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">👤 Encargado / Contacto:</td>
            <td style="padding:10px 6px;color:#111;">${order.contacto}</td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">📱 Teléfono / WhatsApp:</td>
            <td style="padding:10px 6px;color:#111;font-weight:600;">
              <a href="${waLink}" style="color:#1d6f42;text-decoration:none;">${order.telefono} (Escribir en WhatsApp)</a>
            </td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">✉️ Correo Electrónico:</td>
            <td style="padding:10px 6px;color:#111;">${order.email}</td>
          </tr>
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">📍 Dirección del Local:</td>
            <td style="padding:10px 6px;color:#111;">${order.direccion}</td>
          </tr>
          ${order.mapLink ? `
          <tr style="border-bottom:1px solid #eee;">
            <td style="padding:10px 6px;color:#777;">🗺️ Ubicación GPS:</td>
            <td style="padding:10px 6px;">
              <a href="${order.mapLink}" target="_blank" style="color:#1a73e8;font-weight:bold;text-decoration:none;">Abrir en Google Maps ↗</a>
            </td>
          </tr>` : ''}
          <tr>
            <td style="padding:10px 6px;color:#777;">📝 Notas Especiales:</td>
            <td style="padding:10px 6px;color:#111;font-style:italic;">${order.notas}</td>
          </tr>
        </table>

        <!-- Botones de Acción -->
        <div style="text-align:center;margin:25px 0 10px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
          <a href="${waLink}" style="background:#25D366;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:50px;font-weight:bold;font-size:14px;display:inline-block;box-shadow:0 4px 15px rgba(37,211,102,0.3);">
            💬 Chatear con Cliente en WhatsApp
          </a>
          ${order.mapLink ? `
          <a href="${order.mapLink}" target="_blank" style="background:#1a73e8;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:50px;font-weight:bold;font-size:14px;display:inline-block;box-shadow:0 4px 15px rgba(26,115,232,0.3);">
            🗺️ Ver Mapa de Entrega
          </a>` : ''}
        </div>
      </div>

      <!-- Pie de Correo -->
      <div style="background:#f9f6f0;padding:14px 20px;text-align:center;font-size:11px;color:#998877;border-top:1px solid #ece4d8;">
        Este pedido se registró automáticamente en Google Sheets y en tu Google Calendar.<br>
        © 2026 Manteca Puerco y Puerca — Valencia, Carabobo.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: NOTIFICACION_EMAIL,
    subject: asunto,
    htmlBody: htmlBody
  });
}
