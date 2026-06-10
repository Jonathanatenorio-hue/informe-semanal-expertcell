// ============================================================
//  ExpertCell · Envio diario de reportes (Opcion A reforzada)
// ============================================================
//
//  QUE HACE:
//    1. Abre tu reporte.html PUBLICADO (solo lectura) con un navegador headless.
//    2. Por cada centro (CC2 y JV) aplica el filtro de coordinador igual que
//       cuando tu lo eliges en el menu, y captura:
//         - El Excel del coordinador (dashboard + 1 hoja por supervisor + detalle
//           general del centro)  -> tu funcion descargarCoordinadorExcel()
//         - El Excel de inactividad del centro  -> tu funcion descargarInactividadExcel()
//         - Un Excel de la TABLA MAESTRA del centro
//    3. Envia un correo por centro al coordinador, con copia a Jonathan.
//
//  GARANTIA: NO modifica reporte.html ni Supabase. Solo lee. Si la pagina no
//  carga con datos, NO envia nada a los coordinadores y AVISA solo a Jonathan.
//
//  Variables de entorno requeridas (se configuran como Secrets en GitHub):
//    GMAIL_USER          -> jonathan.atenorio@gmail.com
//    GMAIL_APP_PASSWORD  -> la clave de aplicacion de 16 letras (sin espacios)
// ============================================================

import { chromium } from 'playwright';
import nodemailer from 'nodemailer';

// ---------- CONFIGURACION ----------
const URL_REPORTE =
  'https://jonathanatenorio-hue.github.io/informe-semanal-expertcell/reporte.html';

// Copia fija en TODOS los correos (incluye los avisos a Jonathan).
const CC_SIEMPRE = 'j.tenorio@expertcell.com.mx';

// Ruteo por CENTRO (clave estable). Si un coordinador no cae aqui,
// NO se le envia y se reporta como incidencia a Jonathan.
const RUTEO_POR_CENTRO = {
  CC2: { para: 'oficinacc2@solucell.com.mx', etiqueta: 'CC2 Antequera' },
  JV:  { para: 'l.meza@solucell.com.mx',      etiqueta: 'JV' },
};

const GMAIL_USER = process.env.GMAIL_USER;
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;
const REMITENTE = `"Reportes ExpertCell" <${GMAIL_USER}>`;

const TZ = 'America/Mexico_City';
const HOY = new Date().toLocaleDateString('es-MX', {
  timeZone: TZ, year: 'numeric', month: 'long', day: 'numeric',
});
const HOY_ARCHIVO = new Date().toLocaleDateString('en-CA', { timeZone: TZ }); // YYYY-MM-DD

// MODO PRUEBA: si existe el secret MODO_PRUEBA (cualquier valor), TODOS los correos
// se redirigen unicamente a Jonathan (CC_SIEMPRE). Los coordinadores NO reciben nada.
// Para volver al envio normal, basta con borrar ese secret.
const MODO_PRUEBA = !!process.env.MODO_PRUEBA;

// ---------- HELPERS ----------
function clasificarCentro(raw) {
  const c = (raw || '').toUpperCase();
  if (c.includes('CC2')) return 'CC2';
  if (c.includes('JV')) return 'JV';
  return null;
}

function construirCuerpo(nombreCoord, etiquetaCentro, hayInact) {
  const inactLinea = hayInact
    ? 'Excel de inactividad del centro (agentes con dias sin activar / sin programar).'
    : 'Hoy no hubo agentes en alerta de inactividad. \u2705';
  const text =
`Hola ${nombreCoord},

Adjunto el reporte operativo de hoy (${HOY}) para ${etiquetaCentro}:

- Excel del centro: dashboard del coordinador, una hoja de detalle por cada supervisor y el detalle general del centro.
- Tabla maestra del centro (Excel).
- ${inactLinea}

Saludos,
Reportes ExpertCell (envio automatico)`;

  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola ${nombreCoord},</p>
  <p>Adjunto el reporte operativo de hoy (<strong>${HOY}</strong>) para <strong>${etiquetaCentro}</strong>:</p>
  <ul>
    <li><strong>Excel del centro:</strong> dashboard del coordinador, una hoja de detalle por cada supervisor y el detalle general del centro.</li>
    <li><strong>Tabla maestra</strong> del centro (Excel adjunto).</li>
    <li>${inactLinea}</li>
  </ul>
  <p style="color:#6b7280;font-size:12px;margin-top:18px">Reportes ExpertCell &middot; env&iacute;o autom&aacute;tico</p>
</div>`;
  return { text, html };
}

async function avisarJonathan(transporter, asunto, mensaje) {
  try {
    await transporter.sendMail({
      from: REMITENTE,
      to: CC_SIEMPRE,
      subject: asunto,
      text: `${mensaje}\n\n(Los coordinadores NO recibieron ningun reporte incompleto.)`,
    });
  } catch (e) {
    console.error('No se pudo enviar aviso a Jonathan:', e.message);
  }
}

// ---------- PRINCIPAL ----------
async function main() {
  if (!GMAIL_USER || !GMAIL_APP_PASSWORD) {
    console.error('Faltan GMAIL_USER / GMAIL_APP_PASSWORD en los Secrets.');
    process.exit(1);
  }

  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    auth: { user: GMAIL_USER, pass: GMAIL_APP_PASSWORD },
  });

  const browser = await chromium.launch();
  const ctx = await browser.newContext({
    viewport: { width: 1680, height: 1050 },
    deviceScaleFactor: 2,
  });
  const page = await ctx.newPage();
  // Evitar que cualquier alert() del reporte bloquee la ejecucion.
  page.on('dialog', (d) => d.dismiss().catch(() => {}));

  const incidencias = [];

  // --- Cargar el reporte y verificar que trae datos ---
  try {
    await page.goto(URL_REPORTE, { waitUntil: 'domcontentloaded', timeout: 120000 });
    await page.waitForFunction(() => {
      const r = document.getElementById('reporte');
      const t = document.getElementById('avance-tabla');
      return r && getComputedStyle(r).display !== 'none'
          && t && t.querySelectorAll('tr').length > 1;
    }, { timeout: 120000 });
  } catch (e) {
    await browser.close();
    await avisarJonathan(
      transporter,
      `\u26A0\uFE0F Reportes ExpertCell - el reporte no cargo (${HOY})`,
      `El reporte no cargo correctamente con datos, por lo que NO se envio nada a los coordinadores.\n\nDetalle tecnico: ${e.message}`
    );
    process.exit(1);
  }

  // --- Instalar captura de Excel (intercepta XLSX.writeFile, no descarga) ---
  await page.evaluate(() => {
    window.__captured = [];
    if (!window.XLSX.__patched) {
      window.XLSX.__patched = true;
      window.XLSX.writeFile = function (wb, filename) {
        try {
          const b64 = window.XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
          window.__captured.push({ filename: filename || 'archivo.xlsx', b64 });
        } catch (err) {
          window.__captured.push({ error: String(err) });
        }
        return true;
      };
    }
    window.alert = function () {};
  });

  // --- Descubrir coordinadores desde el menu del reporte ---
  const coords = await page.evaluate(() => {
    const sel = document.getElementById('coordinador-select');
    if (!sel) return [];
    return Array.from(sel.options)
      .filter((o) => o.value && o.value !== 'todos')
      .map((o) => ({ id: o.value, label: (o.textContent || '').trim() }));
  });

  if (!coords.length) {
    await browser.close();
    await avisarJonathan(
      transporter,
      `\u26A0\uFE0F Reportes ExpertCell - sin coordinadores (${HOY})`,
      'No se encontraron coordinadores en el menu del reporte. No se envio nada.'
    );
    process.exit(1);
  }

  let enviados = 0;

  for (const c of coords) {
    const m = c.label.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
    const nombre = m ? m[1].trim() : c.label;
    const centroRaw = m ? m[2].trim() : '';
    const centroKey = clasificarCentro(centroRaw);

    if (!centroKey || !RUTEO_POR_CENTRO[centroKey]) {
      incidencias.push(`Coordinador "${nombre}" (centro "${centroRaw}") sin ruteo configurado; se omitio.`);
      continue;
    }
    const ruta = RUTEO_POR_CENTRO[centroKey];

    try {
      // Aplicar el filtro de este coordinador (mismo camino que el menu humano)
      await page.evaluate((id) => {
        const sel = document.getElementById('coordinador-select');
        sel.value = id;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
      }, c.id);
      await page.waitForTimeout(2800); // dejar que recalcule y re-renderice

      // 1) Excel del coordinador (OBLIGATORIO)
      await page.evaluate(() => { window.__captured = []; });
      await page.evaluate((nom) => { descargarCoordinadorExcel(nom); }, nombre);
      const capCoord = await page.evaluate(() => window.__captured.slice());
      const archCoord = capCoord.find((x) => x.b64) || null;

      // 2) Excel de inactividad del centro (OPCIONAL: puede no haber alertas)
      await page.evaluate(() => { window.__captured = []; });
      await page.evaluate(() => { try { descargarInactividadExcel(); } catch (e) {} });
      const capInact = await page.evaluate(() => window.__captured.slice());
      const archInact = capInact.find((x) => x.b64) || null;

      // 3) Excel de la TABLA MAESTRA del centro (OBLIGATORIO)
      await page.evaluate(() => { window.__captured = []; });
      await page.evaluate(() => { try { descargarMaestraExcel(); } catch (e) {} });
      const capMaestra = await page.evaluate(() => window.__captured.slice());
      const archMaestra = capMaestra.find((x) => x.b64) || null;

      // Validacion: sin Excel de coordinador o sin maestra, NO se envia.
      if (!archCoord || !archMaestra) {
        incidencias.push(
          `Paquete incompleto para "${nombre}" (${ruta.etiqueta}); NO se envio. ` +
          `(Excel coordinador: ${archCoord ? 'ok' : 'FALTA'}, Tabla maestra: ${archMaestra ? 'ok' : 'FALTA'})`
        );
        continue;
      }

      const attachments = [
        {
          filename: `Reporte_${centroKey}_${HOY_ARCHIVO}.xlsx`,
          content: Buffer.from(archCoord.b64, 'base64'),
        },
        {
          filename: `Tabla_Maestra_${centroKey}_${HOY_ARCHIVO}.xlsx`,
          content: Buffer.from(archMaestra.b64, 'base64'),
        },
      ];
      if (archInact) {
        attachments.push({
          filename: `Inactividad_${centroKey}_${HOY_ARCHIVO}.xlsx`,
          content: Buffer.from(archInact.b64, 'base64'),
        });
      }

      const cuerpo = construirCuerpo(nombre, ruta.etiqueta, !!archInact);

      // En MODO PRUEBA todo va solo a Jonathan; en normal, al coordinador con copia.
      const destinoTo = MODO_PRUEBA ? CC_SIEMPRE : ruta.para;
      const destinoCc = MODO_PRUEBA ? undefined : CC_SIEMPRE;
      const prefijoAsunto = MODO_PRUEBA ? '[PRUEBA] ' : '';

      await transporter.sendMail({
        from: REMITENTE,
        to: destinoTo,
        cc: destinoCc,
        subject: `${prefijoAsunto}Reporte diario ExpertCell \u00B7 ${ruta.etiqueta} \u00B7 ${HOY}`,
        text: cuerpo.text,
        html: cuerpo.html,
        attachments,
      });

      enviados++;
      console.log(`Enviado a ${destinoTo} (${ruta.etiqueta})${MODO_PRUEBA ? ' [PRUEBA]' : ''}.`);
    } catch (e) {
      incidencias.push(`Error con "${nombre}" (${ruta.etiqueta}): ${e.message}. NO se le envio.`);
    }
  }

  await browser.close();

  if (incidencias.length) {
    await avisarJonathan(
      transporter,
      `\u26A0\uFE0F Reportes ExpertCell - incidencias (${HOY})`,
      `Se enviaron ${enviados} de ${coords.length} reportes.\n\nIncidencias:\n- ${incidencias.join('\n- ')}`
    );
  }

  console.log(`Listo. Enviados: ${enviados}. Incidencias: ${incidencias.length}.`);
}

main().catch((e) => {
  console.error('Fallo no controlado:', e);
  process.exit(1);
});
