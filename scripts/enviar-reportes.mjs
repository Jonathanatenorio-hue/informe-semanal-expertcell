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
const ANIO = parseInt(HOY_ARCHIVO.slice(0, 4), 10);
// getDay() sobre la fecha local de CDMX: 1 = lunes
const ES_LUNES = new Date(new Date().toLocaleString('en-US', { timeZone: TZ })).getDay() === 1;

// MODO PRUEBA: si existe el secret MODO_PRUEBA (cualquier valor), TODOS los correos
// se redirigen unicamente a Jonathan (CC_SIEMPRE). Los coordinadores NO reciben nada.
// Para volver al envio normal, basta con borrar ese secret.
const MODO_PRUEBA = !!process.env.MODO_PRUEBA;

// ENVIAR_SUPERVISORES: si existe el secret (cualquier valor), ADEMAS de los
// coordinadores se le manda a cada supervisor su propio detalle de equipo,
// usando el correo guardado en supervisores_config.email (Supabase).
// Si no existe el secret, solo se envia a coordinadores (comportamiento base).
const ENVIAR_SUPERVISORES = !!process.env.ENVIAR_SUPERVISORES;

// ---------- HELPERS ----------
function clasificarCentro(raw) {
  const c = (raw || '').toUpperCase();
  if (c.includes('CC2')) return 'CC2';
  if (c.includes('JV')) return 'JV';
  return null;
}

function construirCuerpo(nombreCoord, etiquetaCentro, hayInact, hayTarjeta, hayAnual) {
  const inactLinea = hayInact
    ? 'Excel de inactividad del centro (agentes con dias sin activar / sin programar).'
    : 'Hoy no hubo agentes en alerta de inactividad. \u2705';
  const anualLinea = hayAnual
    ? '\n- Análisis Anual de cada supervisor del centro (adjunto, envío de los lunes).'
    : '';
  const text =
`Hola ${nombreCoord},

Adjunto el reporte operativo de hoy (${HOY}) para ${etiquetaCentro}:

- Excel del centro: dashboard del coordinador, una hoja de detalle por cada supervisor y el detalle general del centro.
- Tabla maestra del centro (Excel).
- ${inactLinea}${anualLinea}

Saludos,
Reportes ExpertCell (envio automatico)`;

  const imgTarjeta = hayTarjeta
    ? `<p style="margin:6px 0 16px;"><img src="cid:tarjeta-resumen" alt="Resumen del centro" style="max-width:100%;border-radius:10px;"></p>`
    : '';
  const anualHtml = hayAnual
    ? `<li><strong>Análisis Anual</strong> de cada supervisor del centro (adjunto, env&iacute;o de los lunes).</li>`
    : '';
  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola ${nombreCoord},</p>
  <p>Resumen del centro <strong>${etiquetaCentro}</strong> al ${HOY}:</p>
  ${imgTarjeta}
  <p>Adjuntos:</p>
  <ul>
    <li><strong>Excel del centro:</strong> dashboard del coordinador, una hoja de detalle por cada supervisor y el detalle general del centro.</li>
    <li><strong>Tabla maestra</strong> del centro (Excel adjunto).</li>
    <li>${inactLinea}</li>
    ${anualHtml}
  </ul>
  <p style="color:#6b7280;font-size:12px;margin-top:18px">Reportes ExpertCell &middot; env&iacute;o autom&aacute;tico</p>
</div>`;
  return { text, html };
}

function construirCuerpoSupervisor(nombreSup, hayAnual) {
  const anualT = hayAnual ? '\n- Tu Análisis Anual (envío de los lunes).' : '';
  const anualH = hayAnual ? '<li>Tu <strong>Análisis Anual</strong> (env&iacute;o de los lunes).</li>' : '';
  const text =
`Hola ${nombreSup},

Adjunto tu reporte de equipo de hoy (${HOY}): dashboard de tu equipo y el detalle
de cada uno de tus agentes (avance, meta, % de cumplimiento e inactividad).${anualT}

Saludos,
Reportes ExpertCell (envio automatico)`;
  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola ${nombreSup},</p>
  <p>Adjunto tu reporte de equipo de hoy (<strong>${HOY}</strong>):</p>
  <ul>
    <li>Dashboard de tu equipo y el detalle de cada uno de tus agentes (avance, meta, % de cumplimiento e inactividad).</li>
    ${anualH}
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

      // 4) Imagen de la tarjeta principal del centro (para el cuerpo del correo)
      let pngTarjeta = null;
      try {
        const elTarjeta = page.locator('#avance-tarjeta-principal');
        if ((await elTarjeta.count()) > 0) {
          await elTarjeta.first().scrollIntoViewIfNeeded().catch(() => {});
          pngTarjeta = await elTarjeta.first().screenshot();
        }
      } catch (e) { /* la tarjeta es opcional */ }

      // Lista de supervisores del centro con datos (para anual de lunes y envio a supervisores)
      const centroSups = await page.evaluate(() => {
        const d = window.__detalleAgentesData;
        if (!d || !d.agentesPorSup) return [];
        return Object.keys(d.agentesPorSup).filter((s) => (d.agentesPorSup[s] || []).length > 0);
      });

      // 5) LUNES: Analisis Anual de cada supervisor del centro (para el coordinador)
      const anualesCoord = [];
      if (ES_LUNES) {
        for (const sup of centroSups) {
          await page.evaluate(() => { window.__captured = []; });
          await page.evaluate((s, a) => { try { exportarAnalisisAnualExcel(s, a); } catch (e) {} }, sup, ANIO);
          const cap = await page.evaluate(() => window.__captured.slice());
          const arch = cap.find((x) => x.b64) || null;
          if (arch) anualesCoord.push({ filename: `Anual_${sup}_${ANIO}.xlsx`, content: Buffer.from(arch.b64, 'base64') });
        }
      }

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
      anualesCoord.forEach((a) => attachments.push(a));
      // Imagen de la tarjeta incrustada en el cuerpo
      if (pngTarjeta) {
        attachments.push({ filename: 'resumen.png', content: pngTarjeta, cid: 'tarjeta-resumen' });
      }

      const cuerpo = construirCuerpo(nombre, ruta.etiqueta, !!archInact, !!pngTarjeta, ES_LUNES && anualesCoord.length > 0);

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

      // ---- Envio a cada SUPERVISOR de este centro (opcional) ----
      if (ENVIAR_SUPERVISORES) {
        // El filtro del centro ya esta aplicado, asi que agentesPorSup trae solo
        // los supervisores de este centro. Leemos sus nombres + email de Supabase.
        const sups = await page.evaluate(() => {
          const d = window.__detalleAgentesData;
          const emails = window.__emailsSupervisores || {};
          if (!d || !d.agentesPorSup) return [];
          return Object.keys(d.agentesPorSup)
            .filter((s) => (d.agentesPorSup[s] || []).length > 0)
            .map((s) => ({ sup: s, email: emails[s] || null }));
        });

        for (const s of sups) {
          if (!s.email && !MODO_PRUEBA) {
            incidencias.push(`Supervisor "${s.sup}" (${ruta.etiqueta}) sin email en supervisores_config; no se le envio su detalle.`);
            continue;
          }
          try {
            await page.evaluate(() => { window.__captured = []; });
            await page.evaluate((nom) => { descargarEquipoExcel(nom); }, s.sup);
            const capEq = await page.evaluate(() => window.__captured.slice());
            const archEq = capEq.find((x) => x.b64) || null;
            if (!archEq) {
              incidencias.push(`No se genero el Excel de "${s.sup}" (${ruta.etiqueta}); no se le envio.`);
              continue;
            }

            const supTo = MODO_PRUEBA ? CC_SIEMPRE : s.email;
            const supCc = MODO_PRUEBA ? undefined : CC_SIEMPRE;
            const supPref = MODO_PRUEBA ? '[PRUEBA] ' : '';

            const supAttachments = [
              { filename: `Equipo_${s.sup}_${HOY_ARCHIVO}.xlsx`, content: Buffer.from(archEq.b64, 'base64') },
            ];
            // LUNES: anexar su Analisis Anual
            let supHayAnual = false;
            if (ES_LUNES) {
              await page.evaluate(() => { window.__captured = []; });
              await page.evaluate((nom, a) => { try { exportarAnalisisAnualExcel(nom, a); } catch (e) {} }, s.sup, ANIO);
              const capAn = await page.evaluate(() => window.__captured.slice());
              const archAn = capAn.find((x) => x.b64) || null;
              if (archAn) {
                supAttachments.push({ filename: `Anual_${s.sup}_${ANIO}.xlsx`, content: Buffer.from(archAn.b64, 'base64') });
                supHayAnual = true;
              }
            }
            const cuerpoSup = construirCuerpoSupervisor(s.sup, supHayAnual);

            await transporter.sendMail({
              from: REMITENTE,
              to: supTo,
              cc: supCc,
              subject: `${supPref}Reporte diario ExpertCell \u00B7 Equipo ${s.sup} \u00B7 ${HOY}`,
              text: cuerpoSup.text,
              html: cuerpoSup.html,
              attachments: supAttachments,
            });
            enviados++;
            console.log(`Enviado a ${supTo} (equipo ${s.sup})${MODO_PRUEBA ? ' [PRUEBA]' : ''}.`);
          } catch (e) {
            incidencias.push(`Error con supervisor "${s.sup}" (${ruta.etiqueta}): ${e.message}.`);
          }
        }
      }
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
