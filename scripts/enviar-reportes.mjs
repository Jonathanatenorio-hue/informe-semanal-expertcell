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

// Talento Humano por centro: reciben SOLO imagenes (tarjeta + avance diario)
// para pasar en las pantallas. Sin Excel, sin tops.
const TALENTO_HUMANO_POR_CENTRO = {
  CC2: 'reclutamientocc2@solucell.com.mx',
  JV:  'k.flores@solucell.com.mx',
};

const GMAIL_USER = process.env.GMAIL_USER;
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;
const REMITENTE = `"Jonathan Atenorio" <${GMAIL_USER}>`;

const TZ = 'America/Mexico_City';
// HOY/HOY_ARCHIVO arrancan con la fecha de envio, pero despues de cargar el
// reporte se reasignan a la FECHA DE CORTE real de los datos (normalmente ayer,
// porque hoy aun no hay datos). Por eso son `let`, no `const`.
let HOY = new Date().toLocaleDateString('es-MX', {
  timeZone: TZ, year: 'numeric', month: 'long', day: 'numeric',
});
let HOY_ARCHIVO = new Date().toLocaleDateString('en-CA', { timeZone: TZ }); // YYYY-MM-DD
const ANIO = parseInt(HOY_ARCHIVO.slice(0, 4), 10);
// getDay() sobre la fecha local de CDMX: 1 = lunes
const ES_LUNES = new Date(new Date().toLocaleString('en-US', { timeZone: TZ })).getDay() === 1;

// Convierte 'YYYY-MM-DD' a "9 de junio de 2026"
function fmtFechaEs(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  if (!m) return null;
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return `${parseInt(m[3], 10)} de ${meses[parseInt(m[2], 10) - 1]} de ${m[1]}`;
}

// Devuelve el ULTIMO DIA HABIL esperado en los datos: normalmente ayer; si ayer
// fue domingo (sin operacion), regresa el sabado. Formato 'YYYY-MM-DD' en CDMX.
// Sirve para detectar si el reporte trae datos frescos o si se olvido subirlos.
function ultimoDiaHabilEsperado() {
  const hoyCdmx = new Date(new Date().toLocaleString('en-US', { timeZone: TZ }));
  const d = new Date(hoyCdmx);
  d.setDate(d.getDate() - 1);          // ayer
  if (d.getDay() === 0) d.setDate(d.getDate() - 1); // si es domingo, al sabado
  const y = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${mm}-${dd}`;
}

// ============================================================
//  SEGURIDAD (doble candado):
//  1) Por DEFECTO, TODO va UNICAMENTE a Jonathan (modo prueba). Para que llegue
//     a los destinatarios REALES hay que crear A PROPOSITO el secret ENVIAR_EN_VIVO.
//  2) FRENO DE EMERGENCIA: si existe el secret MODO_PRUEBA, SIEMPRE se queda en
//     prueba, aunque ENVIAR_EN_VIVO este puesto. Sirve para frenar al instante.
//  Resultado: solo llega a coordinadores/supervisores si ENVIAR_EN_VIVO existe
//  Y MODO_PRUEBA no existe. Cualquier olvido cae del lado seguro.
// ============================================================
const ENVIAR_EN_VIVO = !!process.env.ENVIAR_EN_VIVO;
const FRENO_PRUEBA = !!process.env.MODO_PRUEBA;
const ES_PRUEBA = FRENO_PRUEBA || !ENVIAR_EN_VIVO;

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

function resumenHtml(resumen) {
  if (!resumen) return '';
  const md = (resumen.modelos || []);
  const pl = (resumen.planes || []);
  if (!md.length && !pl.length) return '';
  const dineroMx = (n) => '$' + Number(n || 0).toLocaleString('es-MX');
  let h = '<div style="background:#F5F8FB;border-radius:10px;padding:12px 16px;margin:8px 0 16px;">';
  if (md.length) {
    h += '<p style="margin:0 0 6px;"><strong>Modelos que más se venden este mes:</strong></p><ol style="margin:0 0 10px;padding-left:20px;">';
    md.forEach((m) => { h += `<li>${m.modelo} — <strong>${m.count}</strong></li>`; });
    h += '</ol>';
  }
  if (pl.length) {
    h += '<p style="margin:0 0 6px;"><strong>Planes que más se venden:</strong></p><ol style="margin:0;padding-left:20px;">';
    pl.forEach((p) => { h += `<li>${p.plan} — <strong>${p.count}</strong> ${p.arpu ? `(ARPU ${dineroMx(p.arpu)})` : ''}</li>`; });
    h += '</ol>';
  }
  h += '</div>';
  return h;
}

function resumenText(resumen) {
  if (!resumen) return '';
  const md = (resumen.modelos || []);
  const pl = (resumen.planes || []);
  if (!md.length && !pl.length) return '';
  let t = '';
  if (md.length) t += '\nModelos que más se venden este mes: ' + md.map((m) => `${m.modelo} (${m.count})`).join(', ') + '.';
  if (pl.length) t += '\nPlanes que más se venden: ' + pl.map((p) => `${p.plan} (${p.count}${p.arpu ? `, ARPU $${p.arpu}` : ''})`).join(', ') + '.';
  return t;
}

function construirCuerpo(nombreCoord, etiquetaCentro, hayInact, hayTarjeta, hayAnual, resumen, hayVista) {
  const inactLinea = hayInact
    ? 'Excel de inactividad del centro (agentes con dias sin activar / sin programar).'
    : 'Hoy no hubo agentes en alerta de inactividad. \u2705';
  const anualLinea = hayAnual
    ? '\n- El Análisis Anual de cada uno de tus supervisores (te lo mando los lunes).'
    : '';
  const text =
`Hola ${nombreCoord},

Te comparto el resumen de ${etiquetaCentro} al ${HOY}.${resumenText(resumen)}

Te dejo adjunto:
- El Excel del centro: dashboard, una hoja de detalle por cada supervisor y el detalle general.
- La tabla maestra del centro.
- ${inactLinea}${anualLinea}

Revísalo y cualquier cosa que veas me dices.

Saludos,
Jonathan`;

  const imgTarjeta = hayTarjeta
    ? `<p style="margin:6px 0 16px;"><img src="cid:tarjeta-resumen" alt="Resumen del centro" style="max-width:100%;border-radius:10px;"></p>`
    : '';
  const imgVista = hayVista
    ? `<p style="margin:14px 0 6px;"><strong>Venta diaria del mes:</strong></p><p style="margin:0 0 16px;"><img src="cid:vista-diaria" alt="Vista diaria del centro" style="max-width:100%;border-radius:10px;border:1px solid #E5E7EB;"></p>`
    : '';
  const anualHtml = hayAnual
    ? `<li>El <strong>Análisis Anual</strong> de cada uno de tus supervisores (te lo mando los lunes).</li>`
    : '';
  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola ${nombreCoord},</p>
  <p>Te comparto el resumen de <strong>${etiquetaCentro}</strong> al ${HOY}.</p>
  ${imgTarjeta}
  ${resumenHtml(resumen)}
  ${imgVista}
  <p>Te dejo adjunto:</p>
  <ul>
    <li>El <strong>Excel del centro</strong>: dashboard, una hoja de detalle por cada supervisor y el detalle general.</li>
    <li>La <strong>tabla maestra</strong> del centro.</li>
    <li>${inactLinea}</li>
    ${anualHtml}
  </ul>
  <p>Revísalo y cualquier cosa que veas me dices.</p>
  <p style="margin-top:16px;">Saludos,<br>Jonathan</p>
</div>`;
  return { text, html };
}

function construirCuerpoSupervisor(nombreSup, hayAnual, resumen, hayVista, hayInact) {
  const anualT = hayAnual ? '\nTambién te mando tu Análisis Anual (te lo paso los lunes).' : '';
  const anualH = hayAnual ? '<li>Tu <strong>Análisis Anual</strong> (te lo paso los lunes).</li>' : '';
  const inactT = hayInact
    ? '\n- Tu Excel de inactividad (agentes de tu equipo con días sin activar / sin programar).'
    : '\n- Hoy tu equipo no tiene agentes en alerta de inactividad. \u2705';
  const inactH = hayInact
    ? '<li>Tu <strong>Excel de inactividad</strong> (agentes con días sin activar / sin programar).</li>'
    : '<li>Hoy tu equipo no tiene agentes en alerta de inactividad. \u2705</li>';
  const text =
`Hola ${nombreSup},

Te paso tu reporte de equipo de hoy.${resumenText(resumen)}

Te dejo adjunto:
- El Excel de tu equipo: dashboard y detalle de cada uno de tus agentes (avance, meta, % de cumplimiento e inactividad).
- Tu tabla maestra.${inactT}${anualT}

Échale un ojo y cualquier duda me buscas.

Saludos,
Jonathan`;
  const imgVista = hayVista
    ? `<p style="margin:14px 0 6px;"><strong>Tu venta diaria del mes:</strong></p><p style="margin:0 0 16px;"><img src="cid:vista-diaria" alt="Vista diaria del equipo" style="max-width:100%;border-radius:10px;border:1px solid #E5E7EB;"></p>`
    : '';
  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola ${nombreSup},</p>
  <p>Te paso tu reporte de equipo de hoy.</p>
  ${resumenHtml(resumen)}
  ${imgVista}
  <p>Te dejo adjunto:</p>
  <ul>
    <li>El <strong>Excel de tu equipo</strong>: dashboard y detalle de cada uno de tus agentes (avance, meta, % de cumplimiento e inactividad).</li>
    <li>Tu <strong>tabla maestra</strong>.</li>
    ${inactH}
    ${anualH}
  </ul>
  <p>Échale un ojo y cualquier duda me buscas.</p>
  <p style="margin-top:16px;">Saludos,<br>Jonathan</p>
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
// Cuerpo para Talento Humano: solo imagenes grandes (tarjeta + avance diario)
// para proyectar en pantalla. Sin adjuntos ni tablas.
function construirCuerpoTalento(etiquetaCentro, hayTarjeta, hayVista) {
  const text =
`Hola,

Les comparto las imágenes de ${etiquetaCentro} al ${HOY} para las pantallas.

Saludos,
Jonathan`;
  const imgTarjeta = hayTarjeta
    ? `<p style="margin:0 0 18px;"><img src="cid:tarjeta-resumen" alt="Avance de ${etiquetaCentro}" style="max-width:100%;border-radius:10px;"></p>`
    : '';
  const imgVista = hayVista
    ? `<p style="margin:0 0 8px;"><strong>Venta diaria del mes</strong></p><p style="margin:0 0 8px;"><img src="cid:vista-diaria" alt="Avance diario de ${etiquetaCentro}" style="max-width:100%;border-radius:10px;border:1px solid #E5E7EB;"></p>`
    : '';
  const html =
`<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.55">
  <p>Hola,</p>
  <p>Les comparto las imágenes de <strong>${etiquetaCentro}</strong> al ${HOY} para las pantallas.</p>
  ${imgTarjeta}
  ${imgVista}
  <p style="margin-top:16px;">Saludos,<br>Jonathan</p>
</div>`;
  return { text, html };
}

// Cuerpo para Talento Humano: solo imagenes grandes (tarjeta + avance diario)
async function capturarVistaDiaria(page) {
  // Captura la vista diaria CON las tarjetas de arriba (las que explican como leer
  // las lineas) + la grafica. Oculta temporalmente el titulo y los controles de la
  // seccion, fotografia el cuerpo, y restaura todo.
  try {
    await page.evaluate(() => {
      const sb = document.querySelector('#sec-vista-tiempo .section-body');
      if (!sb) return;
      const cards = document.getElementById('vt-resumen-cards');
      const chartWrap = document.querySelector('#sec-vista-tiempo .chart-wrap');
      window.__vtHidden = [];
      Array.from(sb.children).forEach((ch) => {
        if (ch !== cards && ch !== chartWrap) {
          window.__vtHidden.push([ch, ch.style.display]);
          ch.style.display = 'none';
        }
      });
    });
    let png = null;
    const el = page.locator('#sec-vista-tiempo .section-body');
    if ((await el.count()) > 0) {
      await el.first().scrollIntoViewIfNeeded().catch(() => {});
      await page.waitForTimeout(150);
      png = await el.first().screenshot();
    }
    await page.evaluate(() => {
      (window.__vtHidden || []).forEach(([ch, disp]) => { ch.style.display = disp; });
      window.__vtHidden = [];
    });
    return png;
  } catch (e) { return null; }
}

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

  // --- Usar la FECHA DE CORTE real del reporte (normalmente ayer) en vez de hoy ---
  try {
    const corte = await page.evaluate(() => {
      const a = window.__avanceData;
      return a && a.fechaCorteStr ? a.fechaCorteStr : null;
    });
    const corteTxt = fmtFechaEs(corte);
    if (corte && corteTxt) {
      HOY_ARCHIVO = corte;
      HOY = corteTxt;
      console.log(`Fecha de corte del reporte: ${HOY} (${HOY_ARCHIVO}).`);
    }

    // --- REGLA: no enviar si los datos no estan actualizados ---
    // Si la fecha de corte del reporte es ANTERIOR al ultimo dia habil esperado,
    // quiere decir que no se subieron los Excel nuevos. En ese caso NO se envia
    // nada a nadie y solo se avisa a Jonathan.
    const esperado = ultimoDiaHabilEsperado();
    if (corte && corte < esperado) {
      await browser.close();
      await avisarJonathan(
        transporter,
        `\u26A0\uFE0F Reportes ExpertCell - datos sin actualizar (${HOY})`,
        `No se envio ningun reporte porque los datos no estan actualizados.\n\n` +
        `La fecha de corte del reporte sigue en ${corteTxt} (${corte}), y se esperaba ` +
        `al menos ${fmtFechaEs(esperado)} (${esperado}).\n\n` +
        `Probablemente falto subir los Excel a Supabase. Subelos y vuelve a correr el envio ` +
        `(Actions > Run workflow, o espera al disparo de manana).`
      );
      console.log(`Datos sin actualizar (corte ${corte} < esperado ${esperado}). No se envio nada.`);
      process.exit(0);
    }
  } catch (e) { /* si falla, se queda con la fecha de hoy */ }

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

      // 4b) Imagen de la vista diaria del centro (tarjetas + grafica). Respeta el
      //     filtro de centro = vista Global.
      let pngVista = null;
      try {
        await page.evaluate(() => {
          const sec = document.getElementById('sec-vista-tiempo');
          if (sec) sec.classList.add('abierto');
          const btnGlobal = document.querySelector('.vt-view-btn[data-vt-view="total"]');
          if (btnGlobal) btnGlobal.click(); // vista Global (respeta el centro del filtro de arriba)
        });
        await page.waitForTimeout(1600);
        pngVista = await capturarVistaDiaria(page);
      } catch (e) { /* vista diaria opcional */ }

      // Lista de supervisores del centro con datos (para anual de lunes y envio a supervisores)
      const centroSups = await page.evaluate(() => {
        const d = window.__detalleAgentesData;
        if (!d || !d.agentesPorSup) return [];
        return Object.keys(d.agentesPorSup).filter((s) => (d.agentesPorSup[s] || []).length > 0);
      });

      // Resumen de modelos y planes que mas venden en el centro (para el cuerpo)
      const resumenCentro = await page.evaluate(async () => {
        try {
          const v = window.__resumenVentas() || {};
          const modelos = await window.__resumenModelos();
          return { planes: v.planes || [], modelos: modelos || [] };
        } catch (e) { return null; }
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
      // Imagenes incrustadas en el cuerpo
      if (pngTarjeta) {
        attachments.push({ filename: 'resumen.png', content: pngTarjeta, cid: 'tarjeta-resumen' });
      }
      if (pngVista) {
        attachments.push({ filename: 'vista-diaria.png', content: pngVista, cid: 'vista-diaria' });
      }

      const cuerpo = construirCuerpo(nombre, ruta.etiqueta, !!archInact, !!pngTarjeta, ES_LUNES && anualesCoord.length > 0, resumenCentro, !!pngVista);

      // En MODO PRUEBA todo va solo a Jonathan; en normal, al coordinador con copia.
      const destinoTo = ES_PRUEBA ? CC_SIEMPRE : ruta.para;
      const destinoCc = ES_PRUEBA ? undefined : CC_SIEMPRE;
      const prefijoAsunto = ES_PRUEBA ? '[PRUEBA] ' : '';

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
      console.log(`Enviado a ${destinoTo} (${ruta.etiqueta})${ES_PRUEBA ? ' [PRUEBA]' : ''}.`);

      // ---- Envio a TALENTO HUMANO del centro: solo imagenes para pantallas ----
      const thEmail = TALENTO_HUMANO_POR_CENTRO[centroKey];
      if (thEmail && (pngTarjeta || pngVista)) {
        try {
          const thAdjuntos = [];
          if (pngTarjeta) thAdjuntos.push({ filename: 'resumen.png', content: pngTarjeta, cid: 'tarjeta-resumen' });
          if (pngVista) thAdjuntos.push({ filename: 'vista-diaria.png', content: pngVista, cid: 'vista-diaria' });
          const cuerpoTh = construirCuerpoTalento(ruta.etiqueta, !!pngTarjeta, !!pngVista);
          const thTo = ES_PRUEBA ? CC_SIEMPRE : thEmail;
          const thCc = ES_PRUEBA ? undefined : CC_SIEMPRE;
          await transporter.sendMail({
            from: REMITENTE,
            to: thTo,
            cc: thCc,
            subject: `${prefijoAsunto}Avance ExpertCell \u00B7 ${ruta.etiqueta} \u00B7 ${HOY}`,
            text: cuerpoTh.text,
            html: cuerpoTh.html,
            attachments: thAdjuntos,
          });
          enviados++;
          console.log(`Enviado a ${thTo} (Talento Humano ${ruta.etiqueta})${ES_PRUEBA ? ' [PRUEBA]' : ''}.`);
        } catch (e) {
          incidencias.push(`Error enviando a Talento Humano (${ruta.etiqueta}): ${e.message}.`);
        }
      } else if (!thEmail) {
        incidencias.push(`No hay correo de Talento Humano configurado para ${ruta.etiqueta}; no se le envio.`);
      }

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
          if (!s.email && !ES_PRUEBA) {
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

            const supTo = ES_PRUEBA ? CC_SIEMPRE : s.email;
            const supCc = ES_PRUEBA ? undefined : CC_SIEMPRE;
            const supPref = ES_PRUEBA ? '[PRUEBA] ' : '';

            const supAttachments = [
              { filename: `Equipo_${s.sup}_${HOY_ARCHIVO}.xlsx`, content: Buffer.from(archEq.b64, 'base64') },
            ];

            // Tabla maestra del supervisor (filtrada a su equipo)
            await page.evaluate(() => { window.__captured = []; });
            await page.evaluate((nom) => { try { descargarMaestraExcel(nom); } catch (e) {} }, s.sup);
            const capMaeS = await page.evaluate(() => window.__captured.slice());
            const archMaeS = capMaeS.find((x) => x.b64) || null;
            if (archMaeS) {
              supAttachments.push({ filename: `Tabla_Maestra_${s.sup}_${HOY_ARCHIVO}.xlsx`, content: Buffer.from(archMaeS.b64, 'base64') });
            }

            // Inactividad del supervisor (filtrada a su equipo). Opcional: si su equipo
            // no tiene agentes en alerta, no se genera y no se adjunta.
            await page.evaluate(() => { window.__captured = []; });
            await page.evaluate((nom) => { try { descargarInactividadExcel(nom); } catch (e) {} }, s.sup);
            const capInaS = await page.evaluate(() => window.__captured.slice());
            const archInaS = capInaS.find((x) => x.b64) || null;
            if (archInaS) {
              supAttachments.push({ filename: `Inactividad_${s.sup}_${HOY_ARCHIVO}.xlsx`, content: Buffer.from(archInaS.b64, 'base64') });
            }
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
            // Vista diaria del supervisor (vista "Por supervisor" + su selector), con tarjetas
            let pngVistaSup = null;
            try {
              await page.evaluate((nom) => {
                const sec = document.getElementById('sec-vista-tiempo');
                if (sec) sec.classList.add('abierto');
                const btnSup = document.querySelector('.vt-view-btn[data-vt-view="supervisor"]');
                if (btnSup) btnSup.click();
                const sel = document.getElementById('vt-filtro-extra');
                if (sel) { sel.value = nom; sel.dispatchEvent(new Event('change', { bubbles: true })); }
              }, s.sup);
              await page.waitForTimeout(1600);
              pngVistaSup = await capturarVistaDiaria(page);
            } catch (e) { /* vista opcional */ }
            if (pngVistaSup) {
              supAttachments.push({ filename: 'vista-diaria.png', content: pngVistaSup, cid: 'vista-diaria' });
            }

            // Resumen de modelos y planes del equipo del supervisor
            const resumenSup = await page.evaluate(async (nom) => {
              try {
                const v = window.__resumenVentas(nom) || {};
                const modelos = await window.__resumenModelos(nom);
                return { planes: v.planes || [], modelos: modelos || [] };
              } catch (e) { return null; }
            }, s.sup);

            const cuerpoSup = construirCuerpoSupervisor(s.sup, supHayAnual, resumenSup, !!pngVistaSup, !!archInaS);

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
            console.log(`Enviado a ${supTo} (equipo ${s.sup})${ES_PRUEBA ? ' [PRUEBA]' : ''}.`);
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
