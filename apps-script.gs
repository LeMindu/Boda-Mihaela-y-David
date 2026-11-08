/**
 * Recoge las confirmaciones de la web de boda en una hoja de Google Sheets
 * y (opcional) busca canciones en Spotify para el formulario.
 * Pegar en Extensiones > Apps Script de la hoja. Ver instrucciones.
 *
 * PESTAÑAS QUE CREA:
 *  - Invitados:  UNA FILA POR PERSONA (nombre, tipo, asiste, autobús, alergias...). Es la tabla principal.
 *  - Respuestas: una fila por formulario enviado (quién respondió, nº de personas, mensaje).
 *  - Canciones:  una fila por canción propuesta.
 *  - Resumen:    totales automáticos.
 */
const H_INV  = "Invitados";
const H_RESP = "Respuestas";
const H_CAN  = "Canciones";
const H_RES  = "Resumen";
const FORZAR = false; // pon true solo si quieres BORRAR todo y reconfigurar

const CAB_INV = ["Fecha respuesta","Invitado","Tipo","¿Asiste?","Respuesta enviada por","¿Autobús?",
  "¿Alergia o intolerancia?","Sin gluten / Celíaco","Sin lactosa","Vegetariano","Vegano","Frutos secos","Marisco","Otras alergias"];

/** Ejecuta esta función UNA vez para crear las pestañas, cabeceras y el resumen. */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hayDatos = ["Invitados","Respuestas","Asistentes"].some(n => {
    const h = ss.getSheetByName(n);
    return h && h.getLastRow() > 1;
  });
  if (hayDatos && !FORZAR) {
    throw new Error("Ya hay datos guardados. Para reconfigurar (y borrarlos) pon FORZAR = true.");
  }

  // --- Invitados (tabla principal) ---
  const inv = preparar_(ss, H_INV, CAB_INV);
  inv.getRange("A2:A").setNumberFormat("dd/mm/yyyy hh:mm");
  inv.setColumnWidth(2, 230); inv.setColumnWidth(5, 200); inv.setColumnWidth(14, 230);
  inv.setFrozenColumns(2);
  inv.getRange(1, 1, inv.getMaxRows(), CAB_INV.length).createFilter();
  inv.setConditionalFormatRules([
    regla_(inv, "G2:M", "#fde2e2"),   // alergias en rojo suave
    regla_(inv, "F2:F", "#dbeafe")    // autobús en azul suave
  ]);

  // --- Respuestas ---
  const r = preparar_(ss, H_RESP, ["Fecha","Nombre","¿Asiste?","Adultos","Niños","Total personas","¿Autobús?","Plazas autobús","Mensaje"]);
  r.getRange("A2:A").setNumberFormat("dd/mm/yyyy hh:mm");
  r.setColumnWidth(2, 230); r.setColumnWidth(9, 380);

  // --- Canciones ---
  const c = preparar_(ss, H_CAN, ["Fecha","Invitado","Título","Artista","Enlace Spotify"]);
  c.getRange("A2:A").setNumberFormat("dd/mm/yyyy hh:mm");
  c.setColumnWidth(2, 230); c.setColumnWidth(3, 260); c.setColumnWidth(4, 200);

  // --- Resumen (las fórmulas se guardan en inglés; Google Sheets las muestra en castellano) ---
  const s = ss.getSheetByName(H_RES) || ss.insertSheet(H_RES);
  s.clear();
  const filas = [
    ["Personas confirmadas",              '=COUNTIF(Invitados!D:D,"Sí")'],
    ["   Adultos",                        '=COUNTIFS(Invitados!D:D,"Sí",Invitados!C:C,"Adulto")'],
    ["   Niños",                          '=COUNTIFS(Invitados!D:D,"Sí",Invitados!C:C,"Niño")'],
    ["Respuestas «No puedo ir»",          '=COUNTIF(Respuestas!C:C,"No")'],
    ["Plazas de autobús",                 '=COUNTIFS(Invitados!D:D,"Sí",Invitados!F:F,"Sí")'],
    ["", ""],
    ["Con alguna alergia o intolerancia", '=COUNTIF(Invitados!G:G,"Sí")'],
    ["   Sin gluten / Celíaco",           '=COUNTIF(Invitados!H:H,"Sí")'],
    ["   Sin lactosa",                    '=COUNTIF(Invitados!I:I,"Sí")'],
    ["   Vegetariano",                    '=COUNTIF(Invitados!J:J,"Sí")'],
    ["   Vegano",                         '=COUNTIF(Invitados!K:K,"Sí")'],
    ["   Alergia a frutos secos",         '=COUNTIF(Invitados!L:L,"Sí")'],
    ["   Alergia a marisco",              '=COUNTIF(Invitados!M:M,"Sí")'],
    ["   Con otras alergias (ver Invitados)", '=COUNTIF(Invitados!N:N,"?*")'],
    ["", ""],
    ["Canciones propuestas",              '=COUNTA(Canciones!C2:C)']
  ];
  s.getRange(1, 1, filas.length, 1).setValues(filas.map(f => [f[0]]));
  s.getRange(1, 2, filas.length, 1).setFormulas(filas.map(f => [f[1]]));
  s.setColumnWidth(1, 330);
  [1, 4, 5, 7, 16].forEach(n => s.getRange(n, 1, 1, 2).setFontWeight("bold"));

  // Orden de pestañas y limpieza de hojas antiguas
  ss.setActiveSheet(inv); ss.moveActiveSheet(1);
  ["Hoja 1", "Sheet1", "Asistentes"].forEach(n => {
    const h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website) return ok_();                       // trampa antispam
    const quien = t_(d.nombre);
    if (!quien) return ok_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ahora = new Date();
    const asiste = d.asiste === "Sí" ? "Sí" : "No";
    const adultos = asiste === "Sí" ? num_(d.adultos, 1, 10) : 0;
    const ninos   = asiste === "Sí" ? num_(d.ninos, 0, 10) : 0;
    const bus     = asiste === "Sí" && d.bus === "Sí" ? "Sí" : "No";
    const total   = adultos + ninos;

    ss.getSheetByName(H_RESP).appendRow([
      ahora, quien, asiste, adultos, ninos, total, bus, bus === "Sí" ? total : 0, t_(d.mensaje)
    ]);

    // Una fila por persona
    const inv = ss.getSheetByName(H_INV);
    if (asiste === "Sí") {
      const lista = Array.isArray(d.asistentes) && d.asistentes.length ? d.asistentes.slice(0, 20) : [{ tipo: "Adulto", nombre: quien }];
      let nA = 0, nN = 0;
      lista.forEach(p => {
        const esNino = p.tipo === "Niño";
        esNino ? nN++ : nA++;
        const tipo = esNino ? "Niño" : "Adulto";
        const nombre = t_(p.nombre) || (tipo + " " + (esNino ? nN : nA) + " de " + quien);   // nunca queda vacío
        const marcas = [p.gluten, p.lactosa, p.vegetariano, p.vegano, p.frutos, p.marisco].map(si_);
        const otras = t_(p.otras);
        const alergia = (marcas.indexOf("Sí") >= 0 || otras) ? "Sí" : "No";
        inv.appendRow([ahora, nombre, tipo, "Sí", quien, bus, alergia].concat(marcas, [otras]));
      });
    } else {
      inv.appendRow([ahora, quien, "", "No", quien, "No", "", "", "", "", "", "", "", ""]);
    }

    if (asiste === "Sí" && Array.isArray(d.canciones)) {
      const sc = ss.getSheetByName(H_CAN);
      d.canciones.slice(0, 3).forEach(c => {
        const titulo = t_(c.titulo);
        if (!titulo) return;
        const enlace = /^https:\/\/(open\.spotify\.com|spotify\.link)\//.test(String(c.url || "")) ? String(c.url).slice(0, 300) : "";
        sc.appendRow([ahora, quien, titulo, t_(c.artista), enlace]);
      });
    }
    return ok_();
  } catch (err) {
    return ContentService.createTextOutput("error").setMimeType(ContentService.MimeType.TEXT);
  } finally {
    lock.releaseLock();
  }
}

/** GET sin parámetros: comprobación. GET ?q=texto: búsqueda de canciones en Spotify. */
function doGet(e) {
  const q = e && e.parameter && e.parameter.q ? String(e.parameter.q).trim().slice(0, 100) : "";
  if (!q) return ContentService.createTextOutput("Activo");
  return json_(buscarSpotify_(q));
}

/* ---- Spotify (credenciales en Configuración del proyecto > Propiedades del script:
        SPOTIFY_ID y SPOTIFY_SECRET). Sin ellas, la web usa el modo manual. ---- */
function tokenSpotify_() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get("sp_token");
  if (guardado) return guardado;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty("SPOTIFY_ID"), secreto = p.getProperty("SPOTIFY_SECRET");
  if (!id || !secreto) return null;
  const r = UrlFetchApp.fetch("https://accounts.spotify.com/api/token", {
    method: "post", payload: { grant_type: "client_credentials" },
    headers: { Authorization: "Basic " + Utilities.base64Encode(id + ":" + secreto) },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return null;
  const j = JSON.parse(r.getContentText());
  cache.put("sp_token", j.access_token, Math.max(60, (j.expires_in || 3600) - 120));
  return j.access_token;
}

function buscarSpotify_(q) {
  const cache = CacheService.getScriptCache();
  const clave = "q_" + Utilities.base64EncodeWebSafe(q.toLowerCase()).slice(0, 200);
  const hit = cache.get(clave);
  if (hit) return JSON.parse(hit);
  const token = tokenSpotify_();
  if (!token) return { ok: false, error: "sin_credenciales" };
  const r = UrlFetchApp.fetch("https://api.spotify.com/v1/search?type=track&limit=5&q=" + encodeURIComponent(q),
    { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { ok: false, error: "spotify_" + r.getResponseCode() };
  const items = (JSON.parse(r.getContentText()).tracks || {}).items || [];
  const res = { ok: true, items: items.map(t => ({
    id: t.id, titulo: t.name, artista: t.artists.map(a => a.name).join(", "),
    url: (t.external_urls || {}).spotify || "",
    img: ((t.album.images || []).slice(-1)[0] || {}).url || ""
  })).filter(t => t.url) };
  cache.put(clave, JSON.stringify(res), 21600);
  return res;
}

/* ---- utilidades ---- */
function preparar_(ss, nombre, cabecera) {
  const h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  const f = h.getFilter(); if (f) f.remove();
  h.clear(); h.clearConditionalFormatRules();
  h.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight("bold").setBackground("#e8eadf");
  h.setFrozenRows(1);
  return h;
}
function regla_(hoja, rango, color) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Sí").setBackground(color).setRanges([hoja.getRange(rango)]).build();
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ok_() { return ContentService.createTextOutput("ok").setMimeType(ContentService.MimeType.TEXT); }
function si_(v) { return v ? "Sí" : ""; }
function num_(v, min, max) { v = parseInt(v, 10); return isNaN(v) ? min : Math.min(max, Math.max(min, v)); }
// Limita longitud y evita que Sheets interprete el texto como fórmula
function t_(v) {
  v = String(v == null ? "" : v).trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
