/**
 * Recoge las confirmaciones de la web de boda en una hoja de Google Sheets
 * y (opcional) busca canciones en Spotify para el formulario.
 * Pegar en Extensiones > Apps Script de la hoja. Ver instrucciones.
 */
const H_RESP = "Respuestas";
const H_ASIS = "Asistentes";
const H_CAN  = "Canciones";
const H_RES  = "Resumen";
const FORZAR = false; // pon true solo si quieres BORRAR todo y reconfigurar

/** Ejecuta esta función UNA vez para crear las pestañas, cabeceras y el resumen. */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const previa = ss.getSheetByName(H_RESP);
  if (previa && previa.getLastRow() > 1 && !FORZAR) {
    throw new Error("Ya hay respuestas guardadas. Para reconfigurar (y borrarlas) pon FORZAR = true.");
  }

  let r = ss.getSheetByName(H_RESP) || ss.insertSheet(H_RESP);
  r.clear();
  r.appendRow(["Fecha","Nombre","¿Asiste?","Adultos","Niños","Total personas","¿Autobús?","Plazas autobús","Mensaje"]);
  r.setFrozenRows(1);

  let a = ss.getSheetByName(H_ASIS) || ss.insertSheet(H_ASIS);
  a.clear();
  a.appendRow(["Fecha","Respuesta de","Asistente","Tipo","Sin gluten / Celíaco","Sin lactosa","Vegetariano","Vegano","Frutos secos","Marisco","Otras alergias","¿Autobús?"]);
  a.setFrozenRows(1);

  let c = ss.getSheetByName(H_CAN) || ss.insertSheet(H_CAN);
  c.clear();
  c.appendRow(["Fecha","Invitado","Título","Artista","Enlace Spotify"]);
  c.setFrozenRows(1);

  let s = ss.getSheetByName(H_RES) || ss.insertSheet(H_RES);
  s.clear();
  const filas = [
    ["Respuestas «Sí»", '=COUNTIF(Respuestas!C:C,"Sí")'],
    ["Respuestas «No»", '=COUNTIF(Respuestas!C:C,"No")'],
    ["Adultos confirmados", '=SUMIF(Respuestas!C:C,"Sí",Respuestas!D:D)'],
    ["Niños confirmados", '=SUMIF(Respuestas!C:C,"Sí",Respuestas!E:E)'],
    ["Total personas", '=SUM(Respuestas!F:F)'],
    ["Plazas de autobús", '=SUM(Respuestas!H:H)'],
    ["", ""],
    ["Sin gluten / Celíaco", '=COUNTIF(Asistentes!E:E,"Sí")'],
    ["Sin lactosa", '=COUNTIF(Asistentes!F:F,"Sí")'],
    ["Vegetariano", '=COUNTIF(Asistentes!G:G,"Sí")'],
    ["Vegano", '=COUNTIF(Asistentes!H:H,"Sí")'],
    ["Alergia a frutos secos", '=COUNTIF(Asistentes!I:I,"Sí")'],
    ["Alergia a marisco", '=COUNTIF(Asistentes!J:J,"Sí")'],
    ["Con otras alergias (ver pestaña Asistentes)", '=COUNTIF(Asistentes!K:K,"?*")'],
    ["", ""],
    ["Canciones propuestas", '=COUNTA(Canciones!C2:C)']
  ];
  s.getRange(1, 1, filas.length, 2).setValues(filas);
  s.setColumnWidth(1, 320);
  s.getRange("A1:A16").setFontWeight("bold");

  const hojaInicial = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (hojaInicial && ss.getSheets().length > 1) ss.deleteSheet(hojaInicial);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website) return ok_();                       // trampa antispam
    if (!d.nombre) return ok_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ahora = new Date();
    const asiste = d.asiste === "Sí" ? "Sí" : "No";
    const adultos = asiste === "Sí" ? num_(d.adultos, 1, 10) : 0;
    const ninos   = asiste === "Sí" ? num_(d.ninos, 0, 10) : 0;
    const bus     = asiste === "Sí" && d.bus === "Sí" ? "Sí" : "No";
    const total   = adultos + ninos;

    ss.getSheetByName(H_RESP).appendRow([
      ahora, t_(d.nombre), asiste, adultos, ninos, total,
      bus, bus === "Sí" ? total : 0, t_(d.mensaje)
    ]);

    if (asiste === "Sí" && Array.isArray(d.asistentes)) {
      const sh = ss.getSheetByName(H_ASIS);
      d.asistentes.slice(0, 20).forEach(p => {
        sh.appendRow([
          ahora, t_(d.nombre), t_(p.nombre), p.tipo === "Niño" ? "Niño" : "Adulto",
          si_(p.gluten), si_(p.lactosa), si_(p.vegetariano), si_(p.vegano),
          si_(p.frutos), si_(p.marisco), t_(p.otras), bus
        ]);
      });
    }
    if (asiste === "Sí" && Array.isArray(d.canciones)) {
      const sc = ss.getSheetByName(H_CAN);
      d.canciones.slice(0, 3).forEach(c => {
        const titulo = t_(c.titulo);
        if (!titulo) return;
        const enlace = /^https:\/\/(open\.spotify\.com|spotify\.link)\//.test(String(c.url || "")) ? String(c.url).slice(0, 300) : "";
        sc.appendRow([ahora, t_(d.nombre), titulo, t_(c.artista), enlace]);
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
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ok_() { return ContentService.createTextOutput("ok").setMimeType(ContentService.MimeType.TEXT); }
function si_(v) { return v ? "Sí" : ""; }
function num_(v, min, max) { v = parseInt(v, 10); return isNaN(v) ? min : Math.min(max, Math.max(min, v)); }
// Limita longitud y evita que Sheets interprete el texto como fórmula
function t_(v) {
  v = String(v == null ? "" : v).slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
