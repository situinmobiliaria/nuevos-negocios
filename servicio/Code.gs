/**
 * Nuevos Negocios SITÛ: servicio de datos en Google Apps Script.
 * Valida la cuenta @situ.cl en cada llamada y guarda los datos en una planilla
 * y una carpeta de Drive del administrador.
 * Implementar como aplicación web: Ejecutar como "Yo", acceso "Cualquier usuario".
 */
var DOMINIO = 'situ.cl';
var ADMIN = 'cristian.valenzuela@situ.cl';
var CLIENT_ID = '138206776381-eep6qsc0nu0ceb2k6gjd6mgbk10d3c7g.apps.googleusercontent.com';
var TABS = ['proyectos', 'versiones', 'cambios', 'usuarios'];

/** Ejecutar una sola vez desde el editor. */
function setup() {
  var p = PropertiesService.getScriptProperties();
  if (p.getProperty('SHEET_ID')) { Logger.log('Ya estaba configurado.'); return; }
  var ss = SpreadsheetApp.create('Nuevos Negocios SITU - Datos');
  TABS.forEach(function (t, i) {
    var s = i === 0 ? ss.getSheets()[0].setName(t) : ss.insertSheet(t);
    s.getRange(1, 1, 1, 4).setValues([['id', 'json', 'fecha', 'usuario']]);
  });
  var f = DriveApp.createFolder('Nuevos Negocios SITU - Archivos');
  p.setProperties({ SHEET_ID: ss.getId(), FOLDER_ID: f.getId(), VER: '1' });
  Logger.log('Listo. Planilla: ' + ss.getUrl());
}

function doGet() { return out({ ok: true, servicio: 'Nuevos Negocios SITU' }); }

function doPost(e) {
  try {
    var q = JSON.parse(e.postData.contents);
    var email = verificar(q.token);
    if (!email) return out({ error: 'sesion' });
    if (email === 'dominio') return out({ error: 'dominio' });
    var f = OPS[q.op];
    if (!f) return out({ error: 'operacion' });
    return out(f(q, email));
  } catch (x) {
    return out({ error: 'servidor', detalle: String(x.message || x) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function verificar(t) {
  if (!t) return null;
  var key = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t));
  var c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return hit;
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(t), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  var p = JSON.parse(r.getContentText());
  var email = String(p.email || '').toLowerCase();
  var ver = p.email_verified === true || p.email_verified === 'true';
  if (p.aud !== CLIENT_ID) return null;
  if (!ver || p.hd !== DOMINIO || email.slice(-(DOMINIO.length + 1)) !== '@' + DOMINIO) return 'dominio';
  c.put(key, email, 300);
  return email;
}

function hoja(n) { return SpreadsheetApp.openById(prop('SHEET_ID')).getSheetByName(n); }
function prop(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function subirVersion() { PropertiesService.getScriptProperties().setProperty('VER', String(Number(prop('VER') || 0) + 1)); }
function conLock(fn) {
  var l = LockService.getScriptLock(); l.waitLock(20000);
  try { var r = fn(); subirVersion(); SpreadsheetApp.flush(); return r; } finally { l.releaseLock(); }
}
function filas(n) {
  var s = hoja(n), u = s.getLastRow();
  return u < 2 ? [] : s.getRange(2, 1, u - 1, 4).getValues();
}
function buscar(n, id) {
  var v = filas(n);
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return { fila: i + 2, doc: JSON.parse(v[i][1]) };
  return null;
}
function guardar(n, id, doc, email) {
  var r = buscar(n, id), f = [id, JSON.stringify(doc), new Date(), email];
  if (r) hoja(n).getRange(r.fila, 1, 1, 4).setValues([f]); else hoja(n).appendRow(f);
}
function exigeAdmin(email) { if (email !== ADMIN) throw new Error('Solo el administrador puede hacer este cambio.'); }

var OPS = {
  load: function (q, email) {
    if (!buscar('usuarios', email)) conLock(function () { guardar('usuarios', email, { nombre: q.nombre || email, email: email, rol: '' }, email); });
    var ver = prop('VER');
    if (q.ver && String(q.ver) === ver) return { ver: ver, igual: true };
    var o = { ver: ver, email: email, proyectos: {}, usuarios: {}, versiones: [], cambios: [] };
    filas('proyectos').forEach(function (r) { o.proyectos[r[0]] = JSON.parse(r[1]); });
    filas('usuarios').forEach(function (r) { o.usuarios[r[0]] = JSON.parse(r[1]); });
    filas('versiones').forEach(function (r) { var d = JSON.parse(r[1]); d.id = r[0]; o.versiones.push(d); });
    filas('cambios').forEach(function (r) { o.cambios.push(JSON.parse(r[1])); });
    if (q.inventario) {
      var id = prop('INV_ID');
      o.items = id ? JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString()).items : [];
    }
    return o;
  },
  patchProyecto: function (q, email) {
    return conLock(function () {
      var r = buscar('proyectos', q.id), d = r ? r.doc : {};
      Object.keys(q.patch || {}).forEach(function (k) { d[k] = q.patch[k]; });
      d.at = Date.now(); d.by = email;
      guardar('proyectos', q.id, d, email);
      return { ok: true };
    });
  },
  addVersion: function (q, email) {
    return conLock(function () {
      var d = q.doc || {}, id = Utilities.getUuid();
      d.by = email; d.at = Date.now();
      hoja('versiones').appendRow([id, JSON.stringify(d), new Date(), email]);
      return { id: id };
    });
  },
  // El control de cambios solo admite agregar: no existe operación para editar ni borrar.
  addCambio: function (q, email) {
    return conLock(function () {
      var d = q.doc || {}, id = Utilities.getUuid();
      d.by = email; d.at = Date.now();
      hoja('cambios').appendRow([id, JSON.stringify(d), new Date(), email]);
      return { id: id };
    });
  },
  setRol: function (q, email) {
    exigeAdmin(email);
    return conLock(function () {
      var r = buscar('usuarios', q.id); if (!r) throw new Error('Usuario no encontrado');
      r.doc.rol = q.rol; guardar('usuarios', q.id, r.doc, email); return { ok: true };
    });
  },
  setInventario: function (q, email) {
    exigeAdmin(email);
    return conLock(function () {
      var nom = 'inventario.json', c = JSON.stringify({ items: q.items }), id = prop('INV_ID');
      if (id) DriveApp.getFileById(id).setContent(c);
      else {
        var f = DriveApp.getFolderById(prop('FOLDER_ID')).createFile(nom, c, 'application/json');
        PropertiesService.getScriptProperties().setProperty('INV_ID', f.getId());
      }
      return { ok: true, n: q.items.length };
    });
  },
  subir: function (q, email) {
    var raiz = DriveApp.getFolderById(prop('FOLDER_ID')), nom = String(q.pid);
    var it = raiz.getFoldersByName(nom), dir = it.hasNext() ? it.next() : raiz.createFolder(nom);
    var f = dir.createFile(Utilities.newBlob(Utilities.base64Decode(q.b64), q.mime || 'application/octet-stream', q.name));
    f.setDescription('Subido por ' + email);
    f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: f.getUrl() };
  }
};
