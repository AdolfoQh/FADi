// ============================================================
//  PRESENCIA — Sistema de Asistencia QR
//  FAD UNCuyo — Multi-docente
//  ARCHIVO: Codigo.gs
// ============================================================

var SHEET_ID = '14Z0IotObyapJh9qW3unayNOjt7gziVI11jaTqV62LjQ';

var TAB_DOCENTES    = 'Docentes';
var TAB_MATERIAS    = 'Materias';
var TAB_REGISTROS   = 'Registros';
var TAB_SESIONES    = 'Sesiones';
var TAB_RESUMEN     = 'Resumen';
var TAB_DISPOSITIVOS  = 'Dispositivos';
var TAB_GRUPOS        = 'Grupos';
var TAB_SESIONES_EVAL = 'SesionesEval';
var TAB_EVALUACIONES  = 'Evaluaciones';

var NOTA_APROBACION = 60;
var NOTA_PROMOCION  = 80;
var TIPOS_TP        = ['tp1','tp2','tp3','tp4'];
var TIPOS_PARCIAL   = ['parcial1','parcial2'];

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  var params = e.parameter;
  if (params.vista === 'alumno') {
    var html = HtmlService.createTemplateFromFile('Alumno');
    html.token = params.token || '';
    return html.evaluate()
      .setTitle('Marcar Asistencia — PRESENCIA')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Docente')
    .setTitle('PRESENCIA — Panel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var result;

  if      (action === 'login')             result = login(data);
  else if (action === 'getDocentes')       result = getDocentes(data);
  else if (action === 'saveDocente')       result = saveDocente(data);
  else if (action === 'deleteDocente')     result = deleteDocente(data);
  else if (action === 'getMaterias')       result = getMaterias(data);
  else if (action === 'saveMateria')       result = saveMateria(data);
  else if (action === 'editarMateria')     result = editarMateria(data);
  else if (action === 'deleteMateria')     result = deleteMateria(data);
  else if (action === 'getAlumnos')        result = getAlumnos(data.urlAlumnos || '');
  else if (action === 'activarSesion')     result = activarSesion(data);
  else if (action === 'getSesionActiva')   result = getSesionActiva(data);
  else if (action === 'getSesionPorToken') result = getSesionPorToken(data.token);
  else if (action === 'extenderSesion')   result = extenderSesion(data);
  else if (action === 'cerrarSesion')      result = cerrarSesion(data);
  else if (action === 'expirarSesion')     result = expirarSesion(data);
  else if (action === 'marcarPresente')    result = marcarPresente(data);
  else if (action === 'vincularDispositivo') result = vincularDispositivo(data);
  else if (action === 'buscarDispositivo')   result = buscarDispositivo(data);
  else if (action === 'getDashboard')        result = getDashboard(data);
  else if (action === 'toggleManual')      result = toggleManual(data);
  else if (action === 'getRegistrosDia')   result = getRegistrosDia(data);
  else if (action === 'getGrupos')         result = getGrupos(data);
  else if (action === 'saveGrupo')         result = saveGrupo(data);
  else if (action === 'editarGrupo')       result = editarGrupo(data);
  else if (action === 'deleteGrupo')       result = deleteGrupo(data);
  else if (action === 'abrirSesionEval')   result = abrirSesionEval(data);
  else if (action === 'cerrarSesionEval')  result = cerrarSesionEval(data);
  else if (action === 'getSesionesEval')   result = getSesionesEval(data);
  else if (action === 'saveEvaluacion')    result = saveEvaluacion(data);
  else if (action === 'getEvaluaciones')   result = getEvaluaciones(data);
  else if (action === 'getHistorialAlumno') result = getHistorialAlumno(data);
  else if (action === 'getResumenAlumno')  result = getResumenAlumno(data);
  else if (action === 'getResumenGrupo')   result = getResumenGrupo(data);
  else result = { ok: false, error: 'Accion desconocida: ' + action };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TEST ─────────────────────────────────────────────────────
function testAcceso() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Sheet OK: ' + ss.getName());
    var sheets = ss.getSheets().map(function(s){ return s.getName(); });
    Logger.log('Pestanas: ' + sheets.join(', '));
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

// ── SETUP ────────────────────────────────────────────────────
function setupSheets() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Abriendo sheet: ' + ss.getName());

    function crearSiNoExiste(nombre, headers) {
      var sh = ss.getSheetByName(nombre);
      if (!sh) {
        sh = ss.insertSheet(nombre);
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2357e8').setFontColor('#ffffff');
        sh.setFrozenRows(1);
        Logger.log('Creada pestaña: ' + nombre);
      } else {
        Logger.log('Ya existe: ' + nombre);
      }
      return sh;
    }

    crearSiNoExiste(TAB_DOCENTES,  ['ID','Nombre','Password','Email','FechaAlta']);
    crearSiNoExiste(TAB_MATERIAS,  ['ID','DocenteID','NombreMateria','NombreAula','RadioMetros','URL_Alumnos','SheetAsistencia_ID','FechaAlta','PesoTardanza']);
    crearSiNoExiste(TAB_REGISTROS, ['Fecha','Hora','Legajo','Alumno','MateriaID','Aula','Lat','Lng','Distancia','EstadoGPS','Token','DeviceId']);
    crearSiNoExiste(TAB_SESIONES,  ['Token','DocenteID','MateriaID','NombreMateria','NombreAula','Expira','Fecha','NombreClase','HoraInicio','URL_Alumnos','SheetAsistencia_ID']);
    crearSiNoExiste(TAB_DISPOSITIVOS,   ['DeviceId','Legajo','Nombre','MateriaID','FechaVinculo']);
    crearSiNoExiste(TAB_GRUPOS,         ['ID','MateriaID','DocenteID','Nombre','TPNumero','FechaInicio','FechaFin']);
    crearSiNoExiste(TAB_SESIONES_EVAL,  ['ID','MateriaID','DocenteID','GrupoID','TipoActividad','Nombre','Fecha','Estado','AsistenciaToken']);
    crearSiNoExiste(TAB_EVALUACIONES,   ['ID','SesionEvalID','Legajo','Alumno','DocenteID','TipoActividad','Calificacion','Proceso','Presentacion','Comentario','Fecha','Hora','EsRecuperatorio','SesionOriginalID']);

    // Verificar si ya existe admin
    var doc = ss.getSheetByName(TAB_DOCENTES);
    var rows = doc.getDataRange().getValues();
    var tieneAdmin = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === 'admin') { tieneAdmin = true; break; }
    }
    if (!tieneAdmin) {
      doc.appendRow(['admin', 'Administrador', 'admin', '', new Date()]);
      Logger.log('Admin creado');
    } else {
      Logger.log('Admin ya existe');
    }

    Logger.log('Setup completado OK');
    return { ok: true };
  } catch(e) {
    Logger.log('ERROR en setup: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ── ADMIN AUTH ───────────────────────────────────────────────
function esAdmin(password) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var doc = ss.getSheetByName(TAB_DOCENTES);
  if (!doc) return false;
  var rows = doc.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'admin' && String(rows[i][2]).trim() === String(password || '').trim()) return true;
  }
  return false;
}

// ── UTILS ────────────────────────────────────────────────────
function normalizarFecha(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  return String(val).slice(0,10);
}

function normalizarHora(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Argentina/Mendoza', 'HH:mm:ss');
  return String(val);
}

function extractSheetId(val) {
  if (!val) return '';
  var match = String(val).match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return String(val).trim();
}

function generarId() {
  return Utilities.getUuid().split('-')[0];
}

// ── LOGIN ────────────────────────────────────────────────────
function login(data) {
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var doc = ss.getSheetByName(TAB_DOCENTES);
    if (!doc) return { ok: false, error: 'Pestaña Docentes no encontrada. Ejecutá setupSheets primero.' };

    var rows = doc.getDataRange().getValues();
    var nombreBuscado = String(data.nombre || '').trim().toLowerCase();
    var passBuscada   = String(data.password || '').trim();

    for (var i = 1; i < rows.length; i++) {
      var id     = String(rows[i][0]).trim();
      var nombre = String(rows[i][1]).trim().toLowerCase();
      var pass   = String(rows[i][2]).trim();

      if ((nombre === nombreBuscado || id.toLowerCase() === nombreBuscado) && pass === passBuscada) {
        return { ok: true, rol: id === 'admin' ? 'admin' : 'docente', id: id, nombre: rows[i][1] };
      }
    }
    return { ok: false, error: 'Usuario o contraseña incorrectos' };
  } catch(e) {
    return { ok: false, error: 'Error: ' + e.message };
  }
}

// ── DOCENTES ─────────────────────────────────────────────────
function getDocentes(data) {
  if (!esAdmin(data.password)) return { ok: false, error: 'Sin permiso' };
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var doc = ss.getSheetByName(TAB_DOCENTES);
  var rows = doc.getDataRange().getValues().slice(1);
  var docentes = rows.filter(function(r){ return r[0] && String(r[0]) !== 'admin'; }).map(function(r) {
    return { id: r[0], nombre: r[1], email: r[3] || '' };
  });
  return { ok: true, docentes: docentes };
}

function saveDocente(data) {
  if (!esAdmin(data.adminPassword)) return { ok: false, error: 'Sin permiso' };
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var doc = ss.getSheetByName(TAB_DOCENTES);
  var id  = generarId();
  doc.appendRow([id, data.nombre, data.password, data.email || '', new Date()]);
  return { ok: true, id: id };
}

function deleteDocente(data) {
  if (!esAdmin(data.adminPassword)) return { ok: false, error: 'Sin permiso' };
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var doc = ss.getSheetByName(TAB_DOCENTES);
  var rows = doc.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) { doc.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'Docente no encontrado' };
}

// ── MATERIAS ─────────────────────────────────────────────────
function getMaterias(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var mat = ss.getSheetByName(TAB_MATERIAS);
  if (!mat) return { ok: true, materias: [] };
  var rows = mat.getDataRange().getValues().slice(1);
  var materias = rows.filter(function(r) {
    if (!r[0]) return false;
    if (data.rol === 'admin') return true;
    return String(r[1]) === String(data.docenteId);
  }).map(function(r) {
    return { id: r[0], docenteId: r[1], nombre: r[2], aula: r[3], radio: r[4]||200, urlAlumnos: r[5]||'', sheetId: extractSheetId(r[6]), pesoTardanza: r[8] !== undefined && r[8] !== '' ? parseInt(r[8]) : 100 };
  });
  return { ok: true, materias: materias };
}

function saveMateria(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var mat = ss.getSheetByName(TAB_MATERIAS);
  var id  = generarId();
  mat.appendRow([id, data.docenteId, data.nombre, data.aula, data.radio||200, data.urlAlumnos||'', extractSheetId(data.sheetId)||'', new Date(), data.pesoTardanza !== undefined ? data.pesoTardanza : 100]);
  return { ok: true, id: id };
}

function editarMateria(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var mat = ss.getSheetByName(TAB_MATERIAS);
  var rows = mat.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      if (data.nombre)                   mat.getRange(i+1,3).setValue(data.nombre);
      if (data.aula)                     mat.getRange(i+1,4).setValue(data.aula);
      if (data.radio)                    mat.getRange(i+1,5).setValue(data.radio);
      if (data.urlAlumnos !== undefined) mat.getRange(i+1,6).setValue(data.urlAlumnos);
      if (data.sheetId !== undefined)    mat.getRange(i+1,7).setValue(extractSheetId(data.sheetId));
      if (data.docenteId)                mat.getRange(i+1,2).setValue(data.docenteId);
      if (data.pesoTardanza !== undefined) mat.getRange(i+1,9).setValue(data.pesoTardanza);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Materia no encontrada' };
}

function deleteMateria(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var mat = ss.getSheetByName(TAB_MATERIAS);
  var rows = mat.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) { mat.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'Materia no encontrada' };
}

// ── ALUMNOS ──────────────────────────────────────────────────
var COLS_LEGAJO = ['legajo', 'leg', 'id', 'nro', 'numero', 'cod', 'codigo', 'matricula'];
var COLS_ALUMNO = ['alumno', 'nombre', 'apellido', 'apellidos', 'apellido y nombre', 'apellidos y nombres', 'nombre y apellido', 'estudiante'];

function encontrarColumna(headers, candidatos) {
  for (var i = 0; i < candidatos.length; i++) {
    var idx = headers.indexOf(candidatos[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getAlumnos(url) {
  if (!url) return { ok: false, error: 'URL no configurada' };
  try {
    var resp = UrlFetchApp.fetch(url);
    var rows = Utilities.parseCsv(resp.getContentText());
    var headers = rows[0].map(function(h){ return h.trim().toLowerCase(); });
    var idxLegajo = encontrarColumna(headers, COLS_LEGAJO);
    var idxAlumno = encontrarColumna(headers, COLS_ALUMNO);
    var idxEstado = headers.indexOf('estado');
    if (idxLegajo === -1 || idxAlumno === -1) {
      return { ok: false, error: 'Columnas no encontradas. El CSV debe tener una columna de legajo (' + COLS_LEGAJO.join('/') + ') y una de nombre (' + COLS_ALUMNO.join('/') + '). Columnas encontradas: ' + headers.join(', ') };
    }
    var alumnos = rows.slice(1).filter(function(r){ return r[idxLegajo] && r[idxLegajo].trim(); }).map(function(r) {
      return { legajo: r[idxLegajo].trim(), nombre: r[idxAlumno].trim(), estado: idxEstado>=0 ? r[idxEstado].trim() : '' };
    });
    return { ok: true, alumnos: alumnos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── SESION ───────────────────────────────────────────────────
function activarSesion(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var mat = ss.getSheetByName(TAB_MATERIAS);
  var ses = ss.getSheetByName(TAB_SESIONES);

  var matRows = mat.getDataRange().getValues();
  var materia = null;
  for (var i = 1; i < matRows.length; i++) {
    if (String(matRows[i][0]) === String(data.materiaId)) {
      materia = { id: matRows[i][0], docenteId: matRows[i][1], nombre: matRows[i][2], aula: matRows[i][3], radio: matRows[i][4]||200, urlAlumnos: matRows[i][5]||'', sheetId: extractSheetId(matRows[i][6])||'' };
      break;
    }
  }
  if (!materia) return { ok: false, error: 'Materia no encontrada' };

  var token      = Utilities.getUuid();
  var expira     = new Date(Date.now() + data.minutos * 60 * 1000);
  var fecha      = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  var horaInicio = new Date().toISOString();
  var nombreClase = String(data.nombreClase || 'Clase');

  ses.appendRow([token, data.docenteId, materia.id, materia.nombre, materia.aula, expira.toISOString(), fecha, nombreClase, horaInicio, materia.urlAlumnos, materia.sheetId, materia.pesoTardanza !== undefined ? materia.pesoTardanza : 100, expira.toISOString()]);
  ses.getRange(ses.getLastRow(), 8).setNumberFormat('@');

  var webAppUrl = ScriptApp.getService().getUrl();
  return { ok: true, token: token, qrUrl: webAppUrl+'?vista=alumno&token='+token, expira: expira.toISOString(), materia: materia, nombreClase: nombreClase, horaInicio: horaInicio };
}

function vincularDispositivo(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var dis = ss.getSheetByName(TAB_DISPOSITIVOS);
  if (!dis) { setupSheets(); dis = ss.getSheetByName(TAB_DISPOSITIVOS); }

  // Verificar si ya está vinculado
  var rows = dis.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.deviceId && rows[i][3] === data.materiaId) {
      return { ok: true, yaVinculado: true };
    }
  }

  dis.appendRow([data.deviceId, data.legajo, data.nombre, data.materiaId, new Date()]);
  return { ok: true, yaVinculado: false };
}

function buscarDispositivo(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var dis = ss.getSheetByName(TAB_DISPOSITIVOS);
  if (!dis) return { ok: true, alumno: null };

  var rows = dis.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.deviceId && rows[i][3] === data.materiaId) {
      return { ok: true, alumno: { legajo: rows[i][1], nombre: rows[i][2] } };
    }
  }
  return { ok: true, alumno: null };
}

function getDashboard(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var reg = ss.getSheetByName(TAB_REGISTROS);
  var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');

  var rows = reg.getDataRange().getValues().slice(1);
  var registros = rows.filter(function(r) {
    return normalizarFecha(r[0]) === hoy && String(r[4]) === String(data.materiaId) && r[10] === data.token;
  });

  var presentes  = registros.filter(function(r){ return r[9]==='OK'||r[9]==='MANUAL'; }).length;
  var tardanzas  = registros.filter(function(r){ return r[9]==='TARDANZA_MANUAL'; }).length;
  var duplicados = registros.filter(function(r){ return r[9]==='DISPOSITIVO_DUPLICADO'; }).length;
  var total      = presentes + tardanzas;

  return { ok: true, presentes: presentes, tardanzas: tardanzas, duplicados: duplicados, total: total, hora: Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'HH:mm') };
}

function extenderSesion(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var ses = ss.getSheetByName(TAB_SESIONES);
  var rows = ses.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.token) {
      var nuevaExpira = new Date(Date.now() + data.minutos * 60 * 1000);
      ses.getRange(i+1, 6).setValue(nuevaExpira.toISOString());
      return { ok: true, expira: nuevaExpira.toISOString() };
    }
  }
  return { ok: false, error: 'Token no encontrado' };
}

function getSesionActiva(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var ses = ss.getSheetByName(TAB_SESIONES);
  if (!ses) return { ok: true, sesion: null };
  var rows = ses.getDataRange().getValues();
  var now  = new Date();
  for (var i = rows.length-1; i >= 1; i--) {
    if (data.docenteId !== 'admin' && String(rows[i][1]) !== String(data.docenteId)) continue;
    if (new Date(rows[i][5]) > now) return { ok: true, sesion: buildSesion(rows[i]) };
  }
  return { ok: true, sesion: null };
}

function getSesionPorToken(token) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var ses = ss.getSheetByName(TAB_SESIONES);
  var rows = ses.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === token) {
      if (new Date() > new Date(rows[i][5])) return { ok: false, error: 'QR expirado' };
      return { ok: true, sesion: buildSesion(rows[i]) };
    }
  }
  return { ok: false, error: 'Token no encontrado' };
}

function buildSesion(row) {
  var nombreClase = row[7];
  if (nombreClase instanceof Date) nombreClase = Utilities.formatDate(nombreClase, 'America/Argentina/Mendoza', 'dd/MM/yyyy');
  return { token: row[0], docenteId: row[1], materiaId: row[2], materia: row[3], aula: row[4], expira: row[5], fecha: row[6], nombreClase: String(nombreClase||'Clase'), horaInicio: row[8]||new Date().toISOString(), urlAlumnos: row[9]||'', sheetId: row[10]||'', pesoTardanza: row[11] !== undefined && row[11] !== '' ? parseInt(row[11]) : 100, expiraOriginal: row[12] ? String(row[12]) : String(row[5]) };
}

function volcarUltimaSesion() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var ses = ss.getSheetByName(TAB_SESIONES);
  var rows = ses.getDataRange().getValues();
  if (rows.length <= 1) { Logger.log('Sin sesiones'); return; }
  var ultima = rows[rows.length - 1];

  Logger.log('Total columnas: ' + ultima.length);
  for (var i = 0; i < ultima.length; i++) {
    Logger.log('Col ' + i + ': [' + ultima[i] + ']');
  }

  var sesion = buildSesion(ultima);
  Logger.log('sheetId: ' + sesion.sheetId);
  Logger.log('urlAlumnos: ' + sesion.urlAlumnos);

  if (!sesion.sheetId || !sesion.urlAlumnos) {
    Logger.log('ERROR: sheetId o urlAlumnos vacios');
    return;
  }

  var resultado = volcarAsistenciaPlanilla(sesion);
  Logger.log('Resultado: ' + JSON.stringify(resultado));
}

function cerrarSesion(data) {
  try {
    var sesion = null;
    // Buscar sesión por token en el Sheet
    if (data.sesion && data.sesion.token) {
      var ss  = SpreadsheetApp.openById(SHEET_ID);
      var ses = ss.getSheetByName(TAB_SESIONES);
      var rows = ses.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.sesion.token) {
          sesion = buildSesion(rows[i]);
          break;
        }
      }
    }
    if (sesion) {
      Logger.log('Cerrando sesion token: ' + sesion.token);
      Logger.log('sheetId: ' + sesion.sheetId);
      Logger.log('urlAlumnos: ' + sesion.urlAlumnos);
      var resultado = volcarAsistenciaPlanilla(sesion);
      Logger.log('Resultado: ' + JSON.stringify(resultado));
    } else {
      Logger.log('No se encontro sesion para cerrar');
    }
    return { ok: true };
  } catch(e) {
    Logger.log('Error cerrarSesion: ' + e.message);
    return { ok: true };
  }
}

function expirarSesion(data) {
  try {
    var sesion = null;
    if (data.sesion && data.sesion.token) {
      var ss  = SpreadsheetApp.openById(SHEET_ID);
      var ses = ss.getSheetByName(TAB_SESIONES);
      var rows = ses.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.sesion.token) {
          sesion = buildSesion(rows[i]);
          break;
        }
      }
    }
    if (sesion) volcarAsistenciaPlanilla(sesion);
    return { ok: true };
  } catch(e) {
    Logger.log('Error expirarSesion: ' + e.message);
    return { ok: true };
  }
}

// ── MARCAR PRESENTE ──────────────────────────────────────────
function marcarPresente(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var ses = ss.getSheetByName(TAB_SESIONES);
  var reg = ss.getSheetByName(TAB_REGISTROS);

  var sesRows = ses.getDataRange().getValues();
  var sesion = null;
  for (var i = 1; i < sesRows.length; i++) {
    if (sesRows[i][0] === data.token) {
      if (new Date() > new Date(sesRows[i][5])) return { ok: false, error: 'QR expirado' };
      sesion = buildSesion(sesRows[i]);
      break;
    }
  }
  if (!sesion) return { ok: false, error: 'Token invalido' };

  var regRows = reg.getDataRange().getValues();
  var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');

  for (var j = 1; j < regRows.length; j++) {
    if (String(regRows[j][2]).trim() === String(data.legajo).trim() && regRows[j][10] === data.token) {
      return { ok: false, error: 'Ya registraste tu presencia en esta clase' };
    }
  }

  var deviceId = data.deviceId || '';
  if (deviceId) {
    for (var k = 1; k < regRows.length; k++) {
      if (regRows[k][11] === deviceId && regRows[k][10] === data.token) {
        return { ok: false, error: 'Este celular ya registró a ' + regRows[k][3] + '. Cada alumno debe usar su propio celular.' };
      }
    }
  }

  var hora = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'HH:mm:ss');
  reg.appendRow([hoy, hora, data.legajo, data.nombre, sesion.materiaId, sesion.aula, '', '', 0, 'OK', data.token, deviceId]);
  reg.getRange(reg.getLastRow(), 1, 1, 12).setBackground('#d1fae5');

  return { ok: true, dispositivoDuplicado: false };
}

// ── TOGGLE MANUAL ────────────────────────────────────────────
function toggleManual(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var reg = ss.getSheetByName(TAB_REGISTROS);
  var hoy  = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  var hora = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'HH:mm:ss');
  var estado = data.estado || 'MANUAL';

  var rows = reg.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() === String(data.legajo).trim() && rows[i][10] === data.token) {
      reg.deleteRow(i+1);
      return { ok: true, accion: 'eliminado' };
    }
  }
  reg.appendRow([hoy, hora, data.legajo, data.nombre, data.materiaId, data.aula, '', '', 0, estado, data.token, '']);
  var color = estado === 'TARDANZA_MANUAL' ? '#fef3c7' : '#dbeafe';
  reg.getRange(reg.getLastRow(), 1, 1, 12).setBackground(color);
  return { ok: true, accion: 'agregado' };
}

// ── REGISTROS DEL DIA ────────────────────────────────────────
function getRegistrosDia(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var reg = ss.getSheetByName(TAB_REGISTROS);
  if (!reg) return { ok: true, registros: [] };
  var rows = reg.getDataRange().getValues();
  var registros = rows.slice(1).filter(function(r) {
    var rf = normalizarFecha(r[0]);
    var tokenMatch = !data.token || r[10] === data.token;
    var materiaMatch = !data.materiaId || String(r[4]) === String(data.materiaId);
    return rf === data.fecha && tokenMatch && materiaMatch;
  }).map(function(r) {
    return { fecha: r[0], hora: normalizarHora(r[1]), legajo: String(r[2]).trim(), nombre: r[3], materiaId: r[4], gps: r[9], token: r[10]||'', deviceId: r[11]||'' };
  });

  var deviceCount = {};
  registros.forEach(function(r) {
    var key = r.token+'|'+r.deviceId;
    if (r.deviceId) deviceCount[key] = (deviceCount[key]||0)+1;
  });
  registros.forEach(function(r) {
    var key = r.token+'|'+r.deviceId;
    if (r.deviceId && deviceCount[key] > 1) r.gps = 'DISPOSITIVO_DUPLICADO';
  });

  return { ok: true, registros: registros };
}

// ── VOLCAR A PLANILLA ────────────────────────────────────────
function volcarAsistenciaPlanilla(sesion) {
  var planillaId = sesion.sheetId;
  var urlAlumnos = sesion.urlAlumnos;
  if (!planillaId || !urlAlumnos) return { ok: false, error: 'Sin Sheet o URL configurados' };

  var ss = SpreadsheetApp.openById(planillaId);
  var pl = ss.getSheetByName('Asistencia');
  if (!pl) {
    pl = ss.insertSheet('Asistencia');
    pl.getRange('A1').setValue('ALUMNOS').setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    pl.getRange('B1').setValue('%').setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff').setHorizontalAlignment('center');
    pl.setColumnWidth(1, 220);
    pl.setColumnWidth(2, 60);
    pl.setFrozenColumns(1);
    pl.setFrozenRows(1);
  }

  var todosAlumnos = getAlumnos(urlAlumnos).alumnos || [];
  if (pl.getLastRow() <= 1) {
    todosAlumnos.forEach(function(a, i) {
      pl.getRange(i+2, 1).setValue(a.nombre).setBackground(i%2===0 ? '#f8fafc' : '#ffffff');
    });
  }

  var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  var registros = getRegistrosDia({fecha: hoy, token: sesion.token, materiaId: sesion.materiaId}).registros;

  var horaInicioStr = Utilities.formatDate(new Date(sesion.horaInicio), 'America/Argentina/Mendoza', 'HH:mm:ss');
  var hiP = horaInicioStr.split(':');
  var horaInicioLocal = new Date();
  horaInicioLocal.setHours(parseInt(hiP[0]), parseInt(hiP[1]), parseInt(hiP[2]||0), 0);

  // Hora límite de presencia a tiempo = momento en que expiró el QR original (antes de la tardanza)
  var expOrig = sesion.expiraOriginal || sesion.expira;
  var expOrigStr = Utilities.formatDate(new Date(expOrig), 'America/Argentina/Mendoza', 'HH:mm:ss').split(':');
  var expiraOriginalLocal = new Date();
  expiraOriginalLocal.setHours(parseInt(expOrigStr[0]), parseInt(expOrigStr[1]), parseInt(expOrigStr[2]||0), 0);

  var marcasPorLegajo = {};
  registros.forEach(function(r) { marcasPorLegajo[r.legajo] = { hora: r.hora, gps: r.gps }; });

  var nombreALegajo = {};
  todosAlumnos.forEach(function(a) {
    nombreALegajo[a.nombre.trim().toUpperCase().replace(/\s+/g,' ')] = a.legajo;
  });

  var nombreCol = (sesion.nombreClase||'Clase') + '\n' + Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'dd/MM/yy');
  var lastCol = pl.getLastColumn();
  // Clases arrancan en columna 3 (C), columna 2 (B) es el %
  var headers = lastCol >= 3 ? pl.getRange(1, 3, 1, lastCol - 2).getValues()[0] : [];
  var colIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === nombreCol) { colIdx = i + 3; break; }
  }
  if (colIdx === -1) {
    colIdx = lastCol < 3 ? 3 : lastCol + 1;
    pl.getRange(1, colIdx).setValue(nombreCol)
      .setFontWeight('bold').setBackground('#2357e8').setFontColor('#ffffff')
      .setWrap(true).setHorizontalAlignment('center');
    pl.setColumnWidth(colIdx, 90);
  }

  var lastRow = pl.getLastRow();
  if (lastRow <= 1) return { ok: false, error: 'Planilla sin alumnos' };

  var alumnosEnPlanilla = pl.getRange(2, 1, lastRow-1, 1).getValues();
  alumnosEnPlanilla.forEach(function(row, i) {
    var cell = pl.getRange(i+2, colIdx);
    if (cell.getValue() !== '' && cell.getValue() !== null) return;
    var keyNorm = String(row[0]).trim().toUpperCase().replace(/\s+/g,' ');
    var legajo = nombreALegajo[keyNorm];
    if (!legajo) {
      var keys = Object.keys(nombreALegajo);
      for (var k = 0; k < keys.length; k++) {
        if (keys[k].indexOf(keyNorm) !== -1 || keyNorm.indexOf(keys[k]) !== -1) { legajo = nombreALegajo[keys[k]]; break; }
      }
    }
    var estado = 'A', color = '#fee2e2', fontColor = '#991b1b';
    if (legajo && marcasPorLegajo[legajo]) {
      var m = marcasPorLegajo[legajo];
      var marcaDate = new Date();
      var hp = String(m.hora).split(':');
      marcaDate.setHours(parseInt(hp[0]||0), parseInt(hp[1]||0), parseInt(hp[2]||0), 0);
      var diff = (marcaDate - horaInicioLocal) / 60000;
      if (m.gps === 'DISPOSITIVO_DUPLICADO') { estado='PD'; color='#fde8d8'; fontColor='#ea580c'; }
      else if (m.gps === 'TARDANZA_MANUAL') { estado='T'; color='#fef3c7'; fontColor='#92400e'; }
      else if (marcaDate <= expiraOriginalLocal) { estado='P'; color='#d1fae5'; fontColor='#065f46'; }
      else { estado='T'; color='#fef3c7'; fontColor='#92400e'; }
    }
    cell.setValue(estado).setBackground(color).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(fontColor);
    if (estado === 'PD') cell.setNote('Marcado desde dispositivo ya usado por otro alumno');
  });

  actualizarPorcentajes(pl, sesion.pesoTardanza);
  return { ok: true };
}

// ── GRUPOS ───────────────────────────────────────────────────
function getGrupos(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_GRUPOS);
  if (!sh) return { ok: true, grupos: [] };
  var rows = sh.getDataRange().getValues().slice(1);
  var grupos = rows.filter(function(r) {
    if (!r[0]) return false;
    if (data.materiaId && String(r[1]) !== String(data.materiaId)) return false;
    if (data.docenteId && data.docenteId !== 'admin' && String(r[2]) !== String(data.docenteId)) return false;
    return true;
  }).map(function(r) {
    return { id: r[0], materiaId: r[1], docenteId: r[2], nombre: r[3], tpNumero: r[4]||0, fechaInicio: normalizarFecha(r[5]), fechaFin: normalizarFecha(r[6]), urlAlumnos: r[7]||'' };
  });
  return { ok: true, grupos: grupos };
}

function saveGrupo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_GRUPOS);
  var id = generarId();
  sh.appendRow([id, data.materiaId, data.docenteId, data.nombre, data.tpNumero||0, data.fechaInicio||'', data.fechaFin||'', data.urlAlumnos||'']);
  return { ok: true, id: id, urlAlumnos: data.urlAlumnos||'' };
}

function editarGrupo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_GRUPOS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      if (data.nombre !== undefined)      sh.getRange(i+1,4).setValue(data.nombre);
      if (data.tpNumero !== undefined)    sh.getRange(i+1,5).setValue(data.tpNumero);
      if (data.fechaInicio !== undefined) sh.getRange(i+1,6).setValue(data.fechaInicio);
      if (data.fechaFin !== undefined)    sh.getRange(i+1,7).setValue(data.fechaFin);
      if (data.urlAlumnos !== undefined)  sh.getRange(i+1,8).setValue(data.urlAlumnos);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Grupo no encontrado' };
}

function deleteGrupo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_GRUPOS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) { sh.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'Grupo no encontrado' };
}

// ── SESIONES EVAL ────────────────────────────────────────────
function abrirSesionEval(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_SESIONES_EVAL);
  var id = generarId();
  var fecha = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  sh.appendRow([id, data.materiaId, data.docenteId, data.grupoId||'', data.tipoActividad, data.nombre||'', fecha, 'abierta', data.asistenciaToken||'']);
  return { ok: true, id: id, fecha: fecha };
}

function getSesionesEval(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_SESIONES_EVAL);
  if (!sh) return { ok: true, sesiones: [] };
  var rows = sh.getDataRange().getValues().slice(1);
  var sesiones = rows.filter(function(r) {
    if (!r[0]) return false;
    if (data.materiaId && String(r[1]) !== String(data.materiaId)) return false;
    if (data.grupoId   && String(r[3]) !== String(data.grupoId))   return false;
    if (data.tipoActividad && String(r[4]) !== String(data.tipoActividad)) return false;
    return true;
  }).map(function(r) {
    return { id: r[0], materiaId: r[1], docenteId: r[2], grupoId: r[3], tipoActividad: r[4], nombre: r[5], fecha: normalizarFecha(r[6]), estado: r[7], asistenciaToken: r[8]||'' };
  });
  return { ok: true, sesiones: sesiones };
}

function cerrarSesionEval(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_SESIONES_EVAL);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sh.getRange(i+1, 8).setValue('cerrada');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Sesión eval no encontrada' };
}

// ── EVALUACIONES ─────────────────────────────────────────────
function saveEvaluacion(data) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var sh  = ss.getSheetByName(TAB_EVALUACIONES);
  var fecha = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  var hora  = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'HH:mm:ss');

  var proceso      = data.proceso      !== undefined ? data.proceso      : '';
  var presentacion = data.presentacion !== undefined ? data.presentacion : '';
  var calificacion = data.calificacion !== undefined ? data.calificacion : '';
  if (proceso !== '' && presentacion !== '') calificacion = parseInt(proceso) + parseInt(presentacion);

  // Actualizar si ya existe registro para este alumno en esta sesión (no es recuperatorio)
  if (!data.esRecuperatorio) {
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(data.sesionEvalId) && String(rows[i][2]).trim() === String(data.legajo).trim()) {
        sh.getRange(i+1, 7).setValue(calificacion);
        sh.getRange(i+1, 8).setValue(proceso);
        sh.getRange(i+1, 9).setValue(presentacion);
        if (data.comentario !== undefined) sh.getRange(i+1, 10).setValue(data.comentario);
        sh.getRange(i+1, 12).setValue(hora);
        return { ok: true, accion: 'actualizado' };
      }
    }
  }

  var id = generarId();
  sh.appendRow([id, data.sesionEvalId, data.legajo, data.nombre, data.docenteId, data.tipoActividad, calificacion, proceso, presentacion, data.comentario||'', fecha, hora, data.esRecuperatorio ? 1 : 0, data.sesionOriginalId||'']);
  return { ok: true, id: id, accion: 'creado' };
}

function getEvaluaciones(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_EVALUACIONES);
  if (!sh) return { ok: true, evaluaciones: [] };
  var rows = sh.getDataRange().getValues().slice(1);
  var evals = rows.filter(function(r) {
    if (!r[0]) return false;
    if (data.sesionEvalId && String(r[1]) !== String(data.sesionEvalId)) return false;
    if (data.legajo       && String(r[2]).trim() !== String(data.legajo).trim()) return false;
    return true;
  }).map(function(r) {
    return { id: r[0], sesionEvalId: r[1], legajo: String(r[2]).trim(), alumno: r[3], docenteId: r[4], tipoActividad: r[5], calificacion: r[6], proceso: r[7], presentacion: r[8], comentario: r[9], fecha: normalizarFecha(r[10]), hora: normalizarHora(r[11]), esRecuperatorio: !!r[12], sesionOriginalId: r[13]||'' };
  });
  return { ok: true, evaluaciones: evals };
}

function getHistorialAlumno(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var evSh  = ss.getSheetByName(TAB_EVALUACIONES);
  var sesSh = ss.getSheetByName(TAB_SESIONES_EVAL);
  if (!evSh) return { ok: true, evaluaciones: [] };

  var sesMap = {};
  if (sesSh) {
    sesSh.getDataRange().getValues().slice(1).forEach(function(r) {
      if (r[0]) sesMap[r[0]] = { tipoActividad: r[4], nombre: r[5], fecha: normalizarFecha(r[6]) };
    });
  }

  var rows = evSh.getDataRange().getValues().slice(1);
  var evals = rows.filter(function(r) {
    return r[0] && String(r[2]).trim() === String(data.legajo).trim();
  }).map(function(r) {
    var ses = sesMap[r[1]] || {};
    return { id: r[0], sesionEvalId: r[1], tipoActividad: r[5]||ses.tipoActividad||'', nombre: ses.nombre||'', fecha: normalizarFecha(r[10])||ses.fecha||'', calificacion: r[6], proceso: r[7], presentacion: r[8], comentario: r[9], esRecuperatorio: !!r[12] };
  });
  return { ok: true, evaluaciones: evals };
}

// ── RESUMEN ──────────────────────────────────────────────────
function calcularResumenAlumno(legajo, nombre, materiaId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var evSh  = ss.getSheetByName(TAB_EVALUACIONES);
  var sesSh = ss.getSheetByName(TAB_SESIONES_EVAL);
  var regSh = ss.getSheetByName(TAB_REGISTROS);

  var sesMap = {};
  if (sesSh) {
    sesSh.getDataRange().getValues().slice(1).forEach(function(r) {
      if (r[0]) sesMap[String(r[0])] = { tipoActividad: r[4], materiaId: r[1] };
    });
  }

  var evRows = evSh ? evSh.getDataRange().getValues().slice(1).filter(function(r) {
    if (!r[0] || String(r[2]).trim() !== String(legajo).trim()) return false;
    if (materiaId) {
      var ses = sesMap[String(r[1])];
      if (ses && ses.materiaId && String(ses.materiaId) !== String(materiaId)) return false;
    }
    return true;
  }) : [];

  // Consultas
  var consultas = evRows.filter(function(r) { return (r[5]||'') === 'consulta'; });
  var consultaVals = consultas.map(function(r) { return r[6] !== '' && r[6] !== null ? parseFloat(r[6]) : 0; });
  var consultaPromedio = consultaVals.length > 0 ? Math.round(consultaVals.reduce(function(a,b){return a+b;},0) / consultaVals.length * 10) / 10 : null;

  // Scores por tipo — recuperatorio pisa al original
  var scores = {};
  evRows.forEach(function(r) {
    var tipo = String(r[5]||'');
    if (!tipo || tipo === 'consulta') return;
    var cal = r[6] !== '' && r[6] !== null ? parseFloat(r[6]) : null;
    if (cal === null) return;
    var key = r[12] ? tipo + '_recup' : tipo;
    scores[key] = { nota: cal, proceso: r[7], presentacion: r[8], sesionId: r[1] };
  });

  function getScore(tipo) {
    var recup = scores[tipo + '_recup'];
    var orig  = scores[tipo];
    if (recup) return { nota: recup.nota, proceso: recup.proceso, presentacion: recup.presentacion, recuperado: true };
    if (orig)  return { nota: orig.nota,  proceso: orig.proceso,  presentacion: orig.presentacion,  recuperado: false };
    return null;
  }

  var tps      = {}; TIPOS_TP.forEach(function(t)     { tps[t]      = getScore(t); });
  var parciales = {}; TIPOS_PARCIAL.forEach(function(t) { parciales[t] = getScore(t); });

  // Asistencia
  var asistencia = null;
  if (regSh) {
    var regRows = regSh.getDataRange().getValues().slice(1).filter(function(r) {
      return String(r[2]).trim() === String(legajo).trim() && (!materiaId || String(r[4]) === String(materiaId));
    });
    var presentes = regRows.filter(function(r) { return r[9]==='OK'||r[9]==='MANUAL'||r[9]==='TARDANZA_MANUAL'; }).length;
    asistencia = regRows.length > 0 ? Math.round(presentes / regRows.length * 100) : null;
  }

  // Estado final
  var tpNotas      = TIPOS_TP.map(function(t)     { return tps[t]      ? tps[t].nota      : null; }).filter(function(n){ return n!==null; });
  var parcialNotas = TIPOS_PARCIAL.map(function(t) { return parciales[t] ? parciales[t].nota : null; }).filter(function(n){ return n!==null; });
  var estado = 'en_curso';

  if (tpNotas.some(function(n){return n<NOTA_APROBACION;}) || parcialNotas.some(function(n){return n<NOTA_APROBACION;})) {
    estado = 'en_riesgo';
  }
  if (tpNotas.length === 4 && parcialNotas.length === 2) {
    var tpProm  = tpNotas.reduce(function(a,b){return a+b;},0) / 4;
    var parProm = parcialNotas.reduce(function(a,b){return a+b;},0) / 2;
    if (tpNotas.some(function(n){return n<NOTA_APROBACION;}) || parcialNotas.some(function(n){return n<NOTA_APROBACION;})) {
      estado = 'recursa';
    } else if (tpProm >= NOTA_PROMOCION && parProm >= NOTA_PROMOCION) {
      estado = 'promociona';
    } else {
      estado = 'regular';
    }
  }

  return { legajo: legajo, alumno: nombre, asistencia: asistencia, consultas: { total: consultas.length, promedio: consultaPromedio }, tps: tps, parciales: parciales, estado: estado };
}

function getResumenAlumno(data) {
  try {
    return { ok: true, resumen: calcularResumenAlumno(data.legajo, data.alumno||data.legajo, data.materiaId||'') };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getResumenGrupo(data) {
  try {
    var alumnosResult = getAlumnos(data.urlAlumnos||'');
    if (!alumnosResult.ok) return alumnosResult;
    var resumenes = alumnosResult.alumnos.map(function(a) {
      return calcularResumenAlumno(a.legajo, a.nombre, data.materiaId||'');
    });
    return { ok: true, resumenes: resumenes };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function actualizarPorcentajes(pl, pesoTardanza) {
  var peso = pesoTardanza !== undefined ? pesoTardanza : 100;
  var lastRow = pl.getLastRow();
  if (lastRow <= 1) return;

  var headerB = pl.getRange(1, 2).getValue();
  if (headerB !== '%') {
    pl.getRange(1, 2).setValue('%').setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff').setHorizontalAlignment('center');
    pl.setColumnWidth(2, 60);
  }

  var lastCol = pl.getLastColumn();
  var clasesCols = lastCol - 2;

  for (var r = 2; r <= lastRow; r++) {
    var pct = 0;
    if (clasesCols > 0) {
      var rowData = pl.getRange(r, 3, 1, clasesCols).getValues()[0];
      var total = rowData.filter(function(v){ return v!==''&&v!==null&&v!==0; }).length;
      if (total > 0) {
        var puntos = 0;
        rowData.forEach(function(v) {
          if (v==='P' || v==='PD') puntos += 1;
          else if (v==='T') puntos += (peso / 100);
        });
        pct = Math.round((puntos / total) * 100);
      }
    }
    var color = pct>=75 ? '#d1fae5' : pct>=50 ? '#fef3c7' : '#fee2e2';
    var fontColor = pct>=75 ? '#065f46' : pct>=50 ? '#92400e' : '#991b1b';
    pl.getRange(r, 2).setValue(pct+'%').setFontWeight('bold').setHorizontalAlignment('center').setBackground(color).setFontColor(fontColor);
  }
}
