// ═══════════════════════════════════════════════════════════════
// APPS SCRIPT — SISTEMA OT PATAGONIA
// Maneja: Tareas, Sesiones, Usuarios, Listas, Config
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Encabezados por hoja ────────────────────────────────────────
const HEADERS = {
  Tareas:   ['ID','Nombre','Tarea','Solicitante','Area','Tipo','Prioridad',
             'HH_Plan','FechaIngreso','FechaCompromiso','FechaEjecucion','Estado','Comentario',
             'HH_Real','N_Sesiones'],
  Sesiones: ['TareaID','TareaNombre','Responsable','Fecha',
             'HoraInicio','HoraTermino','HH_Sesion','Comentario','FechaRegistro'],
  Usuarios: ['Usuario','Nombre','Password','Rol','UltimoRegistro'],
  Listas:   ['Tipo','Valor'],
  Config:   ['Clave','Valor'],
};

// ── Datos iniciales ─────────────────────────────────────────────
const INIT_USUARIOS = [
  ['lsanchez','Luis Sanchez','pat2024','admin',''],
  ['ldiaz',   'Luis Diaz',   'pat2024','user', ''],
  ['ysilva',  'Yamila Silva','pat2024','user', ''],
];

const INIT_LISTAS = [
  ['Area',      'Ventas'],
  ['Area',      'Oficina Tecnica'],
  ['Area',      'Producción'],
  ['Area',      'Logistica'],
  ['Area',      'Postventa'],
  ['Area',      'Bodega'],
  ['Tipo',      'Coordinación'],
  ['Tipo',      'Plano'],
  ['Tipo',      'Costo'],
  ['Tipo',      'Consulta'],
  ['Tipo',      'EDP'],
  ['Tipo',      'Ejecución'],
  ['Tipo',      'Informe'],
  ['Tipo',      'Vales'],
  ['Prioridad', 'Alta'],
  ['Prioridad', 'Media'],
  ['Prioridad', 'Baja'],
  ['Estado',    'Por Iniciar'],
  ['Estado',    'En Proceso'],
  ['Estado',    'Listo'],
  ['Estado',    'Programada'],
  ['Estado',    'Standby'],
  ['Estado',    'Cancelado'],
  ['Solicitante','Thania'],
  ['Solicitante','Monica'],
  ['Solicitante','Brezzy'],
  ['Solicitante','Rodolfo'],
  ['Solicitante','Javiera'],
  ['Solicitante','Eniliam'],
  ['Solicitante','Patricia'],
  ['Solicitante','Claudia'],
  ['Solicitante','Michael'],
  ['Solicitante','Juan'],
  ['Solicitante','Sofia'],
  ['Solicitante','Tatiana'],
  ['Solicitante','Pedro'],
  ['Solicitante','Roxana'],
];

// ── Inicializar hojas si están vacías ───────────────────────────
function inicializar() {
  Object.entries(HEADERS).forEach(([nombre, hdrs]) => {
    const sheet = SS.getSheetByName(nombre);
    if (!sheet) return;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(hdrs);
      if (nombre === 'Usuarios') INIT_USUARIOS.forEach(r => sheet.appendRow(r));
      if (nombre === 'Listas')   INIT_LISTAS.forEach(r => sheet.appendRow(r));
    }
  });
}

// ── GET handler ─────────────────────────────────────────────────
function doGet(e) {
  inicializar();
  const accion = e.parameter.accion || 'leer';
  const tabla  = e.parameter.tabla  || '';

  let resultado;

  if (accion === 'leer') {
    resultado = leerTabla(tabla);
  } else if (accion === 'login') {
    resultado = login(e.parameter.usuario, e.parameter.password);
  } else if (accion === 'listas') {
    resultado = obtenerListas();
  } else if (accion === 'cumplimiento') {
    resultado = obtenerCumplimiento();
  } else if (accion === 'dashboard') {
    resultado = obtenerDashboard(e.parameter.desde, e.parameter.hasta);
  } else if (accion === 'eliminarSesion') {
    resultado = eliminarSesion(data.tareaId, data.fecha, data.inicio);
  } else if (accion === 'eliminarPorTareaId') {
    resultado = eliminarSesionesPorTarea(data.id);
  } else if (accion === 'actualizarSesion') {
    resultado = actualizarSesion(data.tareaId, data.fechaOrig, data.inicioOrig, data.campos);
  } else {
    resultado = { error: 'Acción no reconocida' };
  }

  return jsonResponse(resultado);
}

// ── POST handler ────────────────────────────────────────────────
function doPost(e) {
  inicializar();
  const data = JSON.parse(e.postData.contents);
  const accion = data.accion || 'insertar';

  let resultado;

  if (accion === 'insertar') {
    resultado = insertarFila(data.tabla, data.fila);
  } else if (accion === 'actualizar') {
    resultado = actualizarFila(data.tabla, data.id, data.campos);
  } else if (accion === 'eliminar') {
    resultado = eliminarFila(data.tabla, data.id);
  } else if (accion === 'guardarLista') {
    resultado = guardarLista(data.items);
  } else if (accion === 'actualizarUltimoRegistro') {
    resultado = actualizarUltimoRegistro(data.usuario);
  } else if (accion === 'eliminarSesion') {
    resultado = eliminarSesion(data.tareaId, data.fecha, data.inicio);
  } else if (accion === 'eliminarPorTareaId') {
    resultado = eliminarSesionesPorTarea(data.id);
  } else if (accion === 'actualizarSesion') {
    resultado = actualizarSesion(data.tareaId, data.fechaOrig, data.inicioOrig, data.campos);
  } else {
    resultado = { error: 'Acción no reconocida' };
  }

  return jsonResponse(resultado);
}

// ── CRUD básico ─────────────────────────────────────────────────
function leerTabla(nombre) {
  const sheet = SS.getSheetByName(nombre);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues();
}

function insertarFila(nombre, fila) {
  const sheet = SS.getSheetByName(nombre);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada' };
  sheet.appendRow(fila);
  return { ok: true };
}

function actualizarFila(nombre, id, campos) {
  const sheet = SS.getSheetByName(nombre);
  if (!sheet) return { ok: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.entries(campos).forEach(([col, val]) => {
        const ci = parseInt(col);
        sheet.getRange(i + 1, ci + 1).setValue(val);
      });
      return { ok: true };
    }
  }
  return { ok: false, error: 'Fila no encontrada' };
}

function eliminarFila(nombre, id) {
  const sheet = SS.getSheetByName(nombre);
  if (!sheet) return { ok: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ── Login ───────────────────────────────────────────────────────
function login(usuario, password) {
  const sheet = SS.getSheetByName('Usuarios');
  if (!sheet) return { ok: false, error: 'Sin hoja Usuarios' };
  const rows = sheet.getDataRange().getValues().slice(1);
  const u = rows.find(r => r[0] === usuario && r[2] === password);
  if (!u) return { ok: false, error: 'Usuario o contraseña incorrectos' };
  return { ok: true, usuario: u[0], nombre: u[1], rol: u[3] };
}

// ── Listas (para dropdowns) ─────────────────────────────────────
function obtenerListas() {
  const sheet = SS.getSheetByName('Listas');
  if (!sheet || sheet.getLastRow() <= 1) return {};
  const rows = sheet.getDataRange().getValues().slice(1);
  const result = {};
  rows.forEach(r => {
    const tipo = r[0]; const val = r[1];
    if (!result[tipo]) result[tipo] = [];
    result[tipo].push(val);
  });
  return result;
}

function guardarLista(items) {
  const sheet = SS.getSheetByName('Listas');
  if (!sheet) return { ok: false };
  // Limpiar desde fila 2
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  items.forEach(item => sheet.appendRow([item.tipo, item.valor]));
  return { ok: true };
}

// ── Cumplimiento diario ─────────────────────────────────────────
function obtenerCumplimiento() {
  const uSheet = SS.getSheetByName('Usuarios');
  const sSheet = SS.getSheetByName('Sesiones');
  if (!uSheet) return [];

  const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  const usuarios = uSheet.getDataRange().getValues().slice(1);

  // Obtener fechas de último registro por usuario (desde Sesiones y Tareas)
  let sesiones = [];
  if (sSheet && sSheet.getLastRow() > 1) {
    sesiones = sSheet.getDataRange().getValues().slice(1);
  }

  return usuarios.map(u => {
    const nombre = u[1];
    // Buscar última sesión de este usuario
    const susSesiones = sesiones.filter(s => s[2] === nombre);
    let ultimaFecha = '';
    if (susSesiones.length > 0) {
      const fechas = susSesiones.map(s => s[3]).filter(f => f).sort().reverse();
      if (fechas.length > 0) {
        ultimaFecha = typeof fechas[0] === 'object'
          ? Utilities.formatDate(fechas[0], 'America/Santiago', 'yyyy-MM-dd')
          : String(fechas[0]).substring(0, 10);
      }
    }

    let estado = 'rojo';
    if (ultimaFecha === hoy) estado = 'verde';
    else if (ultimaFecha) {
      const diff = (new Date(hoy) - new Date(ultimaFecha)) / 86400000;
      if (diff <= 1) estado = 'amarillo';
    }

    return { nombre, ultimaFecha, estado, rol: u[3] };
  });
}

// ── Dashboard ───────────────────────────────────────────────────
function obtenerDashboard(desde, hasta) {
  const tSheet = SS.getSheetByName('Tareas');
  const sSheet = SS.getSheetByName('Sesiones');
  if (!tSheet) return {};

  const tareas = tSheet.getLastRow() > 1
    ? tSheet.getDataRange().getValues().slice(1) : [];
  const sesiones = (sSheet && sSheet.getLastRow() > 1)
    ? sSheet.getDataRange().getValues().slice(1) : [];

  // Filtrar por rango de fechas si se pasa
  const filtrarFecha = (fecha) => {
    if (!desde && !hasta) return true;
    const f = typeof fecha === 'object'
      ? Utilities.formatDate(fecha, 'America/Santiago', 'yyyy-MM-dd')
      : String(fecha).substring(0, 10);
    if (desde && f < desde) return false;
    if (hasta && f > hasta) return false;
    return true;
  };

  const tareasF   = tareas.filter(t => filtrarFecha(t[8]));
  const sesionesF = sesiones.filter(s => filtrarFecha(s[3]));

  // HH por área
  const hhArea = {};
  tareasF.forEach(t => {
    const area = t[4] || 'Sin área';
    const hhr  = parseFloat(t[12]) || 0;
    hhArea[area] = (hhArea[area] || 0) + hhr;
  });

  // HH por tipo
  const hhTipo = {};
  tareasF.forEach(t => {
    const tipo = t[5] || 'Sin tipo';
    const hhr  = parseFloat(t[12]) || 0;
    hhTipo[tipo] = (hhTipo[tipo] || 0) + hhr;
  });

  // HH por persona
  const hhPersona = {};
  sesionesF.forEach(s => {
    const nom = s[2] || 'Sin nombre';
    const hh  = parseFloat(s[6]) || 0;
    hhPersona[nom] = (hhPersona[nom] || 0) + hh;
  });

  // Desviación plan vs real
  let totalPlan = 0, totalReal = 0, nConDesv = 0;
  const desviaciones = tareasF.map(t => {
    const plan = parseFloat(t[7]) || 0;
    const real = parseFloat(t[12]) || 0;
    totalPlan += plan;
    totalReal += real;
    if (plan > 0) nConDesv++;
    return { tarea: t[2], plan, real, desv: real - plan, pct: plan > 0 ? ((real-plan)/plan*100).toFixed(1) : null };
  });

  // Tareas vencidas
  const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  const vencidas = tareasF.filter(t => {
    const comp = t[9] ? (typeof t[9]==='object'
      ? Utilities.formatDate(t[9],'America/Santiago','yyyy-MM-dd')
      : String(t[9]).substring(0,10)) : '';
    return comp && comp < hoy && t[10] !== 'Listo' && t[10] !== 'Cancelado';
  }).length;

  // Estados
  const estados = {};
  tareasF.forEach(t => {
    const e = t[10] || 'Sin estado';
    estados[e] = (estados[e] || 0) + 1;
  });

  // Top tipos por HH
  const topTipos = Object.entries(hhTipo)
    .sort((a,b) => b[1]-a[1]).slice(0, 5)
    .map(([tipo, hh]) => ({ tipo, hh: Math.round(hh*10)/10 }));

  // Promedio HH por tipo (para calibrar estimaciones)
  const contTipo = {};
  tareasF.forEach(t => {
    const tipo = t[5] || 'Sin tipo';
    if (!contTipo[tipo]) contTipo[tipo] = { suma: 0, n: 0 };
    contTipo[tipo].suma += parseFloat(t[12]) || 0;
    contTipo[tipo].n++;
  });
  const promTipo = Object.entries(contTipo).map(([tipo, d]) => ({
    tipo, promHH: d.n > 0 ? Math.round(d.suma/d.n*10)/10 : 0, n: d.n
  })).sort((a,b) => b.promHH - a.promHH);

  return {
    resumen: {
      totalTareas:   tareasF.length,
      totalSesiones: sesionesF.length,
      totalHHPlan:   Math.round(totalPlan*10)/10,
      totalHHReal:   Math.round(totalReal*10)/10,
      desvTotal:     Math.round((totalReal-totalPlan)*10)/10,
      cpi:           totalPlan > 0 ? Math.round(totalReal/totalPlan*100)/100 : null,
      vencidas,
    },
    hhArea,
    hhTipo,
    hhPersona,
    estados,
    topTipos,
    promTipo,
    desviaciones: desviaciones.slice(0, 20),
  };
}

// ── Actualizar último registro ──────────────────────────────────
function actualizarUltimoRegistro(usuario) {
  const sheet = SS.getSheetByName('Usuarios');
  if (!sheet) return { ok: false };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === usuario) {
      const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
      sheet.getRange(i + 1, 5).setValue(hoy);
      return { ok: true };
    }
  }
  return { ok: false };
}


// ── Eliminar sesión por TareaID+Fecha+Inicio ───────────────────────────
function eliminarSesion(tareaId, fecha, inicio) {
  const sheet = SS.getSheetByName('Sesiones');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: false };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowTid   = String(data[i][0]);
    const rowFecha = typeof data[i][3] === 'object'
      ? Utilities.formatDate(data[i][3], 'America/Santiago', 'yyyy-MM-dd')
      : String(data[i][3]).substring(0, 10);
    const rowIni = String(data[i][4]);
    if (rowTid === String(tareaId) && rowFecha === String(fecha) && rowIni === String(inicio)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Sesión no encontrada' };
}

// ── Eliminar todas las sesiones de una tarea ───────────────────────────
function eliminarSesionesPorTarea(tareaId) {
  const sheet = SS.getSheetByName('Sesiones');
  if (!sheet || sheet.getLastRow() <= 1) return { ok: true };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(tareaId)) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// ── Actualizar sesión por TareaID+Fecha+Inicio ────────────────────────
function actualizarSesion(tareaId, fechaOrig, inicioOrig, campos) {
  const sheet = SS.getSheetByName('Sesiones');
  if (!sheet) return { ok: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowTid   = String(data[i][0]);
    const rowFecha = typeof data[i][3] === 'object'
      ? Utilities.formatDate(data[i][3], 'America/Santiago', 'yyyy-MM-dd')
      : String(data[i][3]).substring(0, 10);
    const rowIni = String(data[i][4]);
    if (rowTid === String(tareaId) && rowFecha === String(fechaOrig) && rowIni === String(inicioOrig)) {
      if (campos.fecha)       sheet.getRange(i+1, 4).setValue(campos.fecha);
      if (campos.responsable) sheet.getRange(i+1, 3).setValue(campos.responsable);
      if (campos.inicio)      sheet.getRange(i+1, 5).setValue(campos.inicio);
      if (campos.termino)     sheet.getRange(i+1, 6).setValue(campos.termino);
      if (campos.hh != null)  sheet.getRange(i+1, 7).setValue(campos.hh);
      if (campos.comentario != null) sheet.getRange(i+1, 8).setValue(campos.comentario);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Sesión no encontrada' };
}

// ── Helper JSON response ────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

