// ═══════════════════════════════════════════════════════════════
// APPS SCRIPT — SISTEMA OT PATAGONIA
// Entidad única: Bloques (Express o Simple)
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// Índices de columnas en la hoja Bloques (0-based)
const B = {
  ID:0, TIPO_INGRESO:1, TIPO_BLOQUE:2, TITULO:3, AREA:4,
  SOLICITANTE:5, RESPONSABLE:6, URGENTE:7, IMPORTANTE:8, COMENTARIO:9,
  ESTADO:10, FECHA_REAL:11, HORA_INICIO:12, HORA_FIN:13, HH_REAL:14,
  HH_PLAN:15, FECHA_PROG:16, HORA_INICIO_PROG:17, HORA_FIN_PROG:18, FECHA_INGRESO:19
};

const HEADERS_BLOQUES = [
  'ID','Tipo_Ingreso','Tipo_Bloque','Titulo','Area','Solicitante','Responsable',
  'Urgente','Importante','Comentario','Estado','Fecha_Real','Hora_Inicio','Hora_Fin',
  'HH_Real','HH_Plan','Fecha_Programada','Hora_Inicio_Prog','Hora_Fin_Prog','FechaIngreso'
];

// ── Inicializar headers si la hoja está vacía ───────────────────
function inicializar() {
  const sh = SS.getSheetByName('Bloques');
  if (sh && sh.getLastRow() === 0) sh.appendRow(HEADERS_BLOQUES);
}

// ── Helpers ─────────────────────────────────────────────────────
function fmtFecha(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Santiago', 'yyyy-MM-dd');
  return String(v).substring(0, 10);
}

function bloqueToObj(row) {
  return {
    id:             row[B.ID],
    tipoIngreso:    row[B.TIPO_INGRESO],
    tipoBloque:     row[B.TIPO_BLOQUE],
    titulo:         row[B.TITULO],
    area:           row[B.AREA],
    solicitante:    row[B.SOLICITANTE],
    responsable:    row[B.RESPONSABLE],
    urgente:        row[B.URGENTE],
    importante:     row[B.IMPORTANTE],
    comentario:     row[B.COMENTARIO],
    estado:         row[B.ESTADO],
    fechaReal:      fmtFecha(row[B.FECHA_REAL]),
    horaInicio:     row[B.HORA_INICIO] || '',
    horaFin:        row[B.HORA_FIN] || '',
    hhReal:         parseFloat(row[B.HH_REAL]) || 0,
    hhPlan:         parseFloat(row[B.HH_PLAN]) || 0,
    fechaProg:      fmtFecha(row[B.FECHA_PROG]),
    horaInicioProg: row[B.HORA_INICIO_PROG] || '',
    horaFinProg:    row[B.HORA_FIN_PROG] || '',
    fechaIngreso:   fmtFecha(row[B.FECHA_INGRESO]),
  };
}

// ── GET handler ─────────────────────────────────────────────────
function doGet(e) {
  inicializar();
  const accion = e.parameter.accion || 'leer';
  let resultado;

  if      (accion === 'login')            resultado = login(e.parameter.usuario, e.parameter.password);
  else if (accion === 'leer')             resultado = leerBloques();
  else if (accion === 'listas')           resultado = obtenerListas();
  else if (accion === 'cumplimiento')     resultado = obtenerCumplimiento();
  else if (accion === 'dashboard')        resultado = obtenerDashboard(e.parameter.desde, e.parameter.hasta);
  else if (accion === 'reprogramaciones') resultado = obtenerReprogramaciones(e.parameter.desde, e.parameter.hasta);
  else resultado = { error: 'Acción no reconocida' };

  return jsonResponse(resultado);
}

// ── POST handler ────────────────────────────────────────────────
function doPost(e) {
  inicializar();
  const data = JSON.parse(e.postData.contents);
  const accion = data.accion;
  let resultado;

  if      (accion === 'insertar')                 resultado = insertarBloque(data.bloque);
  else if (accion === 'actualizar')               resultado = actualizarBloque(data.id, data.campos);
  else if (accion === 'eliminar')                 resultado = eliminarBloque(data.id);
  else if (accion === 'reprogramar')              resultado = registrarReprogramacion(data);
  else if (accion === 'guardarLista')             resultado = guardarLista(data.items);
  else if (accion === 'actualizarUltimoRegistro') resultado = actualizarUltimoRegistro(data.usuario);
  else if (accion === 'gestionUsuario')           resultado = gestionUsuario(data);
  else resultado = { error: 'Acción no reconocida' };

  return jsonResponse(resultado);
}

// ── CRUD Bloques ────────────────────────────────────────────────
function leerBloques() {
  const sh = SS.getSheetByName('Bloques');
  if (!sh || sh.getLastRow() <= 1) return [];
  return sh.getDataRange().getValues().slice(1).map(bloqueToObj);
}

function insertarBloque(b) {
  const sh = SS.getSheetByName('Bloques');
  if (!sh) return { ok: false, error: 'Hoja Bloques no encontrada' };

  const ultimaFila = sh.getLastRow();
  const ultimoId = ultimaFila > 1
    ? parseInt(sh.getRange(ultimaFila, 1).getValue()) || 0
    : 0;
  const nuevoId = ultimoId + 1;

  sh.appendRow([
    nuevoId,
    b.tipoIngreso,
    b.tipoBloque,
    b.titulo,
    b.area,
    b.solicitante,
    b.responsable,
    b.urgente,
    b.importante,
    b.comentario || '',
    b.estado,
    b.fechaReal       || '',
    b.horaInicio      || '',
    b.horaFin         || '',
    b.hhReal          || 0,
    b.hhPlan          || 0,
    b.fechaProg       || '',
    b.horaInicioProg  || '',
    b.horaFinProg     || '',
    Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'),
  ]);

  return { ok: true, id: nuevoId };
}

function actualizarBloque(id, campos) {
  const sh = SS.getSheetByName('Bloques');
  if (!sh) return { ok: false };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.entries(campos).forEach(([col, val]) => {
        sh.getRange(i + 1, parseInt(col) + 1).setValue(val);
      });
      return { ok: true };
    }
  }
  return { ok: false, error: 'Bloque no encontrado' };
}

function eliminarBloque(id) {
  const sh = SS.getSheetByName('Bloques');
  if (!sh) return { ok: false };
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Bloque no encontrado' };
}

// ── Login ───────────────────────────────────────────────────────
function login(usuario, password) {
  const sh = SS.getSheetByName('Usuarios');
  if (!sh) return { ok: false, error: 'Sin hoja Usuarios' };
  const u = sh.getDataRange().getValues().slice(1)
    .find(r => r[0] === usuario && r[2] === password);
  if (!u) return { ok: false, error: 'Usuario o contraseña incorrectos' };
  return { ok: true, usuario: u[0], nombre: u[1], rol: u[3] };
}

// ── Listas (dropdowns) ──────────────────────────────────────────
function obtenerListas() {
  const sh = SS.getSheetByName('Listas');
  if (!sh || sh.getLastRow() <= 1) return {};
  const result = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (!result[r[0]]) result[r[0]] = [];
    result[r[0]].push(r[1]);
  });
  return result;
}

function guardarLista(items) {
  const sh = SS.getSheetByName('Listas');
  if (!sh) return { ok: false };
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  items.forEach(item => sh.appendRow([item.tipo, item.valor]));
  return { ok: true };
}

// ── Cumplimiento diario ─────────────────────────────────────────
function obtenerCumplimiento() {
  const uSheet = SS.getSheetByName('Usuarios');
  const bSheet = SS.getSheetByName('Bloques');
  if (!uSheet) return [];

  const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  const bloques = (bSheet && bSheet.getLastRow() > 1)
    ? bSheet.getDataRange().getValues().slice(1) : [];

  return uSheet.getDataRange().getValues().slice(1).map(u => {
    const nombre = u[1];
    const confirmados = bloques.filter(
      b => b[B.RESPONSABLE] === nombre && b[B.ESTADO] === 'Confirmado'
    );
    let ultimaFecha = '';
    if (confirmados.length > 0) {
      const fechas = confirmados
        .map(b => fmtFecha(b[B.FECHA_REAL]))
        .filter(f => f)
        .sort()
        .reverse();
      ultimaFecha = fechas[0] || '';
    }

    let estado = 'rojo';
    if (ultimaFecha === hoy) {
      estado = 'verde';
    } else if (ultimaFecha) {
      const diff = (new Date(hoy) - new Date(ultimaFecha)) / 86400000;
      if (diff <= 1) estado = 'amarillo';
    }

    return { nombre, ultimaFecha, estado, rol: u[3] };
  });
}

// ── Dashboard ───────────────────────────────────────────────────
function obtenerDashboard(desde, hasta) {
  const sh = SS.getSheetByName('Bloques');
  if (!sh || sh.getLastRow() <= 1) return {};

  const filtrar = (fecha) => {
    if (!desde && !hasta) return true;
    const f = fmtFecha(fecha);
    if (desde && f < desde) return false;
    if (hasta && f > hasta) return false;
    return true;
  };

  const todos = sh.getDataRange().getValues().slice(1)
    .filter(b => filtrar(b[B.FECHA_REAL]) || (!b[B.FECHA_REAL] && filtrar(b[B.FECHA_INGRESO])));

  const confirmados = todos.filter(b => b[B.ESTADO] === 'Confirmado');

  // HH Real por área
  const hhArea = {};
  confirmados.forEach(b => {
    const area = b[B.AREA] || 'Sin área';
    hhArea[area] = (hhArea[area] || 0) + (parseFloat(b[B.HH_REAL]) || 0);
  });

  // HH Real por tipo de bloque
  const hhTipo = {};
  confirmados.forEach(b => {
    const tipo = b[B.TIPO_BLOQUE] || 'Sin tipo';
    hhTipo[tipo] = (hhTipo[tipo] || 0) + (parseFloat(b[B.HH_REAL]) || 0);
  });

  // HH Real por responsable
  const hhPersona = {};
  confirmados.forEach(b => {
    const nom = b[B.RESPONSABLE] || 'Sin asignar';
    hhPersona[nom] = (hhPersona[nom] || 0) + (parseFloat(b[B.HH_REAL]) || 0);
  });

  // Conteos Express vs Simple
  const nExpress = confirmados.filter(b => b[B.TIPO_INGRESO] === 'Express').length;
  const nSimple  = confirmados.filter(b => b[B.TIPO_INGRESO] === 'Simple').length;
  const nUrgente = confirmados.filter(b => b[B.URGENTE] === 'Si').length;

  // HH Express vs Simple (para el argumento clave ante gerencia)
  const hhExpress = confirmados
    .filter(b => b[B.TIPO_INGRESO] === 'Express')
    .reduce((s, b) => s + (parseFloat(b[B.HH_REAL]) || 0), 0);
  const hhSimple = confirmados
    .filter(b => b[B.TIPO_INGRESO] === 'Simple')
    .reduce((s, b) => s + (parseFloat(b[B.HH_REAL]) || 0), 0);

  // Estados de todos los bloques
  const estados = {};
  todos.forEach(b => {
    const e = b[B.ESTADO] || 'Sin estado';
    estados[e] = (estados[e] || 0) + 1;
  });

  const totalHHReal = confirmados.reduce((s, b) => s + (parseFloat(b[B.HH_REAL]) || 0), 0);
  const totalHHPlan = confirmados.reduce((s, b) => s + (parseFloat(b[B.HH_PLAN]) || 0), 0);

  return {
    resumen: {
      totalBloques:  todos.length,
      confirmados:   confirmados.length,
      nExpress,
      nSimple,
      nUrgente,
      pctUrgente:    confirmados.length > 0 ? Math.round(nUrgente / confirmados.length * 100) : 0,
      hhExpress:     Math.round(hhExpress * 10) / 10,
      hhSimple:      Math.round(hhSimple * 10) / 10,
      totalHHReal:   Math.round(totalHHReal * 10) / 10,
      totalHHPlan:   Math.round(totalHHPlan * 10) / 10,
    },
    estados,
    hhArea,
    hhTipo,
    hhPersona,
  };
}

// ── Reprogramaciones ────────────────────────────────────────────
function registrarReprogramacion(data) {
  let sh = SS.getSheetByName('Reprogramaciones');
  if (!sh) {
    sh = SS.insertSheet('Reprogramaciones');
    sh.appendRow(['BloqueID','Titulo','Responsable','FechaOriginal','HoraOriginalInicio',
                  'HoraOriginalFin','FechaNueva','HoraNuevaInicio','HoraNuevaFin',
                  'Motivo','EsUrgencia','ReprogramadoPor','FechaRegistro']);
  }
  const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  sh.appendRow([
    data.bloqueId, data.titulo, data.responsable,
    data.fechaOriginal, data.horaInicioOriginal || '', data.horaFinOriginal || '',
    data.fechaNueva,    data.horaInicioNueva   || '', data.horaFinNueva    || '',
    data.motivo, data.esUrgencia ? 'Sí' : 'No', data.reprogramadoPor, hoy,
  ]);

  // Actualizar bloque: nueva fecha/hora programada y estado Programado
  actualizarBloque(data.bloqueId, {
    [B.ESTADO]:          'Programado',
    [B.FECHA_PROG]:      data.fechaNueva,
    [B.HORA_INICIO_PROG]: data.horaInicioNueva || '',
    [B.HORA_FIN_PROG]:   data.horaFinNueva    || '',
  });

  return { ok: true };
}

function obtenerReprogramaciones(desde, hasta) {
  const sh = SS.getSheetByName('Reprogramaciones');
  if (!sh || sh.getLastRow() <= 1) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => {
    const f = fmtFecha(r[12]);
    if (desde && f < desde) return false;
    if (hasta && f > hasta) return false;
    return true;
  }).map(r => ({
    bloqueId:         r[0],
    titulo:           r[1],
    responsable:      r[2],
    fechaOriginal:    fmtFecha(r[3]),
    horaInicioOrig:   r[4],
    horaFinOrig:      r[5],
    fechaNueva:       fmtFecha(r[6]),
    horaInicioNueva:  r[7],
    horaFinNueva:     r[8],
    motivo:           r[9],
    esUrgencia:       r[10],
    reprogPor:        r[11],
    fechaReg:         fmtFecha(r[12]),
  }));
}

// ── Actualizar último registro ──────────────────────────────────
function actualizarUltimoRegistro(usuario) {
  const sh = SS.getSheetByName('Usuarios');
  if (!sh) return { ok: false };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === usuario) {
      sh.getRange(i + 1, 5).setValue(
        Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd')
      );
      return { ok: true };
    }
  }
  return { ok: false };
}

// ── Gestión de usuarios ─────────────────────────────────────────
function gestionUsuario(data) {
  const sh = SS.getSheetByName('Usuarios');
  if (!sh) return { ok: false };
  const rows = sh.getDataRange().getValues();

  if (!data.usuarioOrig) {
    sh.appendRow([data.usuario, data.nombre, data.pass, data.rol, '']);
    return { ok: true };
  }

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.usuarioOrig) {
      sh.getRange(i+1, 1).setValue(data.usuario);
      sh.getRange(i+1, 2).setValue(data.nombre);
      if (data.pass) sh.getRange(i+1, 3).setValue(data.pass);
      sh.getRange(i+1, 4).setValue(data.rol);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Usuario no encontrado' };
}

// ── Helper JSON ─────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
