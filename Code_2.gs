// ═══════════════════════════════════════════════════════════════
// APPS SCRIPT — SISTEMA OT PATAGONIA
// Maneja: Tareas, Sesiones, Usuarios, Listas, Config
// ═══════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Encabezados por hoja ────────────────────────────────────────
const HEADERS = {
  Tareas:   ['ID','Nombre','Tarea','Solicitante','Area','Tipo','Urgente','Importante',
             'HH_Plan','FechaIngreso','FechaComp','Estado','Comentario','HH_Real','N_Sesiones','FechaEjecucion'],
  Sesiones: ['TareaID','TareaNombre','Responsable','Fecha',
             'HoraInicio','HoraTermino','HH_Sesion','Comentario','FechaRegistro','EsExtra','MotivoExtra'],
  Usuarios: ['Usuario','Nombre','Password','Rol','UltimoRegistro'],
  Listas:   ['Tipo','Valor'],
  Config:   ['Clave','Valor'],
  Reprogramaciones: ['TareaID','TareaNombre','Responsable','FechaOriginal',
                     'FechaNueva','Motivo','EsUrgencia','ReprogramadoPor','FechaRegistro'],
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
  ['MotivoReprog','Tarea urgente interrumpió'],
  ['MotivoReprog','Cliente solicitó cambio'],
  ['MotivoReprog','Error de estimación'],
  ['MotivoReprog','Recursos no disponibles'],
  ['MotivoReprog','Cambio de prioridad'],
  ['MotivoReprog','Otro'],
  ['MotivoExtra','Cumplir plazo comprometido'],
  ['MotivoExtra','Urgencia de cliente'],
  ['MotivoExtra','Retraso acumulado'],
  ['MotivoExtra','Solicitud de gerencia'],
  ['MotivoExtra','Otro'],
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
  } else if (accion === 'gestionUsuario') {
    resultado = gestionUsuario(data);
  } else if (accion === 'reprogramar') {
    resultado = registrarReprogramacion(data);
  } else if (accion === 'eliminarSesion') {
    resultado = eliminarSesion(data.tareaId, data.fecha, data.inicio);
  } else if (accion === 'eliminarPorTareaId') {
    resultado = eliminarSesionesPorTarea(data.id);
  } else if (accion === 'actualizarSesion') {
    resultado = actualizarSesion(data.tareaId, data.fechaOrig, data.inicioOrig, data.campos);
  } else if (accion === 'cargaTrabajo') {
    resultado = obtenerCargaTrabajo(e.parameter.periodo);
  } else if (accion === 'kpiProductividad') {
    resultado = obtenerKPIProductividad(e.parameter.desde, e.parameter.hasta);
  } else if (accion === 'reprogramaciones') {
    resultado = obtenerReprogramaciones(e.parameter.desde, e.parameter.hasta);
  } else if (accion === 'weeklyReview') {
    resultado = obtenerWeeklyReview(e.parameter.usuario);
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
  } else if (accion === 'gestionUsuario') {
    resultado = gestionUsuario(data);
  } else if (accion === 'reprogramar') {
    resultado = registrarReprogramacion(data);
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


// ── Carga de trabajo futura ────────────────────────────────────────────
function obtenerCargaTrabajo(periodo) {
  const tSheet = SS.getSheetByName('Tareas');
  if (!tSheet || tSheet.getLastRow() <= 1) return { personas: [] };

  const tareas = tSheet.getDataRange().getValues().slice(1);
  const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');

  // Work hours per day by dow
  const HRS_DIA = { 1:8, 2:9, 3:8, 4:9, 5:8 }; // Mon-Fri

  // Calculate available hours for period
  function getAvailHrs(periodo) {
    const d = new Date();
    let days = [];
    if (periodo === 'dia') {
      days = [new Date()];
    } else if (periodo === 'semana') {
      const mon = new Date(d);
      mon.setDate(d.getDate() - d.getDay() + (d.getDay()===0?-6:1));
      for (let i=0; i<5; i++) {
        const dd = new Date(mon); dd.setDate(mon.getDate()+i);
        if (Utilities.formatDate(dd,'America/Santiago','yyyy-MM-dd') >= hoy) days.push(dd);
      }
    } else { // mes
      const firstDay = new Date(d.getFullYear(), d.getMonth(), 1);
      const lastDay  = new Date(d.getFullYear(), d.getMonth()+1, 0);
      for (let dd=new Date(d); dd<=lastDay; dd.setDate(dd.getDate()+1)) {
        if (dd >= firstDay) days.push(new Date(dd));
      }
    }
    let hrs = 0;
    days.forEach(dd => { const dow = dd.getDay(); hrs += HRS_DIA[dow]||0; });
    return Math.max(hrs, 1);
  }

  const personas = ['Luis Sanchez','Luis Diaz','Yamila Silva'];
  const availHrs = getAvailHrs(periodo);

  const result = personas.map(p => {
    const pTareas = tareas.filter(t =>
      t[1] === p && !['Listo','Cancelado'].includes(t[10])
    );
    const totalPlan = pTareas.reduce((s,t) => s + (parseFloat(t[7])||0), 0);
    const totalReal = pTareas.reduce((s,t) => s + (parseFloat(t[12])||0), 0);
    const pendiente = Math.max(0, totalPlan - totalReal);
    const pct = Math.round(pendiente / availHrs * 100);
    return {
      persona: p,
      hhComprometidas: Math.round(pendiente*10)/10,
      hhDisponibles: availHrs,
      pctOcupacion: pct,
      nTareas: pTareas.length,
      tareas: pTareas.map(t => ({
        id: t[0], tarea: t[2], area: t[4], tipo: t[5],
        hhPlan: t[7], hhReal: t[12], estado: t[10],
        pendiente: Math.max(0, (parseFloat(t[7])||0)-(parseFloat(t[12])||0))
      })).filter(t => t.pendiente > 0).sort((a,b) => b.pendiente - a.pendiente)
    };
  });

  return { personas: result, hhDisponibles: availHrs, periodo };
}

// ── KPI Productividad ─────────────────────────────────────────────────
function obtenerKPIProductividad(desde, hasta) {
  const tSheet = SS.getSheetByName('Tareas');
  const sSheet = SS.getSheetByName('Sesiones');
  const rSheet = SS.getSheetByName('Reprogramaciones');
  if (!tSheet) return {};

  const tareas   = tSheet.getLastRow()>1 ? tSheet.getDataRange().getValues().slice(1) : [];
  const sesiones = sSheet&&sSheet.getLastRow()>1 ? sSheet.getDataRange().getValues().slice(1) : [];
  const reprog   = rSheet&&rSheet.getLastRow()>1 ? rSheet.getDataRange().getValues().slice(1) : [];

  const filtFecha = (f) => {
    if (!desde && !hasta) return true;
    const fd = typeof f==='object'
      ? Utilities.formatDate(f,'America/Santiago','yyyy-MM-dd')
      : String(f).substring(0,10);
    if (desde && fd < desde) return false;
    if (hasta && fd > hasta) return false;
    return true;
  };

  const HRS_DIA = {1:8,2:9,3:8,4:9,5:8};
  function diasLaborales(d1,d2) {
    let hrs=0,d=new Date(d1+'T12:00:00');
    const end=new Date(d2+'T12:00:00');
    while(d<=end){hrs+=HRS_DIA[d.getDay()]||0;d.setDate(d.getDate()+1);}
    return Math.max(hrs,1);
  }
  const hrsDisp = desde&&hasta ? diasLaborales(desde,hasta) : 40;

  const personas = ['Luis Sanchez','Luis Diaz','Yamila Silva'];
  return personas.map(p => {
    const pT  = tareas.filter(t=>t[1]===p&&filtFecha(t[8]));
    const pS  = sesiones.filter(s=>s[2]===p&&filtFecha(s[3]));
    const pR  = reprog.filter(r=>r[2]===p&&filtFecha(r[8]));
    const pRU = reprog.filter(r=>r[2]===p&&r[6]==='Sí'&&filtFecha(r[8]));

    const done  = pT.filter(t=>t[10]==='Listo').length;
    const total = pT.length;
    const hhPlan  = pT.reduce((s,t)=>s+(parseFloat(t[7])||0),0);
    const hhReal  = pS.reduce((s,s2)=>s+(parseFloat(s2[6])||0),0);
    const hhExtr  = pS.filter(s=>s[7]==='Sí').reduce((s,s2)=>s+(parseFloat(s2[6])||0),0);

    // Eficiencia estimación: avg(plan)/avg(real) for completed tasks
    const compT = pT.filter(t=>t[10]==='Listo'&&(parseFloat(t[7])||0)>0);
    const effEst = compT.length>0
      ? Math.round(compT.reduce((s,t)=>s+(parseFloat(t[12])||0)/(parseFloat(t[7])||1),0)/compT.length*100)/100
      : null;

    return {
      persona: p,
      totalTareas: total,
      completadas: done,
      tasaCompletitud: total>0 ? Math.round(done/total*100) : 0,
      hhPlan: Math.round(hhPlan*10)/10,
      hhReal: Math.round(hhReal*10)/10,
      hhExtra: Math.round(hhExtr*10)/10,
      hhDisponibles: hrsDisp,
      productividad: Math.round(hhReal/hrsDisp*100),
      eficienciaEstimacion: effEst,
      reprogramaciones: pR.length,
      interrupcionesUrgentes: pRU.length,
      tasaInterrupcion: total>0 ? Math.round(pRU.length/total*100) : 0,
    };
  });
}

// ── Reprogramaciones ──────────────────────────────────────────────────
function registrarReprogramacion(data) {
  const sheet = SS.getSheetByName('Reprogramaciones') || SS.insertSheet('Reprogramaciones');
  if (sheet.getLastRow()===0) {
    sheet.appendRow(['TareaID','TareaNombre','Responsable','FechaOriginal',
                     'FechaNueva','Motivo','EsUrgencia','ReprogramadoPor','FechaRegistro']);
  }
  const hoy = Utilities.formatDate(new Date(),'America/Santiago','yyyy-MM-dd');
  sheet.appendRow([data.tareaId, data.tareaNombre, data.responsable,
    data.fechaOriginal, data.fechaNueva, data.motivo,
    data.esUrgencia?'Sí':'No', data.reprogramadoPor, hoy]);
  // Also update task FechaEjecucion
  const tSheet = SS.getSheetByName('Tareas');
  if (tSheet) {
    const rows = tSheet.getDataRange().getValues();
    for (let i=1; i<rows.length; i++) {
      if (String(rows[i][0])===String(data.tareaId)) {
        // Update estado to Programada if moving to future
        if (data.fechaNueva > hoy) tSheet.getRange(i+1,11).setValue('Programada');
        break;
      }
    }
  }
  return { ok: true };
}

function obtenerReprogramaciones(desde, hasta) {
  const sheet = SS.getSheetByName('Reprogramaciones');
  if (!sheet||sheet.getLastRow()<=1) return [];
  const rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(r => {
    const f = typeof r[8]==='object'
      ? Utilities.formatDate(r[8],'America/Santiago','yyyy-MM-dd')
      : String(r[8]).substring(0,10);
    if (desde&&f<desde) return false;
    if (hasta&&f>hasta) return false;
    return true;
  }).map(r => ({
    tareaId:r[0], tarea:r[1], responsable:r[2],
    fechaOrig:r[3], fechaNueva:r[4], motivo:r[5],
    esUrgencia:r[6], reprogPor:r[7], fechaReg:r[8]
  }));
}

// ── Weekly Review ─────────────────────────────────────────────────────
function obtenerWeeklyReview(usuario) {
  const tSheet = SS.getSheetByName('Tareas');
  if (!tSheet||tSheet.getLastRow()<=1) return [];
  const hoy = Utilities.formatDate(new Date(),'America/Santiago','yyyy-MM-dd');
  const rows = tSheet.getDataRange().getValues().slice(1);
  return rows.filter(t => {
    if (!['Por Iniciar','Programada','Standby'].includes(t[10])) return false;
    if (usuario&&usuario!=='admin'&&t[1]!==usuario) return false;
    return true;
  }).map(t => ({
    id:t[0], nombre:t[1], tarea:t[2], area:t[4], tipo:t[5],
    hhPlan:t[7], fechaIngreso:t[8], fechaComp:t[9], estado:t[10],
    vencida: t[9]&&String(t[9]).substring(0,10)<hoy
  })).sort((a,b)=>(a.vencida?-1:1));
}



// ── Gestión de usuarios ───────────────────────────────────────────────
function gestionUsuario(data) {
  const sheet = SS.getSheetByName('Usuarios');
  if (!sheet) return { ok: false };
  const rows = sheet.getDataRange().getValues();

  if (!data.usuarioOrig) {
    // New user
    sheet.appendRow([data.usuario, data.nombre, data.pass, data.rol, '']);
    return { ok: true };
  }

  // Edit existing
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.usuarioOrig) {
      sheet.getRange(i+1, 1).setValue(data.usuario);
      sheet.getRange(i+1, 2).setValue(data.nombre);
      if (data.pass) sheet.getRange(i+1, 3).setValue(data.pass);
      sheet.getRange(i+1, 4).setValue(data.rol);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Usuario no encontrado' };
}

// ── Helper JSON response ────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

