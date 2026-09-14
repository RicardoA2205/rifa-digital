/**
 * ============================================================
 *  RIFA DIGITAL — Backend (Google Apps Script)  ·  v3
 *  Validación MANUAL de pagos Yape desde el panel admin.
 *
 *  Hojas: Config | Numeros | Ordenes | Log | Sorteo
 *  Rutas: ?             -> página pública de la rifa
 *         ?p=admin      -> panel del organizador (PIN)
 *         ?p=t&c=CODIGO -> verificación pública de un ticket
 *
 *  CAMBIOS v3
 *  - URL_PUBLICA: se captura sola la primera vez que alguien abre
 *    la web. Antes los correos salían con la URL /dev, que solo
 *    abre el dueño del script; a los compradores les daba error.
 *  - apiMisCompras(celular): recuperar tus números y tu ticket
 *    aunque cierres el navegador o cambies de teléfono.
 *  - POR_VALIDAR ya no expira. Antes, si no validabas en 12 h,
 *    el sistema le soltaba los números a alguien que YA te pagó.
 *  - adminAprobar verifica que los números sigan siendo de esa
 *    orden. Antes podía pisar un número ya vendido a otro.
 *  - Los conteos salen de la hoja Numeros, no de las órdenes,
 *    así "disponibles" deja de mentir.
 * ============================================================
 */

// ----------------------- CONSTANTES -----------------------

const HOJA = { CFG: 'Config', NUM: 'Numeros', ORD: 'Ordenes', LOG: 'Log', SOR: 'Sorteo' };

const N_COL = { NUMERO: 1, ESTADO: 2, ORDEN: 3, ACTUALIZADO: 4 };
const N_ANCHO = 4;

const O_COL = {
  ID: 1, REF: 2, NOMBRE: 3, CELULAR: 4, EMAIL: 5, NUMEROS: 6, CANTIDAD: 7,
  MONTO: 8, ESTADO: 9, COD_OP: 10, COMPROBANTE: 11, TICKET: 12,
  CREADO: 13, EXPIRA: 14, PAGADO: 15, VALIDADO_POR: 16, NOTAS: 17
};
const O_ANCHO = 17;

const EST_NUM = { DISPONIBLE: 'DISPONIBLE', RESERVADO: 'RESERVADO', PAGADO: 'PAGADO', ANULADO: 'ANULADO' };
const EST_ORD = {
  RESERVADA: 'RESERVADA',       // eligió números, aún no manda comprobante
  POR_VALIDAR: 'POR_VALIDAR',   // mandó comprobante, esperando al organizador
  PAGADA: 'PAGADA',
  RECHAZADA: 'RECHAZADA',
  EXPIRADA: 'EXPIRADA'
};

const CFG_DEFAULT = {
  RIFA_NOMBRE: 'Rifa 300',
  PREMIO: '300',
  PRECIO: '5',
  TOTAL_NUMEROS: '300',
  MIN_RESERVA: '20',
  MAX_POR_ORDEN: '10',
  YAPE_NOMBRE: 'Tu Nombre A.',
  YAPE_CELULAR: '999999999',
  FECHA_SORTEO: 'Sábado 3 de octubre',
  VENTAS_ABIERTAS: 'SI',
  CARPETA_DRIVE: '',
  URL_PUBLICA: '',              // se llena sola al primer acceso público
  MENSAJE_CIERRE: 'Las ventas están cerradas. Gracias a todos los participantes.'
};

// ----------------------- INSTALACIÓN Y MANTENIMIENTO -----------------------

function instalar() {
  const ss = _ss();

  let cfg = ss.getSheetByName(HOJA.CFG);
  if (!cfg) {
    cfg = ss.insertSheet(HOJA.CFG);
    cfg.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]).setFontWeight('bold');
    const filas = Object.keys(CFG_DEFAULT).map(k => [k, CFG_DEFAULT[k]]);
    cfg.getRange(2, 1, filas.length, 2).setValues(filas);
    cfg.setFrozenRows(1);
    cfg.setColumnWidth(1, 180);
    cfg.setColumnWidth(2, 420);
  }
  // claves nuevas en instalaciones viejas
  Object.keys(CFG_DEFAULT).forEach(k => { if (_cfg()[k] === undefined) _setCfg(k, CFG_DEFAULT[k]); });

  let num = ss.getSheetByName(HOJA.NUM);
  if (!num) {
    num = ss.insertSheet(HOJA.NUM);
    num.getRange(1, 1, 1, N_ANCHO)
      .setValues([['numero', 'estado', 'orden_id', 'actualizado']]).setFontWeight('bold');
    num.setFrozenRows(1);
  }

  const total = Number(_cfg().TOTAL_NUMEROS);
  if (num.getLastRow() < 2) {
    const filas = [];
    for (let i = 1; i <= total; i++) filas.push([_pad(i), EST_NUM.DISPONIBLE, '', '']);
    // el formato texto va ANTES de escribir, o Sheets convierte "001" en 1
    num.getRange(2, N_COL.NUMERO, filas.length, 1).setNumberFormat('@');
    num.getRange(2, 1, filas.length, N_ANCHO).setValues(filas);
  }

  let ord = ss.getSheetByName(HOJA.ORD);
  if (!ord) {
    ord = ss.insertSheet(HOJA.ORD);
    ord.getRange(1, 1, 1, O_ANCHO).setValues([[
      'orden_id', 'ref', 'nombre', 'celular', 'email', 'numeros', 'cantidad',
      'monto', 'estado', 'cod_operacion', 'comprobante_id', 'codigo_ticket',
      'creado', 'expira', 'pagado', 'validado_por', 'notas'
    ]]).setFontWeight('bold');
    ord.setFrozenRows(1);
    ord.getRange(2, O_COL.CELULAR, 5000, 1).setNumberFormat('@');
    ord.getRange(2, O_COL.NUMEROS, 5000, 1).setNumberFormat('@');
    ord.getRange(2, O_COL.COD_OP, 5000, 1).setNumberFormat('@');
  }

  let log = ss.getSheetByName(HOJA.LOG);
  if (!log) {
    log = ss.insertSheet(HOJA.LOG);
    log.getRange(1, 1, 1, 4).setValues([['fecha', 'actor', 'accion', 'detalle']]).setFontWeight('bold');
    log.setFrozenRows(1);
  }

  let sor = ss.getSheetByName(HOJA.SOR);
  if (!sor) {
    sor = ss.insertSheet(HOJA.SOR);
    sor.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]).setFontWeight('bold');
    sor.setFrozenRows(1);
    sor.setColumnWidth(2, 520);
  }

  if (!_cfg().CARPETA_DRIVE) {
    const carpeta = DriveApp.createFolder('Rifa - Comprobantes ' + new Date().getFullYear());
    _setCfg('CARPETA_DRIVE', carpeta.getId());
  }

  const props = PropertiesService.getScriptProperties();
  let pin = props.getProperty('ADMIN_PIN');
  if (!pin) {
    pin = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('ADMIN_PIN', pin);
  }

  const yaExiste = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'tareaLiberarExpiradas');
  if (!yaExiste) ScriptApp.newTrigger('tareaLiberarExpiradas').timeBased().everyMinutes(5).create();

  _log('sistema', 'INSTALAR', 'Instalación completada');
  Logger.log('LISTO. PIN de administrador: ' + pin +
    '\nPuede ser letras y números: cámbialo en Configuración del proyecto > Propiedades.' +
    '\nAhora despliega: Implementar > Nueva implementación > App web.');
  return pin;
}

/** Arregla una hoja donde Sheets convirtió "001" en 1. Seguro de repetir. */
function repararNumeros() {
  const sh = _sh(HOJA.NUM);
  const n = sh.getLastRow() - 1;
  if (n < 1) { Logger.log('No hay números que reparar.'); return 0; }
  const col = sh.getRange(2, N_COL.NUMERO, n, 1);
  const vals = col.getValues().map(f => [_pad(f[0])]);
  col.setNumberFormat('@');
  col.setValues(vals);
  SpreadsheetApp.flush();
  _log('sistema', 'REPARAR_NUMEROS', n + ' filas');
  Logger.log(n + ' números reparados. La columna A debe verse alineada a la IZQUIERDA.');
  return n;
}

/** Deja la rifa como recién instalada. Para limpiar después de probar. */
function reiniciarRifa() {
  const num = _sh(HOJA.NUM);
  const n = num.getLastRow() - 1;
  if (n > 0) {
    const filas = [];
    for (let i = 0; i < n; i++) filas.push([EST_NUM.DISPONIBLE, '', '']);
    num.getRange(2, N_COL.ESTADO, n, 3).setValues(filas);
  }
  [HOJA.ORD, HOJA.LOG, HOJA.SOR].forEach(function (h) {
    const sh = _sh(h);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  _setCfg('VENTAS_ABIERTAS', 'SI');
  SpreadsheetApp.flush();
  Logger.log('Rifa reiniciada. Los ' + n + ' números están libres otra vez.');
  return n;
}

/** Chequeo antes de vender. Léelo completo. */
function diagnostico() {
  const cfg = _cfg();
  const sh = _sh(HOJA.NUM);
  const filas = Math.max(0, sh.getLastRow() - 1);
  const muestra = filas ? sh.getRange(2, N_COL.NUMERO, Math.min(3, filas), 1).getValues() : [];
  const formatoOk = muestra.length && muestra.every(f => String(f[0]).length === 3);
  const c = _contarNumeros();

  const msg = [
    'Rifa: ' + cfg.RIFA_NOMBRE + ' | premio S/' + cfg.PREMIO + ' | precio S/' + cfg.PRECIO,
    'Números: ' + filas + ' cargados (' + c.disponibles + ' libres, ' +
      c.reservados + ' reservados, ' + c.pagados + ' vendidos)',
    'Formato de números: ' + (formatoOk ? 'OK' : 'MAL -> ejecuta repararNumeros()'),
    'Yape: ' + cfg.YAPE_NOMBRE + ' / ' + cfg.YAPE_CELULAR +
      (cfg.YAPE_CELULAR === '999999999' ? '   <-- todavía es el de ejemplo' : ''),
    'URL pública guardada: ' + (cfg.URL_PUBLICA
      ? cfg.URL_PUBLICA
      : 'VACÍA -> abre la web una vez y se guarda sola, o pégala a mano en Config'),
    'Ventas abiertas: ' + cfg.VENTAS_ABIERTAS,
    'Carpeta de comprobantes: ' + (cfg.CARPETA_DRIVE ? 'OK' : 'FALTA -> ejecuta instalar()'),
    'PIN configurado: ' + (PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') ? 'sí' : 'NO'),
    'Trigger de limpieza: ' + (ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === 'tareaLiberarExpiradas') ? 'activo' : 'FALTA')
  ].join('\n');

  Logger.log(msg);
  return msg;
}

/** Fuerza la URL pública si la automática falló. Llámala desde un envoltorio. */
function fijarUrlPublica(url) {
  const u = String(url || '').trim();
  if (u.indexOf('https://') !== 0 || u.indexOf('/exec') === -1) {
    throw new Error('Pega la URL completa de la implementación, la que termina en /exec');
  }
  _setCfg('URL_PUBLICA', u);
  _log('sistema', 'URL_PUBLICA', u);
  return u;
}

function cambiarPin(nuevo) {
  if (!nuevo || String(nuevo).length < 4) {
    throw new Error('El PIN necesita al menos 4 caracteres. Llámala desde una función envoltorio, ' +
      'o edítalo en Configuración del proyecto > Propiedades de la secuencia de comandos.');
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', String(nuevo));
  _log('sistema', 'CAMBIO_PIN', 'PIN actualizado');
  return 'PIN actualizado';
}

function verPin() {
  const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  Logger.log('PIN actual: ' + pin);
  return pin;
}

// ----------------------- RUTEO WEB -----------------------

function doGet(e) {
  _capturarUrl();

  const p = (e && e.parameter && e.parameter.p) || 'rifa';
  let archivo = 'Index', titulo = _cfg().RIFA_NOMBRE;

  if (p === 'admin') { archivo = 'Admin'; titulo = 'Panel · ' + titulo; }
  else if (p === 't') { archivo = 'Ticket'; titulo = 'Ticket · ' + titulo; }

  const t = HtmlService.createTemplateFromFile(archivo);
  t.codigo = (e && e.parameter && e.parameter.c) || '';
  t.baseUrl = _urlBase();

  return t.evaluate()
    .setTitle(titulo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/**
 * Guarda la URL /exec la primera vez que alguien entra por la web.
 * Es la única forma confiable de conocerla: llamada desde el editor o
 * desde un trigger, getUrl() devuelve la URL /dev, que solo abre el
 * dueño del script. Por eso los correos llevaban a una página muerta.
 */
function _capturarUrl() {
  try {
    if (_cfg().URL_PUBLICA) return;
    const u = ScriptApp.getService().getUrl();
    if (u && u.indexOf('/exec') !== -1) {
      _setCfg('URL_PUBLICA', u);
      _log('sistema', 'URL_PUBLICA', u);
    }
  } catch (err) { /* nunca romper la carga de la página por esto */ }
}

function _urlBase() {
  const u = _cfg().URL_PUBLICA;
  if (u && u.indexOf('http') === 0) return u;
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// ----------------------- API PÚBLICA -----------------------

function apiEstado() {
  const cfg = _cfg();
  const datos = _sh(HOJA.NUM).getDataRange().getValues();
  const numeros = [];
  let pagados = 0, reservados = 0;

  for (let i = 1; i < datos.length; i++) {
    const bruto = String(datos[i][N_COL.NUMERO - 1]).trim();
    if (!bruto) continue;
    const est = String(datos[i][N_COL.ESTADO - 1]);
    let e = 'D';
    if (est === EST_NUM.PAGADO) { e = 'P'; pagados++; }
    else if (est === EST_NUM.RESERVADO) { e = 'R'; reservados++; }
    else if (est === EST_NUM.ANULADO) { e = 'X'; }
    numeros.push({ n: _pad(bruto), e: e });
  }

  return {
    numeros: numeros,
    pagados: pagados,
    reservados: reservados,
    disponibles: numeros.length - pagados - reservados,
    recaudado: pagados * Number(cfg.PRECIO),
    cfg: {
      nombre: cfg.RIFA_NOMBRE,
      premio: Number(cfg.PREMIO),
      precio: Number(cfg.PRECIO),
      maxPorOrden: Number(cfg.MAX_POR_ORDEN),
      minReserva: Number(cfg.MIN_RESERVA),
      abiertas: cfg.VENTAS_ABIERTAS === 'SI',
      mensajeCierre: cfg.MENSAJE_CIERRE,
      yapeNombre: cfg.YAPE_NOMBRE,
      yapeCelular: cfg.YAPE_CELULAR,
      fechaSorteo: cfg.FECHA_SORTEO
    }
  };
}

/** Reserva números. Todo bajo LockService: es lo que evita la doble venta. */
function apiReservar(datos) {
  const cfg = _cfg();
  if (cfg.VENTAS_ABIERTAS !== 'SI') throw new Error('Las ventas están cerradas.');

  const nombre = _limpia(datos.nombre, 60);
  const celular = String(datos.celular || '').replace(/\D/g, '');
  const email = _limpia(datos.email, 80);
  const pedidos = (datos.numeros || []).map(n => _pad(n));

  if (nombre.length < 3) throw new Error('Escribe tu nombre y apellido.');
  if (!/^9\d{8}$/.test(celular)) throw new Error('El celular debe tener 9 dígitos y empezar con 9.');
  if (email && !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(email)) throw new Error('Revisa el correo.');
  if (!pedidos.length) throw new Error('Elige al menos un número.');
  if (pedidos.length > Number(cfg.MAX_POR_ORDEN)) {
    throw new Error('Puedes llevar hasta ' + cfg.MAX_POR_ORDEN + ' números por compra.');
  }
  if (new Set(pedidos).size !== pedidos.length) throw new Error('Hay números repetidos.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('El sistema está ocupado. Vuelve a intentar en unos segundos.');

  try {
    const shNum = _sh(HOJA.NUM);
    const filas = shNum.getDataRange().getValues();
    const indice = _indiceNumeros(filas);

    const tomados = [];
    pedidos.forEach(n => {
      const fila = indice[n];
      if (!fila) throw new Error('El número ' + n + ' no existe en esta rifa.');
      if (String(filas[fila - 1][N_COL.ESTADO - 1]) !== EST_NUM.DISPONIBLE) tomados.push(n);
    });
    if (tomados.length) {
      throw new Error('Estos números ya no están libres: ' + tomados.join(', ') +
        '. Actualiza el tablero y elige otros.');
    }

    const ahora = new Date();
    const expira = new Date(ahora.getTime() + Number(cfg.MIN_RESERVA) * 60000);
    const ordenId = 'ORD-' + ahora.getTime() + '-' + _rand(3);
    const ref = _rand(4);
    const total = pedidos.length * Number(cfg.PRECIO);

    _sh(HOJA.ORD).appendRow([
      ordenId, ref, nombre, celular, email, pedidos.join(','), pedidos.length,
      total, EST_ORD.RESERVADA, '', '', '', ahora, expira, '', '', ''
    ]);

    pedidos.forEach(n => {
      shNum.getRange(indice[n], N_COL.ESTADO, 1, 3).setValues([[EST_NUM.RESERVADO, ordenId, ahora]]);
    });

    SpreadsheetApp.flush();
    _log(celular, 'RESERVA', ref + ' | ' + pedidos.join(',') + ' | S/' + total);

    return {
      ordenId: ordenId, ref: ref, numeros: pedidos, total: total,
      expiraEn: expira.getTime(),
      yape: { nombre: cfg.YAPE_NOMBRE, celular: cfg.YAPE_CELULAR }
    };
  } finally {
    lock.releaseLock();
  }
}

/** Corrige nombre, celular o correo de una orden que todavía no está pagada. */
function apiActualizarDatos(datos) {
  const ordenId = String(datos.ordenId || '');
  const nombre = _limpia(datos.nombre, 60);
  const celular = String(datos.celular || '').replace(/\D/g, '');
  const email = _limpia(datos.email, 80);

  if (nombre.length < 3) throw new Error('Escribe tu nombre y apellido.');
  if (!/^9\d{8}$/.test(celular)) throw new Error('El celular debe tener 9 dígitos y empezar con 9.');
  if (email && !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(email)) throw new Error('Revisa el correo.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('El sistema está ocupado. Intenta de nuevo.');

  try {
    const sh = _sh(HOJA.ORD);
    const fila = _filaOrden(ordenId);
    if (!fila) throw new Error('No encontramos tu compra.');

    const estado = String(sh.getRange(fila, O_COL.ESTADO).getValue());
    if (estado === EST_ORD.PAGADA) {
      throw new Error('Tu pago ya fue confirmado. Escríbele al organizador para corregir tus datos.');
    }
    if (estado !== EST_ORD.RESERVADA && estado !== EST_ORD.POR_VALIDAR) {
      throw new Error('Esta reserva ya venció. Elige tus números otra vez.');
    }

    const antes = sh.getRange(fila, O_COL.NOMBRE, 1, 3).getValues()[0];
    sh.getRange(fila, O_COL.NOMBRE, 1, 3).setValues([[nombre, celular, email]]);
    SpreadsheetApp.flush();

    _log(celular, 'EDITAR_DATOS',
      String(sh.getRange(fila, O_COL.REF).getValue()) + ' | antes: ' + antes.join(' / '));

    return { ok: true, nombre: nombre, celular: celular, email: email };
  } finally {
    lock.releaseLock();
  }
}

/** El comprador manda su código de operación + foto del Yape. */
function apiEnviarComprobante(datos) {
  const ordenId = String(datos.ordenId || '');
  const codOp = _limpia(datos.codOp, 30);
  const archivo = datos.archivo || null;

  if (!codOp && !archivo) throw new Error('Manda el código de operación o la captura del Yape.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('El sistema está ocupado. Intenta de nuevo.');

  try {
    const sh = _sh(HOJA.ORD);
    const fila = _filaOrden(ordenId);
    if (!fila) throw new Error('No encontramos tu compra. Vuelve a elegir tus números.');

    const estado = String(sh.getRange(fila, O_COL.ESTADO).getValue());
    if (estado === EST_ORD.PAGADA) return { ok: true, yaPagada: true };
    if (estado !== EST_ORD.RESERVADA && estado !== EST_ORD.POR_VALIDAR) {
      throw new Error('Esta reserva ya venció. Elige tus números otra vez.');
    }

    let fileId = String(sh.getRange(fila, O_COL.COMPROBANTE).getValue() || '');
    if (archivo && archivo.b64) {
      const ref = String(sh.getRange(fila, O_COL.REF).getValue());
      const blob = Utilities.newBlob(
        Utilities.base64Decode(archivo.b64),
        archivo.mime || 'image/jpeg',
        'yape-' + ref + '-' + Date.now() + '.jpg'
      );
      fileId = DriveApp.getFolderById(_cfg().CARPETA_DRIVE).createFile(blob).getId();
    }

    sh.getRange(fila, O_COL.ESTADO).setValue(EST_ORD.POR_VALIDAR);
    sh.getRange(fila, O_COL.COD_OP).setValue(codOp);
    sh.getRange(fila, O_COL.COMPROBANTE).setValue(fileId);
    // Sin fecha de expiración: quien ya pagó NUNCA pierde sus números por
    // tu demora en validar. Solo tú puedes liberarlos, rechazando la orden.
    sh.getRange(fila, O_COL.EXPIRA).setValue('');
    SpreadsheetApp.flush();

    _log(String(sh.getRange(fila, O_COL.CELULAR).getValue()), 'COMPROBANTE',
      String(sh.getRange(fila, O_COL.REF).getValue()) + ' | op:' + codOp);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Recupera las compras de un celular. Es cómo alguien vuelve a encontrar
 * su ticket después de cerrar el navegador o cambiar de teléfono.
 */
function apiMisCompras(celular) {
  const cel = String(celular || '').replace(/\D/g, '');
  if (!/^9\d{8}$/.test(cel)) throw new Error('Ingresa tu celular de 9 dígitos.');

  const datos = _sh(HOJA.ORD).getDataRange().getValues();
  const salida = [];

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (String(f[O_COL.CELULAR - 1]).replace(/\D/g, '') !== cel) continue;

    const estado = String(f[O_COL.ESTADO - 1]);
    if (estado === EST_ORD.EXPIRADA) continue;      // ruido, ya soltó los números

    const ticket = String(f[O_COL.TICKET - 1]);
    salida.push({
      numeros: _listaNumeros(f[O_COL.NUMEROS - 1]),
      monto: Number(f[O_COL.MONTO - 1]),
      estado: estado,
      ticket: ticket,
      notas: String(f[O_COL.NOTAS - 1]),
      creado: f[O_COL.CREADO - 1] ? new Date(f[O_COL.CREADO - 1]).getTime() : 0
    });
  }

  salida.sort((a, b) => b.creado - a.creado);
  return { nombre: salida.length ? _nombreDe(cel) : '', compras: salida };
}

function _nombreDe(cel) {
  const datos = _sh(HOJA.ORD).getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][O_COL.CELULAR - 1]).replace(/\D/g, '') === cel) {
      return String(datos[i][O_COL.NOMBRE - 1]);
    }
  }
  return '';
}

/** Verificación pública de un ticket por su código. */
function apiTicket(codigo) {
  const cod = String(codigo || '').trim().toUpperCase();
  if (!cod) return null;
  const datos = _sh(HOJA.ORD).getDataRange().getValues();
  const cfg = _cfg();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][O_COL.TICKET - 1]).toUpperCase() === cod) {
      const f = datos[i];
      return {
        valido: String(f[O_COL.ESTADO - 1]) === EST_ORD.PAGADA,
        codigo: cod,
        nombre: String(f[O_COL.NOMBRE - 1]),
        numeros: _listaNumeros(f[O_COL.NUMEROS - 1]),
        monto: Number(f[O_COL.MONTO - 1]),
        estado: String(f[O_COL.ESTADO - 1]),
        fecha: f[O_COL.PAGADO - 1]
          ? Utilities.formatDate(new Date(f[O_COL.PAGADO - 1]), 'America/Lima', 'dd/MM/yyyy HH:mm') : '',
        rifa: cfg.RIFA_NOMBRE,
        premio: Number(cfg.PREMIO)
      };
    }
  }
  return { valido: false, codigo: cod, noExiste: true };
}

// ----------------------- API ADMIN -----------------------

function adminLogin(pin) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  Utilities.sleep(600);
  if (!real || String(pin) !== String(real)) {
    _log('desconocido', 'LOGIN_FALLIDO', '');
    throw new Error('PIN incorrecto.');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sesion_' + token, 'ok', 21600); // 6 h
  _log(_quien(), 'LOGIN', 'Sesión iniciada');
  return { token: token };
}

function _auth(token) {
  if (!token || CacheService.getScriptCache().get('sesion_' + token) !== 'ok') {
    throw new Error('Tu sesión venció. Ingresa el PIN otra vez.');
  }
  return true;
}

function adminResumen(token) {
  _auth(token);
  const cfg = _cfg();
  const datos = _sh(HOJA.ORD).getDataRange().getValues();

  const porValidar = [], recientes = [];
  let pagadas = 0, recaudado = 0, reservadas = 0;
  const ahora = Date.now();

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    const o = {
      id: String(f[O_COL.ID - 1]),
      ref: String(f[O_COL.REF - 1]),
      nombre: String(f[O_COL.NOMBRE - 1]),
      celular: String(f[O_COL.CELULAR - 1]),
      email: String(f[O_COL.EMAIL - 1]),
      numeros: _listaNumeros(f[O_COL.NUMEROS - 1]),
      cantidad: Number(f[O_COL.CANTIDAD - 1]),
      monto: Number(f[O_COL.MONTO - 1]),
      estado: String(f[O_COL.ESTADO - 1]),
      codOp: String(f[O_COL.COD_OP - 1]),
      tieneFoto: !!String(f[O_COL.COMPROBANTE - 1]),
      ticket: String(f[O_COL.TICKET - 1]),
      creado: f[O_COL.CREADO - 1] ? new Date(f[O_COL.CREADO - 1]).getTime() : 0,
      expira: f[O_COL.EXPIRA - 1] ? new Date(f[O_COL.EXPIRA - 1]).getTime() : 0,
      notas: String(f[O_COL.NOTAS - 1])
    };

    if (o.estado === EST_ORD.POR_VALIDAR) porValidar.push(o);
    if (o.estado === EST_ORD.PAGADA) { pagadas++; recaudado += o.monto; }
    if (o.estado === EST_ORD.RESERVADA && o.expira > ahora) reservadas++;
    recientes.push(o);
  }

  porValidar.sort((a, b) => a.creado - b.creado);   // los que más esperan, primero
  recientes.sort((a, b) => b.creado - a.creado);

  // Los conteos salen de la hoja Numeros: es la única fuente que no miente
  const c = _contarNumeros();

  return {
    stats: {
      total: c.total,
      pagados: c.pagados,
      reservados: c.reservados,
      disponibles: c.disponibles,
      recaudado: recaudado,
      premio: Number(cfg.PREMIO),
      utilidad: recaudado - Number(cfg.PREMIO),
      ordenesPagadas: pagadas,
      ordenesReservadas: reservadas,
      porValidar: porValidar.length,
      maximo: c.total * Number(cfg.PRECIO)
    },
    porValidar: porValidar,
    recientes: recientes.slice(0, 40),
    abiertas: cfg.VENTAS_ABIERTAS === 'SI',
    fechaSorteo: cfg.FECHA_SORTEO,
    sorteo: _sorteoEstado(),
    baseUrl: _urlBase()
  };
}

function adminComprobante(token, ordenId) {
  _auth(token);
  const fila = _filaOrden(ordenId);
  if (!fila) throw new Error('Orden no encontrada.');
  const fileId = String(_sh(HOJA.ORD).getRange(fila, O_COL.COMPROBANTE).getValue() || '');
  if (!fileId) return null;
  const blob = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/** Confirma el pago: números a PAGADO, genera el ticket y avisa. */
function adminAprobar(token, ordenId) {
  _auth(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Sistema ocupado, reintenta.');

  try {
    const sh = _sh(HOJA.ORD);
    const fila = _filaOrden(ordenId);
    if (!fila) throw new Error('Orden no encontrada.');

    const f = sh.getRange(fila, 1, 1, O_ANCHO).getValues()[0];
    if (String(f[O_COL.ESTADO - 1]) === EST_ORD.PAGADA) {
      return { ok: true, ticket: String(f[O_COL.TICKET - 1]), yaEstaba: true };
    }

    const numeros = _listaNumeros(f[O_COL.NUMEROS - 1]);
    const shNum = _sh(HOJA.NUM);
    const filasNum = shNum.getDataRange().getValues();
    const indice = _indiceNumeros(filasNum);

    // Antes de tocar nada: ¿esos números siguen siendo de esta orden?
    // Si una reserva venció y otro los compró, aprobar acá se los robaría.
    const conflictos = [];
    numeros.forEach(n => {
      const fn = indice[n];
      if (!fn) { conflictos.push(n + ' no existe'); return; }
      const est = String(filasNum[fn - 1][N_COL.ESTADO - 1]);
      const dueno = String(filasNum[fn - 1][N_COL.ORDEN - 1]);
      if (est === EST_NUM.PAGADO && dueno !== ordenId) conflictos.push(n + ' ya se vendió a otra persona');
      else if (est === EST_NUM.ANULADO) conflictos.push(n + ' está anulado');
      else if (dueno && dueno !== ordenId) conflictos.push(n + ' está reservado por otra orden');
    });
    if (conflictos.length) {
      throw new Error('No se puede confirmar: ' + conflictos.join('; ') +
        '. Devuélvele el dinero o acuerda otros números con esta persona.');
    }

    const codigoTicket = 'RIFA-' + numeros[0] + '-' + _rand(4);
    const ahora = new Date();

    numeros.forEach(n => {
      shNum.getRange(indice[n], N_COL.ESTADO, 1, 3).setValues([[EST_NUM.PAGADO, ordenId, ahora]]);
    });

    sh.getRange(fila, O_COL.ESTADO).setValue(EST_ORD.PAGADA);
    sh.getRange(fila, O_COL.TICKET).setValue(codigoTicket);
    sh.getRange(fila, O_COL.PAGADO).setValue(ahora);
    sh.getRange(fila, O_COL.VALIDADO_POR).setValue(_quien());
    sh.getRange(fila, O_COL.EXPIRA).setValue('');
    SpreadsheetApp.flush();

    _log(_quien(), 'APROBAR', String(f[O_COL.REF - 1]) + ' | ' + numeros.join(',') +
      ' | S/' + f[O_COL.MONTO - 1] + ' | ' + codigoTicket);

    const nombre = String(f[O_COL.NOMBRE - 1]);
    const email = String(f[O_COL.EMAIL - 1]);
    let correoEnviado = false;

    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Ticket confirmado · ' + numeros.join(', ') + ' · ' + _cfg().RIFA_NOMBRE,
          htmlBody: _htmlCorreo(nombre, numeros, f[O_COL.MONTO - 1], codigoTicket),
          name: _cfg().RIFA_NOMBRE
        });
        correoEnviado = true;
      } catch (err) {
        _log('sistema', 'ERROR_CORREO', email + ' | ' + err.message);
      }
    }

    const texto = 'Hola ' + nombre.split(' ')[0] + ', tu pago está confirmado.\n' +
      'Números: ' + numeros.join(', ') + '\n' +
      'Ticket: ' + codigoTicket + '\n' +
      'Sorteo: ' + _cfg().FECHA_SORTEO;

    return {
      ok: true, ticket: codigoTicket, correoEnviado: correoEnviado,
      whatsapp: 'https://wa.me/51' + String(f[O_COL.CELULAR - 1]) + '?text=' + encodeURIComponent(texto)
    };
  } finally {
    lock.releaseLock();
  }
}

function adminRechazar(token, ordenId, motivo) {
  _auth(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Sistema ocupado, reintenta.');

  try {
    const sh = _sh(HOJA.ORD);
    const fila = _filaOrden(ordenId);
    if (!fila) throw new Error('Orden no encontrada.');

    const f = sh.getRange(fila, 1, 1, O_ANCHO).getValues()[0];
    if (String(f[O_COL.ESTADO - 1]) === EST_ORD.PAGADA) {
      throw new Error('Esta orden ya está pagada. Si fue un error, corrígela en la hoja Ordenes.');
    }

    _liberarNumeros(_listaNumeros(f[O_COL.NUMEROS - 1]), ordenId);
    sh.getRange(fila, O_COL.ESTADO).setValue(EST_ORD.RECHAZADA);
    sh.getRange(fila, O_COL.NOTAS).setValue(_limpia(motivo, 200));
    sh.getRange(fila, O_COL.VALIDADO_POR).setValue(_quien());
    SpreadsheetApp.flush();

    _log(_quien(), 'RECHAZAR', String(f[O_COL.REF - 1]) + ' | ' + _limpia(motivo, 200));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function adminLiberarExpiradas(token) {
  _auth(token);
  return { ok: true, liberadas: _liberarExpiradas() };
}

function adminBuscar(token, q) {
  _auth(token);
  const t = String(q || '').trim().toLowerCase();
  if (t.length < 2) return [];
  const datos = _sh(HOJA.ORD).getDataRange().getValues();
  const res = [];

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    const blob = [
      f[O_COL.REF - 1], f[O_COL.NOMBRE - 1], f[O_COL.CELULAR - 1], f[O_COL.EMAIL - 1],
      f[O_COL.NUMEROS - 1], f[O_COL.TICKET - 1], f[O_COL.COD_OP - 1]
    ].join(' ').toLowerCase();

    if (blob.indexOf(t) !== -1) {
      res.push({
        id: String(f[O_COL.ID - 1]), ref: String(f[O_COL.REF - 1]),
        nombre: String(f[O_COL.NOMBRE - 1]), celular: String(f[O_COL.CELULAR - 1]),
        numeros: _listaNumeros(f[O_COL.NUMEROS - 1]), monto: Number(f[O_COL.MONTO - 1]),
        estado: String(f[O_COL.ESTADO - 1]), ticket: String(f[O_COL.TICKET - 1]),
        tieneFoto: !!String(f[O_COL.COMPROBANTE - 1]), codOp: String(f[O_COL.COD_OP - 1])
      });
      if (res.length >= 30) break;
    }
  }
  return res;
}

function adminNumeros(token) {
  _auth(token);
  const nums = _sh(HOJA.NUM).getDataRange().getValues();
  const ords = _sh(HOJA.ORD).getDataRange().getValues();
  const porId = {};
  for (let i = 1; i < ords.length; i++) {
    porId[String(ords[i][O_COL.ID - 1])] = {
      nombre: String(ords[i][O_COL.NOMBRE - 1]),
      celular: String(ords[i][O_COL.CELULAR - 1]),
      ref: String(ords[i][O_COL.REF - 1])
    };
  }
  const salida = [];
  for (let i = 1; i < nums.length; i++) {
    const bruto = String(nums[i][N_COL.NUMERO - 1]).trim();
    if (!bruto) continue;
    const oid = String(nums[i][N_COL.ORDEN - 1]);
    salida.push({
      n: _pad(bruto),
      estado: String(nums[i][N_COL.ESTADO - 1]),
      dueno: porId[oid] ? porId[oid].nombre : '',
      celular: porId[oid] ? porId[oid].celular : '',
      ref: porId[oid] ? porId[oid].ref : ''
    });
  }
  return salida;
}

function adminVentas(token, abrir) {
  _auth(token);
  _setCfg('VENTAS_ABIERTAS', abrir ? 'SI' : 'NO');
  _log(_quien(), abrir ? 'ABRIR_VENTAS' : 'CERRAR_VENTAS', '');
  if (!abrir) return { ok: true, sorteo: _comprometerLista() };
  return { ok: true };
}

// ----------------------- SORTEO VERIFICABLE -----------------------

function _comprometerLista() {
  const datos = _sh(HOJA.ORD).getDataRange().getValues();
  const numeros = [];
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][O_COL.ESTADO - 1]) === EST_ORD.PAGADA) {
      _listaNumeros(datos[i][O_COL.NUMEROS - 1]).forEach(n => numeros.push(n));
    }
  }
  numeros.sort();
  const lista = numeros.join(',');
  const huella = _sha256(lista);

  _setSorteo('CIERRE_FECHA', Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm:ss'));
  _setSorteo('PARTICIPANTES', String(numeros.length));
  _setSorteo('LISTA', lista);
  _setSorteo('HUELLA_SHA256', huella);
  _log(_quien(), 'COMPROMISO_LISTA', numeros.length + ' números | ' + huella);

  return { participantes: numeros.length, huella: huella, lista: lista };
}

function adminSortear(token, semilla, fuente) {
  _auth(token);
  if (_cfg().VENTAS_ABIERTAS === 'SI') throw new Error('Primero cierra las ventas.');
  if (_getSorteo('NUMERO_GANADOR')) return _sorteoEstado();   // irrepetible

  const lista = _getSorteo('LISTA');
  if (!lista) throw new Error('No hay lista comprometida. Cierra las ventas de nuevo.');

  const sem = _limpia(semilla, 120);
  if (sem.length < 3) throw new Error('Ingresa la semilla pública (ej. el resultado de La Tinka).');

  const numeros = lista.split(',');
  const firma = Utilities.computeHmacSha256Signature(sem, lista);
  let indice = 0;
  for (let i = 0; i < firma.length; i++) indice = (indice * 256 + (firma[i] & 0xFF)) % numeros.length;
  const ganador = numeros[indice];

  const ords = _sh(HOJA.ORD).getDataRange().getValues();
  let dueno = '', celular = '', ticket = '';
  for (let i = 1; i < ords.length; i++) {
    if (String(ords[i][O_COL.ESTADO - 1]) !== EST_ORD.PAGADA) continue;
    if (_listaNumeros(ords[i][O_COL.NUMEROS - 1]).indexOf(ganador) !== -1) {
      dueno = String(ords[i][O_COL.NOMBRE - 1]);
      celular = String(ords[i][O_COL.CELULAR - 1]);
      ticket = String(ords[i][O_COL.TICKET - 1]);
      break;
    }
  }

  _setSorteo('SEMILLA', sem);
  _setSorteo('SEMILLA_FUENTE', _limpia(fuente, 160));
  _setSorteo('FIRMA_HMAC', _hex(firma));
  _setSorteo('INDICE', String(indice));
  _setSorteo('NUMERO_GANADOR', ganador);
  _setSorteo('GANADOR', dueno);
  _setSorteo('GANADOR_CELULAR', celular);
  _setSorteo('GANADOR_TICKET', ticket);
  _setSorteo('SORTEO_FECHA', Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm:ss'));

  _log(_quien(), 'SORTEO', 'ganador ' + ganador + ' (' + dueno + ') | semilla: ' + sem);
  return _sorteoEstado();
}

function _sorteoEstado() {
  const sh = _sh(HOJA.SOR);
  if (!sh || sh.getLastRow() < 2) return {};
  const o = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(f => { if (f[0]) o[String(f[0])] = String(f[1]); });
  return o;
}

function adminSorteoEstado(token) { _auth(token); return _sorteoEstado(); }

// ----------------------- TAREAS AUTOMÁTICAS -----------------------

function tareaLiberarExpiradas() { _liberarExpiradas(); }

/**
 * Libera SOLO reservas sin comprobante. Una orden POR_VALIDAR nunca se
 * toca automáticamente: esa persona ya dice haber pagado y soltarle los
 * números por tu demora sería regalarle su plata a otro comprador.
 */
function _liberarExpiradas() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 0;

  try {
    const sh = _sh(HOJA.ORD);
    if (sh.getLastRow() < 2) return 0;
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, O_ANCHO).getValues();
    const ahora = Date.now();
    let contador = 0;

    datos.forEach((f, i) => {
      if (String(f[O_COL.ESTADO - 1]) !== EST_ORD.RESERVADA) return;
      const expira = f[O_COL.EXPIRA - 1] ? new Date(f[O_COL.EXPIRA - 1]).getTime() : 0;
      if (!expira || expira > ahora) return;

      _liberarNumeros(_listaNumeros(f[O_COL.NUMEROS - 1]), String(f[O_COL.ID - 1]));
      sh.getRange(i + 2, O_COL.ESTADO).setValue(EST_ORD.EXPIRADA);
      sh.getRange(i + 2, O_COL.NOTAS).setValue('Liberada automáticamente por vencimiento');
      contador++;
      _log('sistema', 'EXPIRAR', String(f[O_COL.REF - 1]) + ' | ' + f[O_COL.NUMEROS - 1]);
    });

    if (contador) SpreadsheetApp.flush();
    return contador;
  } finally {
    lock.releaseLock();
  }
}

/** Devuelve números al tablero, solo si siguen atados a esa orden. */
function _liberarNumeros(numeros, ordenId) {
  const buscados = {};
  numeros.forEach(function (x) {
    buscados[String(x).trim()] = 1;
    buscados[_pad(x)] = 1;
  });

  const sh = _sh(HOJA.NUM);
  const filas = sh.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    const n = String(filas[i][N_COL.NUMERO - 1]).trim();
    if (!buscados[n] && !buscados[_pad(n)]) continue;
    if (String(filas[i][N_COL.ORDEN - 1]) !== ordenId) continue;
    if (String(filas[i][N_COL.ESTADO - 1]) === EST_NUM.PAGADO) continue;
    sh.getRange(i + 1, N_COL.ESTADO, 1, 3).setValues([[EST_NUM.DISPONIBLE, '', new Date()]]);
  }
}

// ----------------------- UTILIDADES -----------------------

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _sh(nombre) { return _ss().getSheetByName(nombre); }

/** Conteo real, leyendo la hoja Numeros. */
function _contarNumeros() {
  const datos = _sh(HOJA.NUM).getDataRange().getValues();
  const c = { disponibles: 0, reservados: 0, pagados: 0, anulados: 0, total: 0 };
  for (let i = 1; i < datos.length; i++) {
    if (!String(datos[i][N_COL.NUMERO - 1]).trim()) continue;
    const e = String(datos[i][N_COL.ESTADO - 1]);
    if (e === EST_NUM.PAGADO) c.pagados++;
    else if (e === EST_NUM.RESERVADO) c.reservados++;
    else if (e === EST_NUM.ANULADO) c.anulados++;
    else c.disponibles++;
    c.total++;
  }
  return c;
}

/** Índice número -> fila, tolerante a que la hoja tenga 2 en vez de "002". */
function _indiceNumeros(filas) {
  const idx = {};
  for (let i = 1; i < filas.length; i++) {
    const bruto = String(filas[i][N_COL.NUMERO - 1]).trim();
    if (!bruto) continue;
    idx[bruto] = i + 1;
    idx[_pad(bruto)] = i + 1;
  }
  return idx;
}

/** "3,17,45" -> ['003','017','045'] */
function _listaNumeros(celda) {
  return String(celda == null ? '' : celda)
    .split(',').map(s => s.trim()).filter(s => s.length).map(_pad);
}

function _cfg() {
  const sh = _sh(HOJA.CFG);
  const o = Object.assign({}, CFG_DEFAULT);
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(f => { if (f[0]) o[String(f[0]).trim()] = String(f[1]).trim(); });
  return o;
}

function _setCfg(clave, valor) {
  const sh = _sh(HOJA.CFG);
  const datos = sh.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === clave) { sh.getRange(i + 1, 2).setValue(valor); return; }
  }
  sh.appendRow([clave, valor]);
}

function _setSorteo(clave, valor) {
  const sh = _sh(HOJA.SOR);
  const datos = sh.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === clave) { sh.getRange(i + 1, 2).setValue(valor); return; }
  }
  sh.appendRow([clave, valor]);
}

function _getSorteo(clave) { return _sorteoEstado()[clave] || ''; }

function _filaOrden(ordenId) {
  const sh = _sh(HOJA.ORD);
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, O_COL.ID, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(ordenId)) return i + 2;
  return 0;
}

function _log(actor, accion, detalle) {
  try { _sh(HOJA.LOG).appendRow([new Date(), actor, accion, detalle]); } catch (e) { /* no bloquear */ }
}

function _quien() {
  try { return Session.getActiveUser().getEmail() || 'admin(pin)'; } catch (e) { return 'admin(pin)'; }
}

/** "2" -> "002". Válido para números de 1 a 999. */
function _pad(n) {
  const s = String(n == null ? '' : n).replace(/\D/g, '');
  if (!s) return '';
  return ('000' + s).slice(-3);
}

function _rand(largo) {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // sin I, O, 0, 1
  let s = '';
  for (let i = 0; i < largo; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}

function _limpia(texto, max) {
  return String(texto == null ? '' : texto).replace(/[<>]/g, '').trim().slice(0, max || 100);
}

function _sha256(txt) {
  return _hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8));
}

function _hex(bytes) {
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _htmlCorreo(nombre, numeros, monto, codigo) {
  const cfg = _cfg();
  return '<div style="font-family:Arial,sans-serif;max-width:460px;color:#1b2430">' +
    '<p>Hola ' + nombre + ', tu pago está confirmado.</p>' +
    '<div style="border:2px dashed #123b63;border-radius:10px;padding:16px;margin:16px 0">' +
    '<p style="margin:0 0 10px;font-size:13px;color:#5b6677">' + cfg.RIFA_NOMBRE +
    ' · premio S/' + cfg.PREMIO + '</p>' +
    '<p style="margin:0;font-size:26px;font-weight:700;letter-spacing:2px">' + numeros.join('  ') + '</p>' +
    '<p style="margin:12px 0 0;font-size:14px">Pagado: S/' + monto + '</p>' +
    '<p style="margin:6px 0 0;font-size:14px">Código: <b>' + codigo + '</b></p>' +
    '</div>' +
    '<p style="font-size:15px">Sorteo: <b>' + cfg.FECHA_SORTEO + '</b></p>' +
    '<p style="font-size:12px;color:#5b6677">Guarda este código. Puedes consultar tus números ' +
    'cuando quieras desde la página de la rifa, con el celular que usaste al comprar.</p>' +
    '</div>';
}