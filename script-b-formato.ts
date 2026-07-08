// =====================================================================================
// SCRIPT B — FORMATO DEL COMPARADOR DE OFERTAS (solo estética, no toca números)
// =====================================================================================
// Se ejecuta DESPUÉS del Script A (cálculo) sobre el mismo libro. Lee la hoja oculta
// "_CONTRATO" que dejó A para saber cuántas ofertas hay, cómo se llaman las hojas y en
// qué fila/columna está cada cosa. NO recalcula nada: solo lee valores ya escritos
// cuando lo necesita para decidir un formato (p. ej. qué filas son de tipo Capitulo).
//
// Es idempotente: se puede ejecutar todas las veces que haga falta sin romper los datos
// (deshace merges, inmovilizaciones y formatos condicionales antes de reaplicarlos).
// La hoja "_CONTRATO" se deja OCULTA al terminar (no se borra, para poder volver a
// ejecutar este script sin repetir el A).
//
// -------------------------------------------------------------------------------------
// CONTRATO CON EL SCRIPT A
// -------------------------------------------------------------------------------------
// La hoja "_CONTRATO" (oculta) tiene pares clave/valor: col A = clave, col B = valor.
// Las claves están documentadas en la cabecera del Script A. Índices 0-based. Si falta
// la hoja o una clave, este script se detiene con un error claro pidiendo ejecutar A.
//
// -------------------------------------------------------------------------------------
// RESTRICCIONES DEL RUNTIME (Office Scripts desde Power Automate)
// -------------------------------------------------------------------------------------
// Runtime síncrono: nada de .find(), .reduce(), .map(), .filter(), spread (...) ni
// flechas en callbacks de array. Solo bucles `for` clásicos. Formato por RANGOS; el
// semáforo se hace con formato condicional (color scale), una sola llamada por columna.
// =====================================================================================

function main(workbook: ExcelScript.Workbook) {

  // ===================================================================================
  // ZONA DE CONFIGURACIÓN (editable): colores, formatos y estilo
  // ===================================================================================
  const cfg: CfgB = {
    hojaContrato: "_CONTRATO",       // debe coincidir con el Script A

    // Colores de relleno (hex sin #)
    colorCabecera: "1F4E79",         // azul corporativo (fila 1 de cabecera)
    colorCabecera2: "2E75B6",        // azul medio (fila 2 de cabecera)
    colorCapitulo: "BDD7EE",         // fila Tipo="Capitulo" (azul fuerte)
    colorSubcapitulo: "DEEBF7",      // fila Tipo="Subcapitulo" (azul claro)
    colorMejor: "C6EFCE",            // verde: fila de la oferta mas economica
    colorAlerta: "FFC000",           // naranja: avisos y descuadres
    colorTotal: "DDEBF7",            // filas TOTAL
    colorManual: "FFF2CC",           // amarillo suave: celdas de entrada manual (PEM)

    // Colores del semaforo (formato condicional de escala; estos SI llevan #)
    escalaVerde: "#63BE7B",          // valor mas BAJO (barato)
    escalaAmarillo: "#FFEB84",       // valores intermedios
    escalaRojo: "#F8696B",           // valor mas ALTO (caro)

    // NOTA: el grosor del borde que encuadra cada oferta se cambia en la funcion
    // encuadrar() (al final del script). No puede ir aqui: el linter de Office Scripts
    // prohibe guardar valores de la API (como ExcelScript.BorderWeight) en variables
    // u objetos ("Aliasing or assignment of Office Scripts APIs is not allowed").

    // Formatos de numero
    fmtEuro2: "#,##0.00\" €\"",      // importes con decimales
    fmtEuro0: "#,##0\" €\"",         // importes redondeados (capitulos y grupos)
    fmtPct: "0.0%",                  // porcentajes con un decimal
    fmtNum2: "#,##0.00"              // precios unitarios y mediciones
  };

  // Leer el contrato que dejo el Script A (y comprobar que es de la version esperada)
  let meta = leerContrato(workbook, cfg.hojaContrato);
  if (metaNum(meta, "VERSION") !== 2) {
    throw new Error("El contrato es de otra version (" + meta["VERSION"] + "). Ejecuta el Script A actualizado antes que este.");
  }

  formatearResumen(workbook, meta, cfg);
  formatearCapitulos(workbook, meta, cfg);
  formatearAgrupado(workbook, meta, cfg);
  formatearPartidas(workbook, meta, cfg);

  // Limpieza final: quitar hojas residuales, ocultar el contrato y activar el resumen
  limpiezaFinal(workbook, meta, cfg);
}

// =====================================================================================
// LECTURA DEL CONTRATO
// =====================================================================================
function leerContrato(wb: ExcelScript.Workbook, nombreHoja: string): Contrato {
  let ws = wb.getWorksheet(nombreHoja);
  if (!ws) {
    throw new Error("No existe la hoja \"" + nombreHoja + "\". Ejecuta primero el Script A (calculo) sobre este libro.");
  }
  let rango = ws.getUsedRange();
  if (!rango) {
    throw new Error("La hoja \"" + nombreHoja + "\" esta vacia. Ejecuta de nuevo el Script A (calculo).");
  }
  let vals = rango.getValues();
  let meta: Contrato = {};
  for (let r = 0; r < vals.length; r++) {
    let clave = String(vals[r][0]);
    if (clave !== "") { meta[clave] = String(vals[r][1]); }
  }
  return meta;
}

// Devuelve una clave numerica del contrato; error claro si falta
function metaNum(meta: Contrato, clave: string): number {
  if (meta[clave] === undefined) {
    throw new Error("Al contrato le falta la clave \"" + clave + "\". Ejecuta de nuevo el Script A (calculo): puede ser de una version anterior.");
  }
  let v = parseFloat(meta[clave]);
  if (isNaN(v)) {
    throw new Error("La clave \"" + clave + "\" del contrato no es numerica (valor: " + meta[clave] + ").");
  }
  return v;
}

// Devuelve una clave de texto del contrato; error claro si falta
function metaStr(meta: Contrato, clave: string): string {
  if (meta[clave] === undefined) {
    throw new Error("Al contrato le falta la clave \"" + clave + "\". Ejecuta de nuevo el Script A (calculo).");
  }
  return meta[clave];
}

// Obtiene una hoja por su clave del contrato; error claro si no existe
function hojaDelContrato(wb: ExcelScript.Workbook, meta: Contrato, clave: string): ExcelScript.Worksheet {
  let nombre = metaStr(meta, clave);
  let ws = wb.getWorksheet(nombre);
  if (!ws) {
    throw new Error("No existe la hoja \"" + nombre + "\" (clave " + clave + " del contrato). Ejecuta primero el Script A (calculo).");
  }
  return ws;
}

// =====================================================================================
// HOJA 1: "Resumen ofertas"
// =====================================================================================
function formatearResumen(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_RESUMEN");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "RES_NUM_COLS");
  const filaResIni = metaNum(meta, "RES_FILA_RESUMEN_INI");
  const filaAvisoIni = metaNum(meta, "RES_FILA_AVISO_INI");
  const numAvisos = metaNum(meta, "RES_NUM_AVISOS");
  const colPem = metaNum(meta, "RES_COL_PEM");
  const colPec = metaNum(meta, "RES_COL_PEC");
  const colDifEcoEur = metaNum(meta, "RES_COL_DIF_ECO_EUR");
  const colDifEcoPct = metaNum(meta, "RES_COL_DIF_ECO_PCT");
  const colDifMedEur = metaNum(meta, "RES_COL_DIF_MED_EUR");
  const colDifMedPct = metaNum(meta, "RES_COL_DIF_MED_PCT");
  const colAviso = metaNum(meta, "RES_COL_AVISO");

  // --- Cabecera: azul corporativo, texto blanco, negrita ---
  let cab = ws.getRangeByIndexes(0, 0, 1, COLS);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("FFFFFF");
  cab.getFormat().getFill().setColor(cfg.colorCabecera);
  cab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // --- Formatos de numero (por rangos) ---
  ws.getRangeByIndexes(1, colPem, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colPec, filaResIni + 2, 1).setNumberFormat(cfg.fmtEuro2); // incluye filas de resumen
  ws.getRangeByIndexes(1, colDifEcoEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifEcoPct, N, 1).setNumberFormat(cfg.fmtPct);
  ws.getRangeByIndexes(1, colDifMedEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifMedPct, N, 1).setNumberFormat(cfg.fmtPct);

  // --- Celdas de entrada manual (PEM): fondo amarillo suave ---
  ws.getRangeByIndexes(1, colPem, N, 1).getFormat().getFill().setColor(cfg.colorManual);

  // --- Filas de resumen final en negrita ---
  ws.getRangeByIndexes(filaResIni, 0, 3, COLS).getFormat().getFont().setBold(true);

  // --- Resaltados que dependen de valores (solo LECTURA, no se recalcula nada) ---
  let vals = ws.getRangeByIndexes(1, 0, N, COLS).getValues();
  let mn = 0, hayMin = false;
  for (let i = 0; i < N; i++) {
    let v = vals[i][colPec];
    if (typeof v === "number") {
      if (!hayMin || v < mn) { mn = v; hayMin = true; }
    }
  }
  for (let i = 0; i < N; i++) {
    // Verde: la oferta mas economica
    if (hayMin && vals[i][colPec] === mn) {
      ws.getRangeByIndexes(1 + i, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorMejor);
    }
    // Naranja: ofertas con aviso de posible PEM sin convertir
    if (String(vals[i][colAviso]) !== "") {
      let c = ws.getRangeByIndexes(1 + i, colAviso, 1, 1);
      c.getFormat().getFill().setColor(cfg.colorAlerta);
      c.getFormat().getFont().setBold(true);
    }
  }

  // --- Filas de AVISO de descuadre entre hojas: naranja y negrita, que canten ---
  if (numAvisos > 0) {
    let rA = ws.getRangeByIndexes(filaAvisoIni, 0, numAvisos, COLS);
    rA.getFormat().getFill().setColor(cfg.colorAlerta);
    rA.getFormat().getFont().setBold(true);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(200);
  ws.getRangeByIndexes(0, colPem, 1, colAviso - colPem).getFormat().setColumnWidth(115);
  ws.getRangeByIndexes(0, colAviso, 1, 1).getFormat().setColumnWidth(340);

  // --- Inmovilizar: 1 fila y 1 columna ---
  fijarPaneles(ws, 1, 1);
}

// =====================================================================================
// HOJA 2: "Resumen capitulos"
// =====================================================================================
function formatearCapitulos(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_CAPITULOS");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "CAP_NUM_COLS");
  const headFilas = metaNum(meta, "CAP_HEADER_FILAS");
  const numCap = metaNum(meta, "CAP_NUM_CAPITULOS");
  const filaTotal = metaNum(meta, "CAP_FILA_TOTAL");
  const colMedia = metaNum(meta, "CAP_COL_MEDIA");
  const col1a = metaNum(meta, "CAP_COL_PRIMERA_OFERTA");
  const porOferta = metaNum(meta, "CAP_COLS_POR_OFERTA");
  const nDatos = numCap + 1; // capitulos + fila TOTAL

  formatearCabeceraDoble(ws, COLS, N, col1a, porOferta, cfg);

  // --- Formatos de numero por rangos ---
  ws.getRangeByIndexes(headFilas, colMedia, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 2, nDatos, 1).setNumberFormat(cfg.fmtPct);
  }

  // --- SEMÁFORO SOLO EN LOS DIFERENCIALES (sin fila TOTAL). Las columnas de
  //     Importe PEC se dejan en blanco, sin escala: el color va en Δ € y Δ % s/Media,
  //     con el punto amarillo anclado en 0 (la media): por debajo verde, por encima rojo.
  for (let i = 0; i < N; i++) {
    escalaDesviacion(ws, headFilas, col1a + i * porOferta + 1, numCap, cfg); // Δ € s/Media
    escalaDesviacion(ws, headFilas, col1a + i * porOferta + 2, numCap, cfg); // Δ % s/Media
  }

  // --- Fila TOTAL PEC ---
  let filaT = ws.getRangeByIndexes(filaTotal, 0, 1, COLS);
  filaT.getFormat().getFont().setBold(true);
  filaT.getFormat().getFill().setColor(cfg.colorTotal);

  // --- Bordes gruesos encuadrando el bloque de cada oferta ---
  for (let i = 0; i < N; i++) {
    encuadrar(ws, 0, col1a + i * porOferta, filaTotal + 1, porOferta);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(60);
  ws.getRangeByIndexes(0, 1, 1, 1).getFormat().setColumnWidth(230);
  ws.getRangeByIndexes(0, colMedia, 1, 1).getFormat().setColumnWidth(100);
  ws.getRangeByIndexes(0, col1a, 1, COLS - col1a).getFormat().setColumnWidth(95);

  // --- Inmovilizar: 2 filas de cabecera y 2 columnas (codigo + nombre) ---
  fijarPaneles(ws, headFilas, 2);
}

// =====================================================================================
// HOJA 3: "Resumen agrupado"
// =====================================================================================
function formatearAgrupado(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_AGRUPADO");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "AGR_NUM_COLS");
  const headFilas = metaNum(meta, "AGR_HEADER_FILAS");
  const numGrupos = metaNum(meta, "AGR_NUM_GRUPOS");
  const filaTotal = metaNum(meta, "AGR_FILA_TOTAL");
  const filaAvisoIni = metaNum(meta, "AGR_FILA_AVISO_INI");
  const numAvisos = metaNum(meta, "AGR_NUM_AVISOS");
  const colMedia = metaNum(meta, "AGR_COL_MEDIA");
  const col1a = metaNum(meta, "AGR_COL_PRIMERA_OFERTA");
  const porOferta = metaNum(meta, "AGR_COLS_POR_OFERTA");
  const nDatos = numGrupos + 1; // grupos + fila TOTAL

  formatearCabeceraDoble(ws, COLS, N, col1a, porOferta, cfg);

  // --- Formatos de numero ---
  ws.getRangeByIndexes(headFilas, colMedia, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 2, nDatos, 1).setNumberFormat(cfg.fmtPct);
  }

  // --- SEMÁFORO SOLO EN LOS DIFERENCIALES (sin fila TOTAL). Importe PEC en blanco ---
  for (let i = 0; i < N; i++) {
    escalaDesviacion(ws, headFilas, col1a + i * porOferta + 1, numGrupos, cfg);
    escalaDesviacion(ws, headFilas, col1a + i * porOferta + 2, numGrupos, cfg);
  }

  // --- Fila TOTAL ---
  let filaT = ws.getRangeByIndexes(filaTotal, 0, 1, COLS);
  filaT.getFormat().getFont().setBold(true);
  filaT.getFormat().getFill().setColor(cfg.colorTotal);

  // --- Filas de AVISO (capitulos sin grupo o duplicados): naranja y negrita ---
  if (numAvisos > 0) {
    let rA = ws.getRangeByIndexes(filaAvisoIni, 0, numAvisos, COLS);
    rA.getFormat().getFill().setColor(cfg.colorAlerta);
    rA.getFormat().getFont().setBold(true);
  }

  // --- Bordes gruesos encuadrando el bloque de cada oferta ---
  for (let i = 0; i < N; i++) {
    encuadrar(ws, 0, col1a + i * porOferta, filaTotal + 1, porOferta);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(180);
  ws.getRangeByIndexes(0, 1, 1, 1).getFormat().setColumnWidth(220);
  ws.getRangeByIndexes(0, colMedia, 1, 1).getFormat().setColumnWidth(100);
  ws.getRangeByIndexes(0, col1a, 1, COLS - col1a).getFormat().setColumnWidth(95);

  // --- Inmovilizar: 2 filas de cabecera y 1 columna (nombre del grupo) ---
  fijarPaneles(ws, headFilas, 1);
}

// =====================================================================================
// HOJA 4: "Comparativa partidas"
// =====================================================================================
function formatearPartidas(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_PARTIDAS");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "PART_NUM_COLS");
  const headFilas = metaNum(meta, "PART_HEADER_FILAS");
  const nDatos = metaNum(meta, "PART_NUM_FILAS");
  const colsFijas = metaNum(meta, "PART_COLS_FIJAS");
  const colTipo = metaNum(meta, "PART_COL_TIPO");
  const colResumen = metaNum(meta, "PART_COL_RESUMEN");
  const colUd = metaNum(meta, "PART_COL_UD");
  const colMedicion = metaNum(meta, "PART_COL_MEDICION");
  const col1a = metaNum(meta, "PART_COL_PRIMERA_OFERTA");
  const porOferta = metaNum(meta, "PART_COLS_POR_OFERTA");
  const colBarata = metaNum(meta, "PART_COL_MAS_BARATA");
  const colDesv = metaNum(meta, "PART_COL_DESV");

  // --- Cabecera (2 filas). Deshacer merges antes de rehacerlos (idempotencia) ---
  let bloqueCab = ws.getRangeByIndexes(0, 0, headFilas, COLS);
  bloqueCab.unmerge();
  ws.getRangeByIndexes(0, 0, 1, colsFijas).merge(false);              // titulo
  for (let i = 0; i < N; i++) {                                        // una celda por empresa
    ws.getRangeByIndexes(0, col1a + i * porOferta, 1, porOferta).merge(false);
  }
  ws.getRangeByIndexes(0, colBarata, 1, 2).merge(false);              // bloque COMPARACION

  let cab0 = ws.getRangeByIndexes(0, 0, 1, COLS);
  cab0.getFormat().getFont().setBold(true);
  cab0.getFormat().getFont().setColor("FFFFFF");
  cab0.getFormat().getFill().setColor(cfg.colorCabecera);
  cab0.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  let cab1 = ws.getRangeByIndexes(1, 0, 1, COLS);
  cab1.getFormat().getFont().setBold(true);
  cab1.getFormat().getFont().setColor("FFFFFF");
  cab1.getFormat().getFill().setColor(cfg.colorCabecera2);
  cab1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // --- Formatos de numero por rangos ---
  ws.getRangeByIndexes(headFilas, colMedicion, nDatos, 1).setNumberFormat(cfg.fmtNum2);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtNum2);       // P. Unit
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos, 1).setNumberFormat(cfg.fmtEuro2);  // Importe
  }
  ws.getRangeByIndexes(headFilas, colDesv, nDatos, 1).setNumberFormat(cfg.fmtPct);

  // --- SEMÁFORO en la columna de comparacion Δ %: poca desviacion en verde,
  //     mucha desviacion (partidas con gran horquilla de precios) en rojo ---
  escalaSemaforo(ws, headFilas, colDesv, nDatos, cfg);

  // --- Sombreado de filas segun la columna Tipo (una sola lectura en bloque) ---
  // Tipo="Capitulo" → azul fuerte; Tipo="Subcapitulo" → azul claro; ambas en negrita.
  let vals = ws.getRangeByIndexes(headFilas, 0, nDatos, COLS).getValues();
  for (let r = 0; r < nDatos; r++) {
    let tipo = String(vals[r][colTipo]);
    if (tipo === "Capitulo" || tipo === "Subcapitulo") {
      let fila = ws.getRangeByIndexes(headFilas + r, 0, 1, COLS);
      fila.getFormat().getFill().setColor(tipo === "Capitulo" ? cfg.colorCapitulo : cfg.colorSubcapitulo);
      fila.getFormat().getFont().setBold(true);
    }
  }

  // --- Bordes gruesos encuadrando el bloque (P. Unit + Importe) de cada oferta ---
  for (let i = 0; i < N; i++) {
    encuadrar(ws, 0, col1a + i * porOferta, headFilas + nDatos, porOferta);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().setColumnWidth(85);     // por defecto
  ws.getRangeByIndexes(0, colTipo, 1, 1).getFormat().setColumnWidth(85);
  ws.getRangeByIndexes(0, colResumen, 1, 1).getFormat().setColumnWidth(330);
  ws.getRangeByIndexes(0, colUd, 1, 1).getFormat().setColumnWidth(45);
  ws.getRangeByIndexes(0, colMedicion, 1, 1).getFormat().setColumnWidth(90);
  ws.getRangeByIndexes(0, colBarata, 1, 1).getFormat().setColumnWidth(100);

  // --- Inmovilizar: 2 filas de cabecera y las columnas fijas hasta Medicion ---
  fijarPaneles(ws, headFilas, colsFijas);
}

// =====================================================================================
// LIMPIEZA FINAL
// =====================================================================================
function limpiezaFinal(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  // Borrar hojas residuales tipicas de un libro recien creado
  let residuales: string[] = ["Sheet1", "Hoja1"];
  for (let i = 0; i < residuales.length; i++) {
    let s = wb.getWorksheet(residuales[i]);
    if (s && wb.getWorksheets().length > 1) { s.delete(); }
  }
  // El contrato se queda OCULTO (no se borra: permite re-ejecutar B sin repetir A)
  let c = wb.getWorksheet(cfg.hojaContrato);
  if (c) { c.setVisibility(ExcelScript.SheetVisibility.hidden); }
  // Dejar activa la hoja de resumen
  let rs = wb.getWorksheet(metaStr(meta, "HOJA_RESUMEN"));
  if (rs) { rs.activate(); }
}

// =====================================================================================
// AYUDANTES DE FORMATO
// =====================================================================================

// Cabecera doble estandar (hojas de capitulos y agrupado): fila 0 con el nombre de cada
// constructora combinado sobre sus columnas, fila 1 con los titulos de columna.
function formatearCabeceraDoble(ws: ExcelScript.Worksheet, COLS: number, N: number,
  col1a: number, porOferta: number, cfg: CfgB) {
  let bloqueCab = ws.getRangeByIndexes(0, 0, 2, COLS);
  bloqueCab.unmerge(); // idempotencia: deshacer merges de ejecuciones anteriores
  bloqueCab.getFormat().getFont().setBold(true);
  bloqueCab.getFormat().getFont().setColor("FFFFFF");
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorCabecera);
  ws.getRangeByIndexes(1, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorCabecera2);
  bloqueCab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(0, col1a + i * porOferta, 1, porOferta).merge(false);
  }
}

// Borra TODOS los formatos condicionales de la hoja (idempotencia: si B se ejecuta dos
// veces, las escalas no se acumulan).
function limpiarCondicionales(ws: ExcelScript.Worksheet) {
  let ur = ws.getUsedRange();
  if (ur) { ur.clearAllConditionalFormats(); }
}

// SEMÁFORO estandar: escala de 3 colores sobre UNA columna. El valor mas bajo en verde,
// el mas alto en rojo, el percentil 50 en amarillo. Una sola llamada por columna.
// OJO: el linter de Office Scripts prohibe guardar objetos de su API en variables
// (aliasing), asi que TODO va encadenado en una sola expresion, sin variables intermedias.
function escalaSemaforo(ws: ExcelScript.Worksheet, filaIni: number, col: number,
  numFilas: number, cfg: CfgB) {
  if (numFilas <= 0) { return; }
  ws.getRangeByIndexes(filaIni, col, numFilas, 1)
    .addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale)
    .getColorScale()
    .setCriteria({
      minimum: { color: cfg.escalaVerde, type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue },
      midpoint: { color: cfg.escalaAmarillo, formula: "50", type: ExcelScript.ConditionalFormatColorCriterionType.percentile },
      maximum: { color: cfg.escalaRojo, type: ExcelScript.ConditionalFormatColorCriterionType.highestValue }
    });
}

// SEMÁFORO para columnas de DESVIACIÓN: el punto amarillo se ancla en el valor 0
// (la media), de modo que lo negativo (por debajo de media) tira a verde y lo positivo
// (por encima de media) tira a rojo.
// Igual que arriba: sin variables intermedias para no violar la regla de aliasing.
function escalaDesviacion(ws: ExcelScript.Worksheet, filaIni: number, col: number,
  numFilas: number, cfg: CfgB) {
  if (numFilas <= 0) { return; }
  ws.getRangeByIndexes(filaIni, col, numFilas, 1)
    .addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale)
    .getColorScale()
    .setCriteria({
      minimum: { color: cfg.escalaVerde, type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue },
      midpoint: { color: cfg.escalaAmarillo, formula: "0", type: ExcelScript.ConditionalFormatColorCriterionType.number },
      maximum: { color: cfg.escalaRojo, type: ExcelScript.ConditionalFormatColorCriterionType.highestValue }
    });
}

// Encuadra un bloque de columnas con borde grueso por los cuatro lados, para separar
// visualmente las columnas de cada constructora.
// GROSOR DEL BORDE: para cambiarlo, sustituye "thick" por "medium" (o "thin") en las
// cuatro lineas setWeight de abajo.
// OJO: los valores de la API (ExcelScript.BorderIndex..., ExcelScript.BorderWeight...)
// NO pueden guardarse en variables, arrays ni parametros: el linter de Office Scripts
// lo prohibe ("Aliasing or assignment of Office Scripts APIs is not allowed"). Por eso
// los cuatro lados van escritos uno a uno, con el enum en la propia llamada.
function encuadrar(ws: ExcelScript.Worksheet, filaIni: number, colIni: number,
  numFilas: number, numCols: number) {
  let f = ws.getRangeByIndexes(filaIni, colIni, numFilas, numCols).getFormat();
  f.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thick);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thick);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thick);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
  f.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thick);
}

// =====================================================================================
// INMOVILIZACIÓN DE PANELES — POR HOJA, SIN CONTAMINAR OTRAS HOJAS
// =====================================================================================
// CLAVE del fallo que habia antes: la inmovilizacion debe hacerse con getFreezePanes()
// de la MISMA hoja y con un rango creado desde esa MISMA hoja (ws.getRangeByIndexes).
// Ademas se hace unfreeze() antes, para que las ejecuciones repetidas no acumulen
// estados raros. Asi la inmovilizacion de una hoja jamas afecta a otra.
function fijarPaneles(ws: ExcelScript.Worksheet, filas: number, columnas: number) {
  let fp = ws.getFreezePanes();
  fp.unfreeze();
  if (filas > 0 && columnas > 0) {
    // freezeAt congela las filas y columnas que cubre el rango indicado (esquina sup-izq)
    fp.freezeAt(ws.getRangeByIndexes(0, 0, filas, columnas));
  } else if (filas > 0) {
    fp.freezeRows(filas);
  } else if (columnas > 0) {
    fp.freezeColumns(columnas);
  }
}

// =====================================================================================
// INTERFACES (tipado estricto, sin any)
// =====================================================================================
interface Contrato { [clave: string]: string; }
interface CfgB {
  hojaContrato: string;
  colorCabecera: string;
  colorCabecera2: string;
  colorCapitulo: string;
  colorSubcapitulo: string;
  colorMejor: string;
  colorAlerta: string;
  colorTotal: string;
  colorManual: string;
  escalaVerde: string;
  escalaAmarillo: string;
  escalaRojo: string;
  fmtEuro2: string;
  fmtEuro0: string;
  fmtPct: string;
  fmtNum2: string;
}
