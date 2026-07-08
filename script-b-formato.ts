// =====================================================================================
// SCRIPT B — FORMATO DEL COMPARADOR DE OFERTAS (solo estética, no toca números)
// =====================================================================================
// Se ejecuta DESPUÉS del Script A (cálculo) sobre el mismo libro. Lee la hoja oculta
// "_CONTRATO" que dejó A para saber cuántas ofertas hay, cómo se llaman las hojas y en
// qué fila/columna está cada cosa. NO recalcula nada: solo lee valores ya escritos
// cuando lo necesita para decidir un color (p. ej. cuál es la oferta mínima).
//
// Es idempotente: se puede ejecutar todas las veces que haga falta sin romper los datos
// (deshace merges e inmovilizaciones antes de volver a aplicarlos).
//
// -------------------------------------------------------------------------------------
// CONTRATO CON EL SCRIPT A
// -------------------------------------------------------------------------------------
// La hoja "_CONTRATO" (oculta) tiene pares clave/valor: columna A = clave, columna B =
// valor. Las claves están documentadas en la cabecera del Script A. Todos los índices
// son 0-based. Si falta la hoja o una clave, este script se detiene con un error claro
// pidiendo ejecutar primero el Script A.
//
// -------------------------------------------------------------------------------------
// RESTRICCIONES DEL RUNTIME (Office Scripts desde Power Automate)
// -------------------------------------------------------------------------------------
// Runtime síncrono: nada de .find(), .reduce(), .map(), spread (...) ni flechas en
// callbacks de array. Solo bucles `for` clásicos. El coloreo se hace por RANGOS siempre
// que se puede; el coloreo celda a celda se limita a los pocos casos que lo requieren
// (mínimos/máximos y desviaciones que superan el umbral).
// =====================================================================================

function main(workbook: ExcelScript.Workbook) {

  // ===================================================================================
  // ZONA DE CONFIGURACIÓN (editable): colores, formatos y umbrales
  // ===================================================================================
  const cfg: CfgB = {
    hojaContrato: "_CONTRATO",       // debe coincidir con el Script A

    // Colores (hex sin #)
    colorCabecera: "1F4E79",         // azul corporativo (fila 1 de cabecera)
    colorCabecera2: "2E75B6",        // azul medio (fila 2 de cabecera)
    colorCapPadre: "BDD7EE",         // fila de capitulo padre
    colorCapSub: "DEEBF7",           // fila de subcapitulo
    colorBanda: "F2F6FC",            // banda alterna por empresa
    colorMejor: "C6EFCE",            // verde: valor mas bajo
    colorPeor: "FFC7CE",             // rojo suave: valor mas alto
    colorAlerta: "FFC000",           // naranja: desviacion sobre umbral y avisos PEM
    colorTotal: "DDEBF7",            // fila TOTAL PEC
    colorManual: "FFF2CC",           // amarillo suave: celdas de entrada manual (PEM)

    // Umbral de desviacion por partida: se resalta en naranja la celda Δ % cuando la
    // oferta mas cara supera a la mas barata en mas de este valor (1.0 = +100 %).
    umbralDesviacionPartida: 1.0,

    // Formatos de numero
    fmtEuro2: "#,##0.00\" €\"",      // importes con decimales
    fmtEuro0: "#,##0\" €\"",         // importes redondeados (capitulos)
    fmtPct: "0.0%",                  // porcentajes con un decimal
    fmtNum2: "#,##0.00"              // precios unitarios y mediciones
  };

  // Leer el contrato que dejo el Script A
  let meta = leerContrato(workbook, cfg.hojaContrato);

  formatearResumen(workbook, meta, cfg);
  formatearCapitulos(workbook, meta, cfg);
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
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "RES_NUM_COLS");
  const filaResIni = metaNum(meta, "RES_FILA_RESUMEN_INI");
  const colPem = metaNum(meta, "RES_COL_PEM");
  const colPec = metaNum(meta, "RES_COL_PEC");
  const colDifEcoEur = metaNum(meta, "RES_COL_DIF_ECO_EUR");
  const colDifEcoPct = metaNum(meta, "RES_COL_DIF_ECO_PCT");
  const colDifMedEur = metaNum(meta, "RES_COL_DIF_MED_EUR");
  const colDifMedPct = metaNum(meta, "RES_COL_DIF_MED_PCT");
  const colAviso = metaNum(meta, "RES_COL_AVISO");
  const totalFilas = filaResIni + 3; // datos + blanco + 3 filas de resumen

  // --- Cabecera: azul corporativo, texto blanco, negrita ---
  let cab = ws.getRangeByIndexes(0, 0, 1, COLS);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("FFFFFF");
  cab.getFormat().getFill().setColor(cfg.colorCabecera);
  cab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // --- Formatos de numero (por rangos, no celda a celda) ---
  ws.getRangeByIndexes(1, colPem, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colPec, totalFilas - 1, 1).setNumberFormat(cfg.fmtEuro2); // incluye filas de resumen
  ws.getRangeByIndexes(1, colDifEcoEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifEcoPct, N, 1).setNumberFormat(cfg.fmtPct);
  ws.getRangeByIndexes(1, colDifMedEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifMedPct, N, 1).setNumberFormat(cfg.fmtPct);

  // --- Celdas de entrada manual (PEM): fondo amarillo suave para que se vean ---
  ws.getRangeByIndexes(1, colPem, N, 1).getFormat().getFill().setColor(cfg.colorManual);

  // --- Filas de resumen final en negrita ---
  ws.getRangeByIndexes(filaResIni, 0, 3, COLS).getFormat().getFont().setBold(true);

  // --- Resaltados que dependen de los valores (solo LECTURA, no se recalcula nada) ---
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
    let aviso = String(vals[i][colAviso]);
    if (aviso !== "") {
      let c = ws.getRangeByIndexes(1 + i, colAviso, 1, 1);
      c.getFormat().getFill().setColor(cfg.colorAlerta);
      c.getFormat().getFont().setBold(true);
    }
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(200);
  ws.getRangeByIndexes(0, colPem, 1, colAviso - colPem).getFormat().setColumnWidth(115);
  ws.getRangeByIndexes(0, colAviso, 1, 1).getFormat().setColumnWidth(340);

  // --- Inmovilizar: 1 fila y 1 columna, SIEMPRE con un rango de ESTA hoja ---
  fijarPaneles(ws, 1, 1);
}

// =====================================================================================
// HOJA 2: "Resumen capitulos"
// =====================================================================================
function formatearCapitulos(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_CAPITULOS");
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "CAP_NUM_COLS");
  const headFilas = metaNum(meta, "CAP_HEADER_FILAS");
  const numCap = metaNum(meta, "CAP_NUM_CAPITULOS");
  const filaTotal = metaNum(meta, "CAP_FILA_TOTAL");
  const colPem = metaNum(meta, "CAP_COL_PEM");
  const colMedia = metaNum(meta, "CAP_COL_MEDIA");
  const col1a = metaNum(meta, "CAP_COL_PRIMERA_OFERTA");
  const porOferta = metaNum(meta, "CAP_COLS_POR_OFERTA");
  const nDatos = numCap + 1; // capitulos + fila TOTAL

  // --- Cabecera (2 filas). Primero deshacer merges por si B ya se ejecuto antes ---
  let bloqueCab = ws.getRangeByIndexes(0, 0, headFilas, COLS);
  bloqueCab.unmerge();
  bloqueCab.getFormat().getFont().setBold(true);
  bloqueCab.getFormat().getFont().setColor("FFFFFF");
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorCabecera);
  ws.getRangeByIndexes(1, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorCabecera2);
  bloqueCab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Merge del nombre de cada constructora sobre sus 3 columnas
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(0, col1a + i * porOferta, 1, porOferta).merge(false);
  }

  // --- Formatos de numero por rangos ---
  ws.getRangeByIndexes(headFilas, colPem, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
  ws.getRangeByIndexes(headFilas, colMedia, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos, 1).setNumberFormat(cfg.fmtEuro0);
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 2, nDatos, 1).setNumberFormat(cfg.fmtPct);
  }

  // --- Celdas de entrada manual (PEM) ---
  ws.getRangeByIndexes(headFilas, colPem, numCap, 1).getFormat().getFill().setColor(cfg.colorManual);

  // --- Min/max por capitulo: verde el importe mas bajo, rojo el mas alto ---
  // Se LEEN los valores ya calculados por A; no se recalcula nada.
  let vals = ws.getRangeByIndexes(headFilas, 0, numCap, COLS).getValues();
  for (let r = 0; r < numCap; r++) {
    let mn = 0, mx = 0, hay = false;
    for (let i = 0; i < N; i++) {
      let v = vals[r][col1a + i * porOferta];
      if (typeof v === "number") {
        if (!hay) { mn = v; mx = v; hay = true; }
        else {
          if (v < mn) { mn = v; }
          if (v > mx) { mx = v; }
        }
      }
    }
    if (!hay || mn === mx) { continue; }
    for (let i = 0; i < N; i++) {
      let v = vals[r][col1a + i * porOferta];
      if (typeof v !== "number") { continue; }
      let celda = ws.getRangeByIndexes(headFilas + r, col1a + i * porOferta, 1, 1);
      if (v === mn) { celda.getFormat().getFill().setColor(cfg.colorMejor); }
      else if (v === mx) { celda.getFormat().getFill().setColor(cfg.colorPeor); }
    }
  }

  // --- Fila TOTAL PEC ---
  let filaT = ws.getRangeByIndexes(filaTotal, 0, 1, COLS);
  filaT.getFormat().getFont().setBold(true);
  filaT.getFormat().getFill().setColor(cfg.colorTotal);

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(60);
  ws.getRangeByIndexes(0, 1, 1, 1).getFormat().setColumnWidth(230);
  ws.getRangeByIndexes(0, colPem, 1, 2).getFormat().setColumnWidth(100);
  ws.getRangeByIndexes(0, col1a, 1, COLS - col1a).getFormat().setColumnWidth(95);

  // --- Inmovilizar: 2 filas de cabecera y 2 columnas (codigo + nombre de capitulo) ---
  fijarPaneles(ws, headFilas, 2);
}

// =====================================================================================
// HOJA 3: "Comparativa partidas"
// =====================================================================================
function formatearPartidas(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_PARTIDAS");
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "PART_NUM_COLS");
  const headFilas = metaNum(meta, "PART_HEADER_FILAS");
  const nDatos = metaNum(meta, "PART_NUM_FILAS");
  const colsFijas = metaNum(meta, "PART_COLS_FIJAS");
  const colTipo = metaNum(meta, "PART_COL_TIPO");
  const colNivel = metaNum(meta, "PART_COL_NIVEL");
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

  // --- Bandas alternas por empresa (por rangos completos, no celda a celda) ---
  for (let i = 0; i < N; i++) {
    if (i % 2 === 1) {
      ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, porOferta)
        .getFormat().getFill().setColor(cfg.colorBanda);
    }
  }

  // --- Formatos de numero por rangos ---
  ws.getRangeByIndexes(headFilas, colMedicion, nDatos, 1).setNumberFormat(cfg.fmtNum2);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtNum2);       // P. Unit
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos, 1).setNumberFormat(cfg.fmtEuro2);  // Importe
  }
  ws.getRangeByIndexes(headFilas, colDesv, nDatos, 1).setNumberFormat(cfg.fmtPct);

  // --- Sombreado de filas de capitulo y resaltado de desviaciones ---
  // Se leen los valores UNA sola vez en bloque; el coloreo por fila se limita a los
  // capitulos (pocos) y a las partidas que superan el umbral (pocas).
  let vals = ws.getRangeByIndexes(headFilas, 0, nDatos, COLS).getValues();
  for (let r = 0; r < nDatos; r++) {
    let tipo = String(vals[r][colTipo]);
    if (tipo === "Capitulo") {
      let nivel = String(vals[r][colNivel]);
      let fila = ws.getRangeByIndexes(headFilas + r, 0, 1, COLS);
      fila.getFormat().getFill().setColor(nivel === "padre" ? cfg.colorCapPadre : cfg.colorCapSub);
      fila.getFormat().getFont().setBold(true);
    } else if (tipo === "Partida") {
      let d = vals[r][colDesv];
      if (typeof d === "number" && d >= cfg.umbralDesviacionPartida) {
        let celda = ws.getRangeByIndexes(headFilas + r, colDesv, 1, 1);
        celda.getFormat().getFill().setColor(cfg.colorAlerta);
        celda.getFormat().getFont().setBold(true);
      }
    }
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().setColumnWidth(85);     // por defecto
  ws.getRangeByIndexes(0, colTipo, 1, 1).getFormat().setColumnWidth(70);
  ws.getRangeByIndexes(0, colNivel, 1, 1).getFormat().setColumnWidth(55);
  ws.getRangeByIndexes(0, colResumen, 1, 1).getFormat().setColumnWidth(330);
  ws.getRangeByIndexes(0, colUd, 1, 1).getFormat().setColumnWidth(45);
  ws.getRangeByIndexes(0, colMedicion, 1, 1).getFormat().setColumnWidth(90);
  ws.getRangeByIndexes(0, colBarata, 1, 1).getFormat().setColumnWidth(100);

  // --- Inmovilizar: 2 filas de cabecera y las 6 columnas fijas ---
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
  // Asegurar que el contrato sigue oculto
  let c = wb.getWorksheet(cfg.hojaContrato);
  if (c) { c.setVisibility(ExcelScript.SheetVisibility.hidden); }
  // Dejar activa la hoja de resumen
  let rs = wb.getWorksheet(metaStr(meta, "HOJA_RESUMEN"));
  if (rs) { rs.activate(); }
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
  colorCapPadre: string;
  colorCapSub: string;
  colorBanda: string;
  colorMejor: string;
  colorPeor: string;
  colorAlerta: string;
  colorTotal: string;
  colorManual: string;
  umbralDesviacionPartida: number;
  fmtEuro2: string;
  fmtEuro0: string;
  fmtPct: string;
  fmtNum2: string;
}
