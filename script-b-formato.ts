// =====================================================================================
// SCRIPT B — FORMATO DEL COMPARADOR DE OFERTAS (solo estética, no toca números)
// =====================================================================================
// Se ejecuta DESPUÉS del Script A (cálculo) sobre el mismo libro. Lee la hoja oculta
// "_CONTRATO" que dejó A para saber cuántas ofertas hay, cómo se llaman las hojas y en
// qué fila/columna está cada cosa. NO recalcula nada: solo lee valores ya escritos
// cuando lo necesita para decidir un formato (p. ej. qué filas son de tipo Capitulo, o
// si la celda de cuadre dice "CUADRE OK" o "DESCUADRE").
//
// Es idempotente: se puede ejecutar todas las veces que haga falta sin romper los datos
// (deshace merges, inmovilizaciones y formatos condicionales antes de reaplicarlos).
// La hoja "_CONTRATO" se deja OCULTA al terminar (no se borra, para poder volver a
// ejecutar este script sin repetir el A).
//
// -------------------------------------------------------------------------------------
// CONTRATO CON EL SCRIPT A (VERSION 3)
// -------------------------------------------------------------------------------------
// La hoja "_CONTRATO" (oculta) tiene pares clave/valor: col A = clave, col B = valor.
// Las claves están documentadas en la cabecera del Script A. Índices 0-based. Si falta
// la hoja o una clave, este script se detiene con un error claro pidiendo ejecutar A.
// Este script exige VERSION 3 (hojas Portada y Top desviaciones, celdas de CUADRE,
// columnas Ranking/Plazo en Resumen y Alcance/Desv. s/media en Comparativa).
//
// Celdas de CUADRE (*_FILA_CUADRE/*_COL_CUADRE y POR_FILA_CUADRE): verde si el texto
// empieza por "CUADRE OK", rojo con letra blanca si empieza por "DESCUADRE".
//
// -------------------------------------------------------------------------------------
// RESTRICCIONES DEL RUNTIME (Office Scripts desde Power Automate)
// -------------------------------------------------------------------------------------
// Runtime síncrono: nada de .find(), .reduce(), .map(), .filter(), spread (...) ni
// flechas en callbacks de array. Solo bucles `for` clásicos. Formato por RANGOS; el
// semáforo se hace con formato condicional (color scale), una sola llamada por columna.
// Los enums de la API (BorderWeight, ConditionalFormatType...) NUNCA se guardan en
// variables (regla de aliasing del linter): van escritos en la propia llamada.
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
    colorAlerta: "FFC000",           // naranja: avisos, descuadres y alcance singular
    colorTotal: "DDEBF7",            // filas TOTAL
    colorManual: "FFF2CC",           // amarillo suave: celdas de entrada manual (PEM, Plazo)

    // Colores de la celda de CUADRE (estado del cuadre global)
    colorCuadreOk: "C6EFCE",         // fondo verde si "CUADRE OK"
    colorCuadreOkTexto: "006100",    // letra verde oscuro
    colorCuadreMal: "C00000",        // fondo rojo si "DESCUADRE"
    colorCuadreMalTexto: "FFFFFF",   // letra blanca sobre rojo

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
  if (metaNum(meta, "VERSION") !== 3) {
    throw new Error("El contrato es de otra version (" + meta["VERSION"] + "); este Script B necesita la VERSION 3. Ejecuta el Script A actualizado antes que este.");
  }

  formatearPortada(workbook, meta, cfg);
  formatearResumen(workbook, meta, cfg);
  formatearCapitulos(workbook, meta, cfg);
  formatearAgrupado(workbook, meta, cfg);
  formatearPartidas(workbook, meta, cfg);
  formatearTop(workbook, meta, cfg);

  // Limpieza final: quitar hojas residuales, ocultar el contrato y activar la Portada
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
// HOJA "Portada": resumen ejecutivo
// =====================================================================================
// El formato de cada celda de valor se decide por el TEXTO de su etiqueta (nada de
// indices de fila fijos): si la etiqueta termina en "(%)" es porcentaje; si contiene
// "(€)", "(PEC)" o "máx − mín" es un importe en euros; el resto (conteos, textos,
// fecha, obra) se queda en formato general.
function formatearPortada(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_PORTADA");
  const numFilas = metaNum(meta, "POR_NUM_FILAS");
  const colEti = metaNum(meta, "POR_COL_ETIQUETA");
  const colVal = metaNum(meta, "POR_COL_VALOR");
  const filaCuadre = metaNum(meta, "POR_FILA_CUADRE");

  // --- Titulo grande en la fila 0 ---
  let titulo = ws.getRangeByIndexes(0, 0, 1, 2);
  titulo.getFormat().getFont().setBold(true);
  titulo.getFormat().getFont().setSize(16);
  titulo.getFormat().getFont().setColor(cfg.colorCabecera);

  // --- Columna de etiquetas en negrita (desde la fila 1: la 0 es el titulo) ---
  if (numFilas > 1) {
    ws.getRangeByIndexes(1, colEti, numFilas - 1, 1).getFormat().getFont().setBold(true);
  }

  // --- Formato de los valores segun su etiqueta ---
  let vals = ws.getRangeByIndexes(0, 0, numFilas, colVal + 1).getValues();
  for (let r = 1; r < numFilas; r++) {
    let eti = String(vals[r][colEti]);
    if (eti === "") { continue; }
    let celda = ws.getRangeByIndexes(r, colVal, 1, 1);
    if (eti.endsWith("(%)")) {
      celda.setNumberFormat(cfg.fmtPct);
    } else if (eti.indexOf("(€)") >= 0 || eti.indexOf("(PEC)") >= 0 || eti.indexOf("máx − mín") >= 0) {
      celda.setNumberFormat(cfg.fmtEuro2);
    }
    // Conteos (Nº de ofertas), textos y fecha: sin formato de moneda
  }

  // --- Celda de estado de cuadre: en la Portada el texto esta en POR_COL_VALOR
  //     (no existe POR_COL_CUADRE) y su etiqueta en POR_COL_ETIQUETA ---
  pintarCuadre(ws, filaCuadre, colEti, colVal, cfg);

  // --- Anchos amplios ---
  ws.getRangeByIndexes(0, colEti, 1, 1).getFormat().setColumnWidth(300);
  ws.getRangeByIndexes(0, colVal, 1, 1).getFormat().setColumnWidth(280);
}

// =====================================================================================
// HOJA "Resumen ofertas"
// =====================================================================================
// La hoja llega ORDENADA por PEC ascendente desde el Script A, con la columna Ranking
// en RES_COL_RANKING y la de Plazo (manual) en RES_COL_PLAZO. Todos los indices salen
// del contrato: aqui no hay numeros de columna fijos.
function formatearResumen(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_RESUMEN");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "RES_NUM_COLS");
  const filaResIni = metaNum(meta, "RES_FILA_RESUMEN_INI");
  const colRanking = metaNum(meta, "RES_COL_RANKING");
  const colEmpresa = metaNum(meta, "RES_COL_EMPRESA");
  const colPem = metaNum(meta, "RES_COL_PEM");
  const colPec = metaNum(meta, "RES_COL_PEC");
  const colDifEcoEur = metaNum(meta, "RES_COL_DIF_ECO_EUR");
  const colDifEcoPct = metaNum(meta, "RES_COL_DIF_ECO_PCT");
  const colDifMedEur = metaNum(meta, "RES_COL_DIF_MED_EUR");
  const colDifMedPct = metaNum(meta, "RES_COL_DIF_MED_PCT");
  const colAviso = metaNum(meta, "RES_COL_AVISO");
  const colPlazo = metaNum(meta, "RES_COL_PLAZO");
  const filaCuadre = metaNum(meta, "RES_FILA_CUADRE");
  const colCuadre = metaNum(meta, "RES_COL_CUADRE");

  // --- Cabecera: azul corporativo, texto blanco, negrita ---
  let cab = ws.getRangeByIndexes(0, 0, 1, COLS);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("FFFFFF");
  cab.getFormat().getFill().setColor(cfg.colorCabecera);
  cab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // --- Formatos de numero (por rangos; los calculados por el contrato, sin aritmetica
  //     a ojo: debajo de las filas de resumen hay ahora blanco + fila de CUADRE) ---
  ws.getRangeByIndexes(1, colPem, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colPec, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(filaResIni, colPec, 3, 1).setNumberFormat(cfg.fmtEuro2); // economica, media, horquilla
  ws.getRangeByIndexes(1, colDifEcoEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifEcoPct, N, 1).setNumberFormat(cfg.fmtPct);
  ws.getRangeByIndexes(1, colDifMedEur, N, 1).setNumberFormat(cfg.fmtEuro2);
  ws.getRangeByIndexes(1, colDifMedPct, N, 1).setNumberFormat(cfg.fmtPct);
  // OJO: Plazo NO lleva formato de euros (es una duracion, no un importe): general.

  // --- Ranking: centrado y en negrita ---
  let rk = ws.getRangeByIndexes(1, colRanking, N, 1);
  rk.getFormat().getFont().setBold(true);
  rk.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // --- Verde: la oferta mas economica. El Script A garantiza el orden ascendente por
  //     PEC, asi que la primera fila de datos es SIEMPRE la ganadora ---
  ws.getRangeByIndexes(1, 0, 1, COLS).getFormat().getFill().setColor(cfg.colorMejor);

  // --- Celdas de entrada manual (PEM y Plazo): fondo amarillo suave, mismo tratamiento.
  //     Se pintan DESPUES del verde para que se vea que son editables tambien en la
  //     fila ganadora ---
  ws.getRangeByIndexes(1, colPem, N, 1).getFormat().getFill().setColor(cfg.colorManual);
  ws.getRangeByIndexes(1, colPlazo, N, 1).getFormat().getFill().setColor(cfg.colorManual);

  // --- Filas de resumen final en negrita ---
  ws.getRangeByIndexes(filaResIni, 0, 3, COLS).getFormat().getFont().setBold(true);

  // --- Naranja: ofertas con aviso (PEM sin convertir y/o oferta incompleta; pueden
  //     venir los dos concatenados en la misma celda). Solo LECTURA para decidir ---
  let vals = ws.getRangeByIndexes(1, 0, N, COLS).getValues();
  for (let i = 0; i < N; i++) {
    if (String(vals[i][colAviso]) !== "") {
      let c = ws.getRangeByIndexes(1 + i, colAviso, 1, 1);
      c.getFormat().getFill().setColor(cfg.colorAlerta);
      c.getFormat().getFont().setBold(true);
    }
  }

  // --- Celda de CUADRE visible ---
  pintarCuadre(ws, filaCuadre, 0, colCuadre, cfg);

  // --- Anchos ---
  ws.getRangeByIndexes(0, colRanking, 1, 1).getFormat().setColumnWidth(70);
  ws.getRangeByIndexes(0, colEmpresa, 1, 1).getFormat().setColumnWidth(200);
  ws.getRangeByIndexes(0, colPem, 1, colAviso - colPem).getFormat().setColumnWidth(115);
  ws.getRangeByIndexes(0, colAviso, 1, 1).getFormat().setColumnWidth(340);
  ws.getRangeByIndexes(0, colPlazo, 1, 1).getFormat().setColumnWidth(115);

  // --- Inmovilizar: 1 fila y las columnas Ranking + Empresa ---
  fijarPaneles(ws, 1, colEmpresa + 1);
}

// =====================================================================================
// HOJA "Resumen capitulos"
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
  const filaCuadre = metaNum(meta, "CAP_FILA_CUADRE");
  const colCuadre = metaNum(meta, "CAP_COL_CUADRE");
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

  // --- Celda de CUADRE visible (fuera del marco de cada oferta, a proposito) ---
  pintarCuadre(ws, filaCuadre, 0, colCuadre, cfg);

  // --- Bordes gruesos encuadrando el bloque de cada oferta (hasta la fila TOTAL:
  //     la fila de CUADRE queda fuera del marco) ---
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
// HOJA "Resumen agrupado"
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
  const filaCuadre = metaNum(meta, "AGR_FILA_CUADRE");
  const colCuadre = metaNum(meta, "AGR_COL_CUADRE");
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

  // --- Celda de CUADRE visible (fuera del marco de cada oferta, a proposito) ---
  pintarCuadre(ws, filaCuadre, 0, colCuadre, cfg);

  // --- Bordes gruesos encuadrando el bloque de cada oferta (hasta la fila TOTAL) ---
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
// HOJA "Comparativa partidas"
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
  const colAlcance = metaNum(meta, "PART_COL_ALCANCE");
  const colDesvMedia = metaNum(meta, "PART_COL_DESV_MEDIA");
  const filaTotal = metaNum(meta, "PART_FILA_TOTAL");
  const filaCuadre = metaNum(meta, "PART_FILA_CUADRE");
  const colCuadre = metaNum(meta, "PART_COL_CUADRE");

  // --- Cabecera (2 filas). Deshacer merges antes de rehacerlos (idempotencia) ---
  let bloqueCab = ws.getRangeByIndexes(0, 0, headFilas, COLS);
  bloqueCab.unmerge();
  ws.getRangeByIndexes(0, 0, 1, colsFijas).merge(false);              // titulo
  for (let i = 0; i < N; i++) {                                        // una celda por empresa
    ws.getRangeByIndexes(0, col1a + i * porOferta, 1, porOferta).merge(false);
  }
  // Bloque COMPARACION: ahora abarca Mas barata, Δ %, Alcance y Desv. s/media
  ws.getRangeByIndexes(0, colBarata, 1, COLS - colBarata).merge(false);

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

  // --- Formatos de numero por rangos. Los Importes se extienden UNA fila mas para
  //     cubrir la fila "TOTAL partidas" (filaTotal = headFilas + nDatos) ---
  ws.getRangeByIndexes(headFilas, colMedicion, nDatos, 1).setNumberFormat(cfg.fmtNum2);
  for (let i = 0; i < N; i++) {
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta, nDatos, 1).setNumberFormat(cfg.fmtNum2);           // P. Unit
    ws.getRangeByIndexes(headFilas, col1a + i * porOferta + 1, nDatos + 1, 1).setNumberFormat(cfg.fmtEuro2);  // Importe + TOTAL
  }
  ws.getRangeByIndexes(headFilas, colDesv, nDatos, 1).setNumberFormat(cfg.fmtPct);

  // --- SEMÁFORO en la columna de comparacion Δ %: poca desviacion en verde,
  //     mucha desviacion (partidas con gran horquilla de precios) en rojo ---
  escalaSemaforo(ws, headFilas, colDesv, nDatos, cfg);

  // --- Sombreado por filas segun la columna Tipo + resaltado de Alcance (una sola
  //     lectura en bloque; solo lectura para decidir formatos) ---
  // Tipo="Capitulo" → azul fuerte; Tipo="Subcapitulo" → azul claro; ambas en negrita.
  // Alcance no vacio ("SOLO <oferta>") → celda naranja + negrita: diferencia de alcance.
  let vals = ws.getRangeByIndexes(headFilas, 0, nDatos, COLS).getValues();
  for (let r = 0; r < nDatos; r++) {
    let tipo = String(vals[r][colTipo]);
    if (tipo === "Capitulo" || tipo === "Subcapitulo") {
      let fila = ws.getRangeByIndexes(headFilas + r, 0, 1, COLS);
      fila.getFormat().getFill().setColor(tipo === "Capitulo" ? cfg.colorCapitulo : cfg.colorSubcapitulo);
      fila.getFormat().getFont().setBold(true);
    }
    if (String(vals[r][colAlcance]) !== "") {
      let c = ws.getRangeByIndexes(headFilas + r, colAlcance, 1, 1);
      c.getFormat().getFill().setColor(cfg.colorAlerta);
      c.getFormat().getFont().setBold(true);
    }
  }

  // --- Fila TOTAL partidas: negrita y mismo fondo que las filas TOTAL de otras hojas ---
  let filaT = ws.getRangeByIndexes(filaTotal, 0, 1, COLS);
  filaT.getFormat().getFont().setBold(true);
  filaT.getFormat().getFill().setColor(cfg.colorTotal);

  // --- Celda de CUADRE visible ---
  pintarCuadre(ws, filaCuadre, 0, colCuadre, cfg);

  // --- Bordes gruesos encuadrando el bloque (P. Unit + Importe) de cada oferta,
  //     incluida la fila TOTAL partidas (el CUADRE queda fuera del marco) ---
  for (let i = 0; i < N; i++) {
    encuadrar(ws, 0, col1a + i * porOferta, filaTotal + 1, porOferta);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().setColumnWidth(85);     // por defecto
  ws.getRangeByIndexes(0, colTipo, 1, 1).getFormat().setColumnWidth(85);
  ws.getRangeByIndexes(0, colResumen, 1, 1).getFormat().setColumnWidth(330);
  ws.getRangeByIndexes(0, colUd, 1, 1).getFormat().setColumnWidth(45);
  ws.getRangeByIndexes(0, colMedicion, 1, 1).getFormat().setColumnWidth(90);
  ws.getRangeByIndexes(0, colBarata, 1, 1).getFormat().setColumnWidth(100);
  ws.getRangeByIndexes(0, colAlcance, 1, 1).getFormat().setColumnWidth(140);
  ws.getRangeByIndexes(0, colDesvMedia, 1, 1).getFormat().setColumnWidth(260);

  // --- Inmovilizar: 2 filas de cabecera y las columnas fijas hasta Medicion ---
  fijarPaneles(ws, headFilas, colsFijas);
}

// =====================================================================================
// HOJA "Top desviaciones"
// =====================================================================================
// Las partidas con mayor diferencia en euros entre ofertas. Semaforo sobre "Dif €"
// para que el dinero en juego salte a la vista.
function formatearTop(wb: ExcelScript.Workbook, meta: Contrato, cfg: CfgB) {
  let ws = hojaDelContrato(wb, meta, "HOJA_TOP");
  limpiarCondicionales(ws);
  const N = metaNum(meta, "NUM_OFERTAS");
  const COLS = metaNum(meta, "TOP_NUM_COLS");
  const headFilas = metaNum(meta, "TOP_HEADER_FILAS");
  const nDatos = metaNum(meta, "TOP_NUM_FILAS");
  const colCodigo = metaNum(meta, "TOP_COL_CODIGO");
  const colResumen = metaNum(meta, "TOP_COL_RESUMEN");
  const col1a = metaNum(meta, "TOP_COL_PRIMERA_OFERTA");
  const porOferta = metaNum(meta, "TOP_COLS_POR_OFERTA");
  const colDifEur = metaNum(meta, "TOP_COL_DIF_EUR");
  const colDifPct = metaNum(meta, "TOP_COL_DIF_PCT");
  const colBarata = metaNum(meta, "TOP_COL_MAS_BARATA");

  // --- Cabecera corporativa (1 fila) ---
  let cab = ws.getRangeByIndexes(0, 0, headFilas, COLS);
  cab.getFormat().getFont().setBold(true);
  cab.getFormat().getFont().setColor("FFFFFF");
  cab.getFormat().getFill().setColor(cfg.colorCabecera);
  cab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  if (nDatos > 0) {
    // --- Formatos de numero: importes por oferta (1 col/oferta, contiguas) y Dif €
    //     en euros; Dif % en porcentaje ---
    ws.getRangeByIndexes(headFilas, col1a, nDatos, N * porOferta).setNumberFormat(cfg.fmtEuro2);
    ws.getRangeByIndexes(headFilas, colDifEur, nDatos, 1).setNumberFormat(cfg.fmtEuro2);
    ws.getRangeByIndexes(headFilas, colDifPct, nDatos, 1).setNumberFormat(cfg.fmtPct);

    // --- SEMÁFORO sobre Dif €: las mayores diferencias en rojo ---
    escalaSemaforo(ws, headFilas, colDifEur, nDatos, cfg);
  }

  // --- Anchos ---
  ws.getRangeByIndexes(0, 0, 1, COLS).getFormat().setColumnWidth(110);   // por defecto
  ws.getRangeByIndexes(0, colCodigo, 1, 1).getFormat().setColumnWidth(85);
  ws.getRangeByIndexes(0, colResumen, 1, 1).getFormat().setColumnWidth(330);
  ws.getRangeByIndexes(0, colBarata, 1, 1).getFormat().setColumnWidth(110);

  // --- Inmovilizar la fila de cabecera ---
  fijarPaneles(ws, headFilas, 0);
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
  // Dejar activa la Portada (el resumen ejecutivo es lo primero que se quiere ver)
  let por = wb.getWorksheet(metaStr(meta, "HOJA_PORTADA"));
  if (por) { por.activate(); }
}

// =====================================================================================
// AYUDANTES DE FORMATO
// =====================================================================================

// Celda de CUADRE: verde + negrita si el texto empieza por "CUADRE OK"; rojo con letra
// blanca + negrita si empieza por "DESCUADRE". Se pinta el tramo desde la etiqueta
// (colEtiqueta) hasta la celda del texto (colTexto), ambas incluidas. La lectura del
// valor es SOLO para decidir el formato.
function pintarCuadre(ws: ExcelScript.Worksheet, fila: number, colEtiqueta: number,
  colTexto: number, cfg: CfgB) {
  let v = ws.getRangeByIndexes(fila, colTexto, 1, 1).getValues();
  let texto = String(v[0][0]);
  let esOk = texto.indexOf("CUADRE OK") === 0;
  let esMal = texto.indexOf("DESCUADRE") === 0;
  if (!esOk && !esMal) { return; } // contenido inesperado: no pintar nada

  let ancho = colTexto - colEtiqueta + 1;
  let r = ws.getRangeByIndexes(fila, colEtiqueta, 1, ancho);
  r.getFormat().getFont().setBold(true);
  if (esOk) {
    r.getFormat().getFill().setColor(cfg.colorCuadreOk);
    r.getFormat().getFont().setColor(cfg.colorCuadreOkTexto);
  } else {
    r.getFormat().getFill().setColor(cfg.colorCuadreMal);
    r.getFormat().getFont().setColor(cfg.colorCuadreMalTexto);
  }
}

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
  colorCuadreOk: string;
  colorCuadreOkTexto: string;
  colorCuadreMal: string;
  colorCuadreMalTexto: string;
  escalaVerde: string;
  escalaAmarillo: string;
  escalaRojo: string;
  fmtEuro2: string;
  fmtEuro0: string;
  fmtPct: string;
  fmtNum2: string;
}
