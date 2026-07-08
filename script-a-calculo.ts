// =====================================================================================
// SCRIPT A — CÁLCULO DEL COMPARADOR DE OFERTAS (solo números, sin formato)
// =====================================================================================
// Recibe `ofertasJson` desde Power Automate y genera TRES hojas con SOLO valores:
//   1. "Resumen ofertas"       → totales PEC, desviaciones y avisos de anomalía PEM/PEC
//   2. "Resumen capitulos"     → importes por capítulo padre y desviaciones vs media
//   3. "Comparativa partidas"  → una fila por partida/capítulo con precios e importes
//
// Este script NO aplica colores, ni merges, ni anchos, ni paneles inmovilizados.
// De la estética se encarga el SCRIPT B (formato), que se ejecuta después.
//
// -------------------------------------------------------------------------------------
// CONTRATO ENTRE SCRIPT A Y SCRIPT B
// -------------------------------------------------------------------------------------
// Para que el Script B no tenga que "adivinar" dónde está cada cosa, este script deja
// una hoja OCULTA llamada "_CONTRATO" con pares clave/valor (columna A = clave,
// columna B = valor). El Script B lee esa hoja y se orienta con ella. Si en el futuro
// se añade una columna o cambia el número de ofertas, B seguirá funcionando sin tocarlo,
// porque todo lo que necesita (número de ofertas, nombres, índices de columnas y filas
// clave de cada hoja) viaja en el contrato.
//
// Claves que escribe este script (todas 0-based, es decir, la primera fila/columna es 0):
//   VERSION, NUM_OFERTAS, NOMBRE_1..NOMBRE_n
//   HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_PARTIDAS          → nombres reales de las hojas
//   RES_HEADER_FILAS, RES_FILA_RESUMEN_INI, RES_NUM_COLS,
//   RES_COL_EMPRESA, RES_COL_PEM, RES_COL_PEC, RES_COL_DIF_ECO_EUR, RES_COL_DIF_ECO_PCT,
//   RES_COL_DIF_MED_EUR, RES_COL_DIF_MED_PCT, RES_COL_AVISO
//   CAP_HEADER_FILAS, CAP_NUM_CAPITULOS, CAP_FILA_TOTAL, CAP_NUM_COLS,
//   CAP_COL_COD, CAP_COL_NOMBRE, CAP_COL_PEM, CAP_COL_MEDIA,
//   CAP_COL_PRIMERA_OFERTA, CAP_COLS_POR_OFERTA, CAP_COLS_FIJAS
//   PART_HEADER_FILAS, PART_NUM_FILAS, PART_NUM_COLS, PART_COLS_FIJAS,
//   PART_COL_CODIGO, PART_COL_TIPO, PART_COL_NIVEL, PART_COL_RESUMEN, PART_COL_UD,
//   PART_COL_MEDICION, PART_COL_PRIMERA_OFERTA, PART_COLS_POR_OFERTA,
//   PART_COL_MAS_BARATA, PART_COL_DESV
//
// Si el Script B no encuentra la hoja "_CONTRATO" o le falta una clave, se detiene con
// un error claro pidiendo ejecutar primero este script.
//
// -------------------------------------------------------------------------------------
// RESTRICCIONES DEL RUNTIME (Office Scripts desde Power Automate)
// -------------------------------------------------------------------------------------
// El runtime es síncrono y NO admite .find(), .reduce(), .map(), spread (...), ni
// funciones flecha en callbacks de array. TODO se hace con bucles `for` clásicos.
// La escritura en hojas se hace SIEMPRE en bloque con setValues() por rendimiento.
// =====================================================================================

function main(workbook: ExcelScript.Workbook, ofertasJson?: string) {

  // ===================================================================================
  // ZONA DE CONFIGURACIÓN (editable)
  // ===================================================================================
  // Nombres de las hojas de salida. Si los cambias, el Script B se adapta solo,
  // porque los lee del contrato.
  const HOJA_RESUMEN = "Resumen ofertas";
  const HOJA_CAPITULOS = "Resumen capitulos";
  const HOJA_PARTIDAS = "Comparativa partidas";
  const HOJA_CONTRATO = "_CONTRATO"; // hoja oculta de metadatos (NO renombrar sin cambiar B)

  // Umbral de anomalía PEM/PEC: si el total de una oferta está más de este porcentaje
  // POR DEBAJO de la media, se escribe un aviso en la columna "Aviso" de la hoja resumen.
  // Suele indicar una oferta entregada en PEM (sin GG, BI ni IVA) en vez de en PEC.
  // 0.20 = 20 %. Es una red de seguridad, NO una garantía: verificar siempre a mano.
  const UMBRAL_ANOMALIA_PEM = 0.20;

  // ===================================================================================
  // VALIDACIÓN DE ENTRADA
  // ===================================================================================
  if (!ofertasJson || ofertasJson.trim().length === 0) {
    throw new Error("El parametro ofertasJson llego vacio. Revisa el paso anterior del flujo de Power Automate (script lector).");
  }
  let ofertas: Oferta[];
  try {
    ofertas = JSON.parse(ofertasJson) as Oferta[];
  } catch (e) {
    throw new Error("ofertasJson no es un JSON valido. Detalle: " + String(e));
  }
  if (!ofertas || ofertas.length === 0) {
    throw new Error("El JSON no contiene ninguna oferta.");
  }

  // Nombres de oferta limpios (sin extension .xlsx/.xlsm/.xls, por si llegan sucios)
  let nombres: string[] = [];
  for (let i = 0; i < ofertas.length; i++) {
    nombres.push(limpiarNombre(ofertas[i].nombre));
  }
  const N = nombres.length;

  // ===================================================================================
  // AGREGACIÓN DE DATOS
  // ===================================================================================
  // Estructuras de acumulación:
  //  - orden / info: union de todos los codigos (capitulos y partidas) en el orden en
  //    que aparecen por primera vez (respeta el orden original del presupuesto).
  //  - precioMap / importeMap: por oferta, precio unitario e importe de cada codigo.
  //  - capOrden / capNombre / capImporte: capitulos PADRE deducidos y sus importes
  //    acumulados por oferta.
  //  - totalPEC: suma de importes de partidas de cada oferta.
  let orden: string[] = [];
  let seen: { [c: string]: boolean } = {};
  let info: { [c: string]: CodeInfo } = {};
  let precioMap: { [c: string]: number }[] = [];
  let importeMap: { [c: string]: number }[] = [];

  let capOrden: string[] = [];
  let capSeen: { [c: string]: boolean } = {};
  let capNombre: { [c: string]: string } = {};
  let capImporte: { [c: string]: number }[] = [];
  let totalPEC: number[] = [];

  for (let k = 0; k < N; k++) {
    let o = ofertas[k];
    let fs: Fila[] = o.filas ? o.filas : [];
    if (fs.length === 0) {
      throw new Error("La oferta \"" + nombres[k] + "\" no contiene filas. Revisa su Excel de origen.");
    }
    let pM: { [c: string]: number } = {};
    let iM: { [c: string]: number } = {};
    let cM: { [c: string]: number } = {};
    let tot = 0;
    // Capitulo padre "en curso": se actualiza cada vez que aparece una fila de capitulo.
    // Las partidas que vienen despues se acumulan bajo el.
    let padreActual = "";

    for (let j = 0; j < fs.length; j++) {
      let f = fs[j];

      if (f.tipo === "capitulo") {
        // Registrar el capitulo/subcapitulo como fila propia de la comparativa
        if (!seen[f.codigo]) {
          seen[f.codigo] = true;
          orden.push(f.codigo);
          info[f.codigo] = { tipo: "Capitulo", nivel: f.nivel ? f.nivel : "", resumen: f.resumen, unidad: "", medicion: 0 };
        }
        if (iM[f.codigo] === undefined) { iM[f.codigo] = f.importe; }
        // Deducir el capitulo PADRE logico a partir del codigo
        padreActual = capituloPadre(f.codigo);
        if (!capSeen[padreActual]) {
          capSeen[padreActual] = true;
          capOrden.push(padreActual);
          capNombre[padreActual] = f.resumen;
        } else if (f.codigo === padreActual || capNombre[padreActual] === "") {
          // Si aparece la fila del propio padre, su resumen manda sobre el del primer sub
          capNombre[padreActual] = f.resumen;
        }
        continue;
      }

      // ------- Fila de partida -------
      if (!seen[f.codigo]) {
        seen[f.codigo] = true;
        orden.push(f.codigo);
        // Nota: la medicion se toma de la PRIMERA oferta donde aparece la partida.
        // Si una oferta trae una medicion distinta, aqui no se detecta (suele coincidir).
        info[f.codigo] = { tipo: "Partida", nivel: "", resumen: f.resumen, unidad: f.unidad, medicion: f.medicion };
      }
      if (pM[f.codigo] === undefined) { pM[f.codigo] = f.precio; iM[f.codigo] = f.importe; }

      // Capitulo padre de la partida: normalmente el del ultimo capitulo visto (las filas
      // vienen en orden). Si aun no ha aparecido ningun capitulo, se deduce del codigo
      // de la propia partida como red de seguridad.
      let padre = padreActual !== "" ? padreActual : capituloPadre(f.codigo);
      if (padre !== "") {
        if (!capSeen[padre]) { capSeen[padre] = true; capOrden.push(padre); capNombre[padre] = ""; }
        cM[padre] = (cM[padre] || 0) + f.importe;
      }
      tot += f.importe;
    }

    precioMap.push(pM);
    importeMap.push(iM);
    capImporte.push(cM);
    totalPEC.push(tot);
  }

  // ===================================================================================
  // ESCRITURA DE LAS TRES HOJAS (solo valores)
  // ===================================================================================
  let layR = hojaResumenA(workbook, HOJA_RESUMEN, nombres, totalPEC, UMBRAL_ANOMALIA_PEM);
  let layC = hojaCapitulosA(workbook, HOJA_CAPITULOS, nombres, capOrden, capNombre, capImporte);
  let layP = hojaPartidasA(workbook, HOJA_PARTIDAS, nombres, orden, info, precioMap, importeMap);

  // ===================================================================================
  // ESCRITURA DEL CONTRATO (hoja oculta que orienta al Script B)
  // ===================================================================================
  escribirContrato(workbook, HOJA_CONTRATO, HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_PARTIDAS,
    nombres, layR, layC, layP);
}

// =====================================================================================
// TABLA DE EXCEPCIONES DE CAPÍTULO — ¡¡CRITERIO DEL USUARIO, VERIFICAR EN CADA PROYECTO!!
// =====================================================================================
// Algunos codigos no revelan su capitulo padre real. Esta tabla mapea el primer token
// numerico del codigo a un capitulo concreto. Ejemplos actuales:
//   - Telecomunicaciones: codigos que empiezan por 1., 2., 3. o 4  →  C33
//   - Aparatos sanitarios: codigos que empiezan por 23             →  C16
// Para anadir una excepcion, anade una linea t.push({ token: "XX", capitulo: "Cnn" });
function tablaExcepciones(): ExcepcionCapitulo[] {
  let t: ExcepcionCapitulo[] = [];
  t.push({ token: "1", capitulo: "C33" });  // telecomunicaciones
  t.push({ token: "2", capitulo: "C33" });  // telecomunicaciones
  t.push({ token: "3", capitulo: "C33" });  // telecomunicaciones
  t.push({ token: "4", capitulo: "C33" });  // telecomunicaciones
  t.push({ token: "23", capitulo: "C16" }); // aparatos sanitarios
  return t;
}

function excepcion(token: string): string {
  let t = tablaExcepciones();
  for (let i = 0; i < t.length; i++) {
    if (t[i].token === token) { return t[i].capitulo; }
  }
  return "";
}

// Deduce el capitulo padre "Cnn" a partir de un codigo.
// Patrones observados en los presupuestos reales:
//   F2_Cnn...  |  F2_nn....  |  Cnn...  |  nn.  |  nn   →   capitulo "Cnn"
// Si ningun patron encaja, devuelve el codigo tal cual (se agrupa bajo si mismo).
function capituloPadre(code: string): string {
  let c = code.trim();
  let m = c.match(/^F2[_\- ]?C0*(\d{1,2})/i);
  if (m) { return "C" + pad2(m[1]); }
  m = c.match(/^F2[_\- ]?(\d{1,2})[.]/i);
  if (m) { return "C" + pad2(m[1]); }
  m = c.match(/^C(\d{2})/i);
  if (m) { return "C" + m[1]; }
  m = c.match(/^(\d{1,2})[.]/);
  if (m) { let e = excepcion(m[1]); return e !== "" ? e : "C" + pad2(m[1]); }
  m = c.match(/^(\d{1,2})$/);
  if (m) { let e = excepcion(m[1]); return e !== "" ? e : "C" + pad2(m[1]); }
  return c;
}

// =====================================================================================
// HOJA 1: "Resumen ofertas" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: cabecera
//   Filas 1..N: una por oferta
//   Fila N+1: en blanco
//   Filas N+2..N+4: oferta mas economica, media, horquilla
function hojaResumenA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  totalPEC: number[], umbralAnomalia: number): LayoutResumen {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS = 8;

  // Minimo, maximo y media de los totales
  let mn = totalPEC[0], mx = totalPEC[0], suma = 0;
  for (let i = 0; i < N; i++) {
    let v = totalPEC[i];
    if (v < mn) { mn = v; }
    if (v > mx) { mx = v; }
    suma += v;
  }
  let media = suma / N;

  let filas: (string | number)[][] = [];
  filas.push(["EMPRESA", "PEM (manual)", "IMPORTE PEC", "Δ vs +econ. (€)", "Δ vs +econ. (%)", "Δ vs media (€)", "Δ vs media (%)", "Aviso"]);

  for (let i = 0; i < N; i++) {
    let v = totalPEC[i];
    // Deteccion de anomalia PEM/PEC: oferta muy por debajo de la media.
    // Solo avisa, NO corrige nada. Verificar manualmente el Excel de origen.
    let aviso = "";
    if (N >= 2 && media > 0) {
      let pctBajo = (media - v) / media;
      if (pctBajo > umbralAnomalia) {
        let pctTxt = Math.round(pctBajo * 1000) / 10;
        aviso = "AVISO: importe un " + pctTxt + "% por debajo de la media. Posible oferta en PEM sin convertir a PEC. Verificar manualmente.";
      }
    }
    filas.push([
      nombres[i],
      "", // PEM se rellena a mano
      r2(v),
      r2(v - mn),
      mn > 0 ? r4((v - mn) / mn) : 0,
      r2(v - media),
      media > 0 ? r4((v - media) / media) : 0,
      aviso
    ]);
  }

  filas.push(["", "", "", "", "", "", "", ""]);
  let filaResumenIni = filas.length; // indice 0-based de la fila "Oferta mas economica"
  filas.push(["Oferta más económica", "", r2(mn), "", "", "", "", ""]);
  filas.push(["Media " + N + " ofertas", "", r2(media), "", "", "", "", ""]);
  filas.push(["Horquilla (máx − mín)", "", r2(mx - mn), "", "", "", "", ""]);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { filaResumenIni: filaResumenIni, numCols: COLS };
}

// =====================================================================================
// HOJA 2: "Resumen capitulos" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: nombres de empresa (en la primera columna de su bloque; B hace el merge)
//   Fila 1: cabecera de columnas
//   Filas 2..2+numCap-1: una por capitulo padre
//   Ultima fila: TOTAL PEC
// Columnas: 0=Cod, 1=Capitulo, 2=PEM (manual), 3=Media PEC,
//           luego 3 columnas por oferta (Importe PEC, Delta EUR s/Media, Delta % s/Media)
function hojaCapitulosA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  capOrden: string[], capNombre: { [c: string]: string },
  capImporte: { [c: string]: number }[]): LayoutCapitulos {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS_FIJAS = 4;
  const COLS = COLS_FIJAS + N * 3;

  let cab1: (string | number)[] = ["", "", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); cab1.push(""); }
  let cab2: (string | number)[] = ["Cód.", "Capítulo", "PEM (manual)", "Media PEC"];
  for (let i = 0; i < N; i++) { cab2.push("Importe PEC"); cab2.push("Δ € s/Media"); cab2.push("Δ % s/Media"); }

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  let totCol: number[] = [];
  for (let i = 0; i < N; i++) { totCol.push(0); }

  for (let c = 0; c < capOrden.length; c++) {
    let cap = capOrden[c];
    // Media del capitulo entre las ofertas que lo tienen
    let sumaCap = 0, cuenta = 0;
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) { sumaCap += r2(capImporte[i][cap]); cuenta++; }
    }
    let media = cuenta > 0 ? sumaCap / cuenta : 0;

    let fila: (string | number)[] = [cap, capNombre[cap] || "", "", r2(media)];
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) {
        let v = r2(capImporte[i][cap]);
        totCol[i] += v;
        fila.push(v);
        fila.push(r2(v - media));
        fila.push(media > 0 ? r4((v - media) / media) : 0);
      } else {
        fila.push(""); fila.push(""); fila.push("");
      }
    }
    filas.push(fila);
  }

  // Fila TOTAL PEC
  let mediaTot = 0;
  for (let i = 0; i < N; i++) { mediaTot += totCol[i]; }
  mediaTot = mediaTot / N;
  let filaTotal = filas.length; // indice 0-based de la fila TOTAL
  let fT: (string | number)[] = ["", "TOTAL PEC", "", r2(mediaTot)];
  for (let i = 0; i < N; i++) {
    fT.push(r2(totCol[i]));
    fT.push(r2(totCol[i] - mediaTot));
    fT.push(mediaTot > 0 ? r4((totCol[i] - mediaTot) / mediaTot) : 0);
  }
  filas.push(fT);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numCapitulos: capOrden.length, filaTotal: filaTotal, numCols: COLS };
}

// =====================================================================================
// HOJA 3: "Comparativa partidas" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: titulo + nombres de empresa (primera columna de cada bloque; B hace el merge)
//   Fila 1: cabecera de columnas
//   Filas 2..: una por capitulo/partida en el orden original
// Columnas fijas: 0=Codigo, 1=Tipo (Capitulo/Partida), 2=Nivel (padre/sub/vacio),
//                 3=Resumen, 4=Ud, 5=Medicion
// Luego 2 columnas por oferta (P. Unit, Importe) y al final: Mas barata, Delta %.
function hojaPartidasA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  orden: string[], info: { [c: string]: CodeInfo },
  precioMap: { [c: string]: number }[], importeMap: { [c: string]: number }[]): LayoutPartidas {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const FIX = 6;
  const colBarata = FIX + N * 2;
  const colDesv = colBarata + 1;
  const COLS = colDesv + 1;

  let cab1: (string | number)[] = ["COMPARATIVA POR PARTIDA", "", "", "", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); }
  cab1.push("COMPARACIÓN"); cab1.push("");
  let cab2: (string | number)[] = ["Código", "Tipo", "Nivel", "Resumen", "Ud", "Medición"];
  for (let i = 0; i < N; i++) { cab2.push("P. Unit"); cab2.push("Importe"); }
  cab2.push("Más barata"); cab2.push("Δ %");

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  for (let r = 0; r < orden.length; r++) {
    let cod = orden[r];
    let ci = info[cod];
    let esCap = ci.tipo === "Capitulo";

    let fila: (string | number)[] = [cod, ci.tipo, ci.nivel, ci.resumen, ci.unidad, esCap ? "" : r2(ci.medicion)];

    // Valores a comparar: precios unitarios en partidas, importes en capitulos
    let vals: number[] = [];
    let idxs: number[] = [];
    for (let i = 0; i < N; i++) {
      let tienePrecio = precioMap[i][cod] !== undefined;
      let tieneImporte = importeMap[i][cod] !== undefined;
      let pu = tienePrecio ? precioMap[i][cod] : 0;
      let im = tieneImporte ? importeMap[i][cod] : 0;
      if (esCap) {
        fila.push("");
        fila.push(tieneImporte ? r2(im) : "");
        if (tieneImporte && im > 0) { vals.push(im); idxs.push(i); }
      } else {
        fila.push(tienePrecio ? r2(pu) : "");
        fila.push(tieneImporte ? r2(im) : "");
        if (tienePrecio && pu > 0) { vals.push(pu); idxs.push(i); }
      }
    }
    let cmp = comparar(vals, idxs, nombres);
    fila.push(cmp.emp);
    fila.push(cmp.pct);
    filas.push(fila);
  }

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numFilas: orden.length, numCols: COLS, colMasBarata: colBarata, colDesv: colDesv };
}

// =====================================================================================
// HOJA OCULTA "_CONTRATO": metadatos para el Script B
// =====================================================================================
function escribirContrato(wb: ExcelScript.Workbook, nombreHoja: string,
  hojaResumen: string, hojaCapitulos: string, hojaPartidas: string, nombres: string[],
  layR: LayoutResumen, layC: LayoutCapitulos, layP: LayoutPartidas) {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;

  let pares: (string | number)[][] = [];
  pares.push(["VERSION", 1]);
  pares.push(["NUM_OFERTAS", N]);
  for (let i = 0; i < N; i++) { pares.push(["NOMBRE_" + (i + 1), nombres[i]]); }
  pares.push(["HOJA_RESUMEN", hojaResumen]);
  pares.push(["HOJA_CAPITULOS", hojaCapitulos]);
  pares.push(["HOJA_PARTIDAS", hojaPartidas]);

  // --- Hoja "Resumen ofertas" ---
  pares.push(["RES_HEADER_FILAS", 1]);
  pares.push(["RES_FILA_RESUMEN_INI", layR.filaResumenIni]);
  pares.push(["RES_NUM_COLS", layR.numCols]);
  pares.push(["RES_COL_EMPRESA", 0]);
  pares.push(["RES_COL_PEM", 1]);
  pares.push(["RES_COL_PEC", 2]);
  pares.push(["RES_COL_DIF_ECO_EUR", 3]);
  pares.push(["RES_COL_DIF_ECO_PCT", 4]);
  pares.push(["RES_COL_DIF_MED_EUR", 5]);
  pares.push(["RES_COL_DIF_MED_PCT", 6]);
  pares.push(["RES_COL_AVISO", 7]);

  // --- Hoja "Resumen capitulos" ---
  pares.push(["CAP_HEADER_FILAS", 2]);
  pares.push(["CAP_NUM_CAPITULOS", layC.numCapitulos]);
  pares.push(["CAP_FILA_TOTAL", layC.filaTotal]);
  pares.push(["CAP_NUM_COLS", layC.numCols]);
  pares.push(["CAP_COLS_FIJAS", 4]);
  pares.push(["CAP_COL_COD", 0]);
  pares.push(["CAP_COL_NOMBRE", 1]);
  pares.push(["CAP_COL_PEM", 2]);
  pares.push(["CAP_COL_MEDIA", 3]);
  pares.push(["CAP_COL_PRIMERA_OFERTA", 4]);
  pares.push(["CAP_COLS_POR_OFERTA", 3]);

  // --- Hoja "Comparativa partidas" ---
  pares.push(["PART_HEADER_FILAS", 2]);
  pares.push(["PART_NUM_FILAS", layP.numFilas]);
  pares.push(["PART_NUM_COLS", layP.numCols]);
  pares.push(["PART_COLS_FIJAS", 6]);
  pares.push(["PART_COL_CODIGO", 0]);
  pares.push(["PART_COL_TIPO", 1]);
  pares.push(["PART_COL_NIVEL", 2]);
  pares.push(["PART_COL_RESUMEN", 3]);
  pares.push(["PART_COL_UD", 4]);
  pares.push(["PART_COL_MEDICION", 5]);
  pares.push(["PART_COL_PRIMERA_OFERTA", 6]);
  pares.push(["PART_COLS_POR_OFERTA", 2]);
  pares.push(["PART_COL_MAS_BARATA", layP.colMasBarata]);
  pares.push(["PART_COL_DESV", layP.colDesv]);

  ws.getRangeByIndexes(0, 0, pares.length, 2).setValues(pares);
  ws.setVisibility(ExcelScript.SheetVisibility.hidden);
}

// =====================================================================================
// UTILIDADES
// =====================================================================================

// Compara una lista de valores y devuelve la empresa mas barata y la desviacion
// porcentual entre la mas cara y la mas barata: (max - min) / min.
function comparar(vals: number[], idxs: number[], nombres: string[]): Cmp {
  if (vals.length === 0) { return { emp: "", pct: 0 }; }
  if (vals.length === 1) { return { emp: nombres[idxs[0]], pct: 0 }; }
  let mn = vals[0], mx = vals[0], mni = 0;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i] < mn) { mn = vals[i]; mni = i; }
    if (vals[i] > mx) { mx = vals[i]; }
  }
  return { emp: nombres[idxs[mni]], pct: mn > 0 ? r4((mx - mn) / mn) : 0 };
}

// Quita extensiones .xlsx/.xlsm/.xls del nombre de la oferta, por si llegan sin limpiar
function limpiarNombre(n: string): string {
  let s = (n ? n : "").toString().trim();
  return s.replace(/\.(xlsx|xlsm|xls)$/i, "");
}

// Borra la hoja si existe y la vuelve a crear vacia (asi cada ejecucion parte de cero)
function recrear(wb: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  let ex = wb.getWorksheet(name);
  if (ex) { ex.delete(); }
  return wb.addWorksheet(name);
}

function pad2(t: string): string { return t.length === 1 ? "0" + t : t; }
function r2(x: number): number { return Math.round(x * 100) / 100; }
function r4(x: number): number { return Math.round(x * 10000) / 10000; }

// =====================================================================================
// INTERFACES (tipado estricto, sin any)
// =====================================================================================
interface Fila {
  tipo: string;      // "capitulo" | "partida"
  nivel: string;     // "padre" | "sub" | ""
  capitulo: string;
  codigo: string;
  resumen: string;
  unidad: string;
  medicion: number;
  precio: number;
  importe: number;
}
interface Oferta { nombre: string; filas: Fila[]; }
interface CodeInfo { tipo: string; nivel: string; resumen: string; unidad: string; medicion: number; }
interface Cmp { emp: string; pct: number; }
interface ExcepcionCapitulo { token: string; capitulo: string; }
interface LayoutResumen { filaResumenIni: number; numCols: number; }
interface LayoutCapitulos { numCapitulos: number; filaTotal: number; numCols: number; }
interface LayoutPartidas { numFilas: number; numCols: number; colMasBarata: number; colDesv: number; }
