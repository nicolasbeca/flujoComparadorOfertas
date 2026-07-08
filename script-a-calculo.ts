// =====================================================================================
// SCRIPT A — CÁLCULO DEL COMPARADOR DE OFERTAS (solo números, sin formato)
// =====================================================================================
// Recibe `ofertasJson` desde Power Automate y genera CUATRO hojas con SOLO valores:
//   1. "Resumen ofertas"       → totales PEC, desviaciones, aviso PEM/PEC y avisos de
//                                descuadre entre hojas
//   2. "Resumen capitulos"     → importes por capítulo padre y desviaciones vs media
//   3. "Resumen agrupado"      → capítulos agrupados por oficios (tabla editable)
//   4. "Comparativa partidas"  → una fila por partida/capítulo con precios e importes
//
// Este script NO aplica colores, ni merges, ni anchos, ni paneles inmovilizados.
// De la estética se encarga el SCRIPT B (formato), que se ejecuta después.
//
// -------------------------------------------------------------------------------------
// VALIDACIÓN CRÍTICA DE CUADRE
// -------------------------------------------------------------------------------------
// Para cada oferta se comprueba que coinciden (con una tolerancia configurable):
//   T1 = total PEC (suma de todas las partidas)            → "Resumen ofertas"
//   T2 = suma de todos los capítulos padre                 → "Resumen capitulos"
//   T3 = suma de todos los grupos de oficios               → "Resumen agrupado"
//   T4 = suma de partidas únicas de la comparativa         → "Comparativa partidas"
// Cualquier descuadre genera una fila "AVISO ..." al final de "Resumen ofertas".
// Además, los capítulos que no estén en la tabla de grupos NO se pierden: van a un
// grupo automático "Sin clasificar (revisar)" y se avisa en "Resumen agrupado".
// Un capítulo repetido en dos grupos también genera aviso (se cuenta solo en el 1º).
//
// -------------------------------------------------------------------------------------
// CONTRATO ENTRE SCRIPT A Y SCRIPT B
// -------------------------------------------------------------------------------------
// Este script deja una hoja OCULTA "_CONTRATO" con pares clave/valor (col A = clave,
// col B = valor). El Script B lee esa hoja y se orienta con ella sin adivinar nada.
// Todos los índices son 0-based (la primera fila/columna es 0).
//
// Claves escritas:
//   VERSION (=2), NUM_OFERTAS, NOMBRE_1..NOMBRE_n
//   HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_AGRUPADO, HOJA_PARTIDAS
//   RES_HEADER_FILAS, RES_FILA_RESUMEN_INI, RES_FILA_AVISO_INI, RES_NUM_AVISOS,
//   RES_NUM_COLS, RES_COL_EMPRESA, RES_COL_PEM, RES_COL_PEC, RES_COL_DIF_ECO_EUR,
//   RES_COL_DIF_ECO_PCT, RES_COL_DIF_MED_EUR, RES_COL_DIF_MED_PCT, RES_COL_AVISO
//   CAP_HEADER_FILAS, CAP_NUM_CAPITULOS, CAP_FILA_TOTAL, CAP_NUM_COLS, CAP_COLS_FIJAS,
//   CAP_COL_COD, CAP_COL_NOMBRE, CAP_COL_MEDIA, CAP_COL_PRIMERA_OFERTA, CAP_COLS_POR_OFERTA
//   AGR_HEADER_FILAS, AGR_NUM_GRUPOS, AGR_FILA_TOTAL, AGR_FILA_AVISO_INI, AGR_NUM_AVISOS,
//   AGR_NUM_COLS, AGR_COLS_FIJAS, AGR_COL_GRUPO, AGR_COL_CAPS, AGR_COL_MEDIA,
//   AGR_COL_PRIMERA_OFERTA, AGR_COLS_POR_OFERTA
//   PART_HEADER_FILAS, PART_NUM_FILAS, PART_NUM_COLS, PART_COLS_FIJAS, PART_COL_CODIGO,
//   PART_COL_TIPO, PART_COL_RESUMEN, PART_COL_UD, PART_COL_MEDICION,
//   PART_COL_PRIMERA_OFERTA, PART_COLS_POR_OFERTA, PART_COL_MAS_BARATA, PART_COL_DESV
//
// -------------------------------------------------------------------------------------
// RESTRICCIONES DEL RUNTIME (Office Scripts desde Power Automate)
// -------------------------------------------------------------------------------------
// Runtime síncrono. PROHIBIDO: .find(), .reduce(), .map(), .filter(), spread (...) y
// funciones flecha en callbacks de array. TODO con bucles `for` clásicos.
// Escritura SIEMPRE en bloque con setValues() (miles de filas).
// =====================================================================================

function main(workbook: ExcelScript.Workbook, ofertasJson?: string) {

  // ===================================================================================
  // ZONA DE CONFIGURACIÓN (editable)
  // ===================================================================================
  // Nombres de las hojas de salida (el Script B los lee del contrato, no hay que tocarlo)
  const HOJA_RESUMEN = "Resumen ofertas";
  const HOJA_CAPITULOS = "Resumen capitulos";
  const HOJA_AGRUPADO = "Resumen agrupado";
  const HOJA_PARTIDAS = "Comparativa partidas";
  const HOJA_CONTRATO = "_CONTRATO"; // hoja oculta de metadatos (NO renombrar sin cambiar B)

  // Umbral de anomalía PEM/PEC: si el total de una oferta está más de este porcentaje
  // POR DEBAJO de la media, se avisa (suele indicar oferta en PEM sin convertir a PEC).
  // 0.20 = 20 %. Es una red de seguridad, NO una garantía: verificar siempre a mano.
  const UMBRAL_ANOMALIA_PEM = 0.20;

  // Tolerancia (en €) para las comprobaciones de cuadre entre hojas. Cubre los céntimos
  // de redondeo; cualquier descuadre real (partidas o capítulos perdidos) es mucho mayor.
  const TOLERANCIA_DESCUADRE = 1.0;

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
  //  - orden / info: union de todos los codigos (capitulos y partidas) en el orden en
  //    que aparecen por primera vez (respeta el orden original del presupuesto).
  //  - precioMap / importeMap: por oferta, precio unitario e importe de cada codigo.
  //  - capOrden / capNombre / capImporte: capitulos PADRE deducidos y sus importes.
  //  - totalPEC: suma de importes de todas las partidas de cada oferta.
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
    let padreActual = "";

    for (let j = 0; j < fs.length; j++) {
      let f = fs[j];

      if (f.tipo === "capitulo") {
        // Los capitulos padre se llaman "Capitulo" y los subcapitulos "Subcapitulo"
        // (asi el Script B puede darles formato distinto leyendo la columna Tipo).
        let etiqueta = f.nivel === "padre" ? "Capitulo" : "Subcapitulo";
        if (!seen[f.codigo]) {
          seen[f.codigo] = true;
          orden.push(f.codigo);
          info[f.codigo] = { tipo: etiqueta, resumen: f.resumen, unidad: "", medicion: 0 };
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
        info[f.codigo] = { tipo: "Partida", resumen: f.resumen, unidad: f.unidad, medicion: f.medicion };
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
  // AGRUPACIÓN POR OFICIOS + VALIDACIÓN DE CUADRE ENTRE HOJAS
  // ===================================================================================
  let agr = calcularAgrupado(capOrden, capImporte, N);

  // T1..T4: los cuatro totales que DEBEN coincidir por oferta (ver cabecera del script)
  let avisosTotales: string[] = [];
  for (let i = 0; i < N; i++) {
    let t1 = totalPEC[i];
    let t2 = 0;
    for (let c = 0; c < capOrden.length; c++) {
      if (capImporte[i][capOrden[c]] !== undefined) { t2 += capImporte[i][capOrden[c]]; }
    }
    let t3 = 0;
    for (let g = 0; g < agr.grupos.length; g++) { t3 += agr.sumas[g][i]; }
    let t4 = 0;
    for (let r = 0; r < orden.length; r++) {
      if (info[orden[r]].tipo === "Partida" && importeMap[i][orden[r]] !== undefined) {
        t4 += importeMap[i][orden[r]];
      }
    }
    if (Math.abs(t2 - t1) > TOLERANCIA_DESCUADRE) {
      avisosTotales.push("AVISO " + nombres[i] + ": la suma de capitulos (" + r2(t2) + " EUR) no coincide con el total PEC (" + r2(t1) + " EUR). Diferencia: " + r2(t2 - t1) + " EUR. Puede haber partidas sin capitulo asignado.");
    }
    if (Math.abs(t3 - t2) > TOLERANCIA_DESCUADRE) {
      avisosTotales.push("AVISO " + nombres[i] + ": la suma de grupos (" + r2(t3) + " EUR) no coincide con la suma de capitulos (" + r2(t2) + " EUR). Diferencia: " + r2(t3 - t2) + " EUR. Revisa la tabla de grupos.");
    }
    if (Math.abs(t4 - t1) > TOLERANCIA_DESCUADRE) {
      avisosTotales.push("AVISO " + nombres[i] + ": la suma de partidas de la comparativa (" + r2(t4) + " EUR) no coincide con el total PEC (" + r2(t1) + " EUR). Diferencia: " + r2(t4 - t1) + " EUR. Hay codigos de partida repetidos dentro de la oferta (la comparativa solo muestra la primera aparicion).");
    }
  }

  // ===================================================================================
  // ESCRITURA DE LAS CUATRO HOJAS (solo valores) + CONTRATO
  // ===================================================================================
  let layR = hojaResumenA(workbook, HOJA_RESUMEN, nombres, totalPEC, UMBRAL_ANOMALIA_PEM, avisosTotales);
  let layC = hojaCapitulosA(workbook, HOJA_CAPITULOS, nombres, capOrden, capNombre, capImporte);
  let layG = hojaAgrupadoA(workbook, HOJA_AGRUPADO, nombres, agr);
  let layP = hojaPartidasA(workbook, HOJA_PARTIDAS, nombres, orden, info, precioMap, importeMap);

  escribirContrato(workbook, HOJA_CONTRATO, HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_AGRUPADO,
    HOJA_PARTIDAS, nombres, layR, layC, layG, layP);
}

// =====================================================================================
// TABLA DE GRUPOS (OFICIOS) — ¡¡CRITERIO DEL USUARIO, VERIFICAR EN CADA PROYECTO!!
// =====================================================================================
// Cada grupo suma los capitulos indicados. Los capitulos del proyecto que NO aparezcan
// aqui NO se pierden: van automaticamente al grupo "Sin clasificar (revisar)" y se
// genera un AVISO para que decidas donde colocarlos (p. ej. C05, C30, C45, C49...).
// Para editar: anade/quita g.push({ nombre: "...", capitulos: ["Cxx", ...] });
function tablaGrupos(): Grupo[] {
  let g: Grupo[] = [];
  g.push({ nombre: "Estructuras", capitulos: ["C01", "C02", "C03", "C04"] });
  g.push({ nombre: "Obra sucia", capitulos: ["C06", "C07", "C08"] });
  g.push({ nombre: "Carpintería y Cerrajería", capitulos: ["C09", "C10", "C11"] });
  g.push({ nombre: "Pinturas", capitulos: ["C12"] });
  g.push({ nombre: "Urbanización", capitulos: ["C13", "C14"] });
  g.push({ nombre: "Varios", capitulos: ["C15"] });
  // Instalaciones: C16 + todos los capitulos de C31 a C44 (ambos incluidos)
  let inst: string[] = ["C16"];
  for (let n = 31; n <= 44; n++) { inst.push("C" + String(n)); }
  g.push({ nombre: "Instalaciones", capitulos: inst });
  g.push({ nombre: "Cierre", capitulos: ["C46", "C47"] });
  return g;
}

// =====================================================================================
// TABLA DE EXCEPCIONES DE CAPÍTULO — ¡¡CRITERIO DEL USUARIO, VERIFICAR EN CADA PROYECTO!!
// =====================================================================================
// Algunos codigos no revelan su capitulo padre real. Esta tabla mapea el primer token
// numerico del codigo a un capitulo concreto:
//   - Telecomunicaciones: codigos que empiezan por 1., 2., 3. o 4  →  C33
//   - Aparatos sanitarios: codigos que empiezan por 23             →  C16
// Para anadir una excepcion: t.push({ token: "XX", capitulo: "Cnn" });
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
// Patrones observados: F2_Cnn... | F2_nn.... | Cnn... | nn. | nn  →  "Cnn"
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
// CÁLCULO DE LA AGRUPACIÓN POR OFICIOS
// =====================================================================================
// Reparte cada capitulo padre en su grupo. Garantias para que el cuadre NUNCA falle:
//  - Capitulo repetido en dos grupos → se cuenta SOLO en el primero + AVISO.
//  - Capitulo presente en las ofertas pero ausente de la tabla → grupo automatico
//    "Sin clasificar (revisar)" + AVISO. Asi la suma de grupos siempre es completa.
function calcularAgrupado(capOrden: string[], capImporte: { [c: string]: number }[], N: number): AgrupadoCalc {
  let base = tablaGrupos();
  let capGrupo: { [c: string]: number } = {};
  let avisos: string[] = [];

  for (let g = 0; g < base.length; g++) {
    let caps = base[g].capitulos;
    for (let j = 0; j < caps.length; j++) {
      if (capGrupo[caps[j]] !== undefined) {
        avisos.push("AVISO: el capitulo " + caps[j] + " aparece en los grupos \"" + base[capGrupo[caps[j]]].nombre + "\" y \"" + base[g].nombre + "\". Se contabiliza solo en el primero. Corrige la tabla de grupos.");
      } else {
        capGrupo[caps[j]] = g;
      }
    }
  }

  // Capitulos presentes en las ofertas que no estan en ningun grupo de la tabla
  let sinGrupo: string[] = [];
  for (let c = 0; c < capOrden.length; c++) {
    if (capGrupo[capOrden[c]] === undefined) {
      sinGrupo.push(capOrden[c]);
      capGrupo[capOrden[c]] = base.length; // indice del grupo automatico
    }
  }

  let grupos: Grupo[] = [];
  for (let g = 0; g < base.length; g++) { grupos.push(base[g]); }
  if (sinGrupo.length > 0) {
    grupos.push({ nombre: "Sin clasificar (revisar)", capitulos: sinGrupo });
    avisos.push("AVISO: capitulos fuera de la tabla de grupos: " + sinGrupo.join(", ") + ". Se han sumado en el grupo \"Sin clasificar (revisar)\" para que el total cuadre. Revisa la tabla de grupos del Script A.");
  }

  // Sumas por grupo y oferta (valores brutos; se redondea al escribir)
  let sumas: number[][] = [];
  let presencia: boolean[][] = [];
  for (let g = 0; g < grupos.length; g++) {
    let s: number[] = [];
    let p: boolean[] = [];
    for (let i = 0; i < N; i++) { s.push(0); p.push(false); }
    sumas.push(s);
    presencia.push(p);
  }
  for (let c = 0; c < capOrden.length; c++) {
    let cap = capOrden[c];
    let g = capGrupo[cap];
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) {
        sumas[g][i] += capImporte[i][cap];
        presencia[g][i] = true;
      }
    }
  }

  return { grupos: grupos, sumas: sumas, presencia: presencia, avisos: avisos };
}

// =====================================================================================
// HOJA 1: "Resumen ofertas" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: cabecera | Filas 1..N: ofertas | N+1: blanco | N+2..N+4: economica, media,
//   horquilla | (si hay avisos de cuadre) blanco + una fila de AVISO por descuadre
function hojaResumenA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  totalPEC: number[], umbralAnomalia: number, avisosTotales: string[]): LayoutResumen {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS = 8;

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
    // Deteccion de anomalia PEM/PEC: oferta muy por debajo de la media. Solo avisa.
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
  let filaResumenIni = filas.length;
  filas.push(["Oferta más económica", "", r2(mn), "", "", "", "", ""]);
  filas.push(["Media " + N + " ofertas", "", r2(media), "", "", "", "", ""]);
  filas.push(["Horquilla (máx − mín)", "", r2(mx - mn), "", "", "", "", ""]);

  // Avisos de descuadre entre hojas (la validacion mas importante del script)
  let filaAvisoIni = 0;
  if (avisosTotales.length > 0) {
    filas.push(["", "", "", "", "", "", "", ""]);
    filaAvisoIni = filas.length;
    for (let a = 0; a < avisosTotales.length; a++) {
      filas.push([avisosTotales[a], "", "", "", "", "", "", ""]);
    }
  }

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { filaResumenIni: filaResumenIni, filaAvisoIni: filaAvisoIni, numAvisos: avisosTotales.length, numCols: COLS };
}

// =====================================================================================
// HOJA 2: "Resumen capitulos" (solo valores, SIN columna PEM)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: nombres de empresa (primera columna de su bloque; B hace el merge)
//   Fila 1: cabecera | Filas 2..: capitulos padre | Ultima fila: TOTAL PEC
// Columnas: 0=Cod, 1=Capitulo, 2=Media PEC, luego 3 por oferta (Importe, Δ€, Δ%)
function hojaCapitulosA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  capOrden: string[], capNombre: { [c: string]: string },
  capImporte: { [c: string]: number }[]): LayoutCapitulos {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS_FIJAS = 3;
  const COLS = COLS_FIJAS + N * 3;

  let cab1: (string | number)[] = ["", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); cab1.push(""); }
  let cab2: (string | number)[] = ["Cód.", "Capítulo", "Media PEC"];
  for (let i = 0; i < N; i++) { cab2.push("Importe PEC"); cab2.push("Δ € s/Media"); cab2.push("Δ % s/Media"); }

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  let totCol: number[] = [];
  for (let i = 0; i < N; i++) { totCol.push(0); }

  for (let c = 0; c < capOrden.length; c++) {
    let cap = capOrden[c];
    let sumaCap = 0, cuenta = 0;
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) { sumaCap += r2(capImporte[i][cap]); cuenta++; }
    }
    let media = cuenta > 0 ? sumaCap / cuenta : 0;

    let fila: (string | number)[] = [cap, capNombre[cap] || "", r2(media)];
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
  let filaTotal = filas.length;
  let fT: (string | number)[] = ["", "TOTAL PEC", r2(mediaTot)];
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
// HOJA 3: "Resumen agrupado" (NUEVA, solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: nombres de empresa | Fila 1: cabecera | Filas 2..: un grupo (oficio) por
//   fila | fila TOTAL | (si hay avisos) blanco + filas de AVISO
// Columnas: 0=Grupo, 1=Capitulos que suma, 2=Media PEC, luego 3 por oferta
function hojaAgrupadoA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  agr: AgrupadoCalc): LayoutAgrupado {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS_FIJAS = 3;
  const COLS = COLS_FIJAS + N * 3;
  const G = agr.grupos.length;

  let cab1: (string | number)[] = ["", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); cab1.push(""); }
  let cab2: (string | number)[] = ["Grupo", "Capítulos", "Media PEC"];
  for (let i = 0; i < N; i++) { cab2.push("Importe PEC"); cab2.push("Δ € s/Media"); cab2.push("Δ % s/Media"); }

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  let totCol: number[] = [];
  for (let i = 0; i < N; i++) { totCol.push(0); }

  for (let g = 0; g < G; g++) {
    // Media del grupo entre las ofertas que tienen algun capitulo del grupo
    let sumaG = 0, cuenta = 0;
    for (let i = 0; i < N; i++) {
      if (agr.presencia[g][i]) { sumaG += r2(agr.sumas[g][i]); cuenta++; }
    }
    let media = cuenta > 0 ? sumaG / cuenta : 0;

    let fila: (string | number)[] = [agr.grupos[g].nombre, agr.grupos[g].capitulos.join("+"), r2(media)];
    for (let i = 0; i < N; i++) {
      if (agr.presencia[g][i]) {
        let v = r2(agr.sumas[g][i]);
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

  // Fila TOTAL
  let mediaTot = 0;
  for (let i = 0; i < N; i++) { mediaTot += totCol[i]; }
  mediaTot = mediaTot / N;
  let filaTotal = filas.length;
  let fT: (string | number)[] = ["TOTAL", "", r2(mediaTot)];
  for (let i = 0; i < N; i++) {
    fT.push(r2(totCol[i]));
    fT.push(r2(totCol[i] - mediaTot));
    fT.push(mediaTot > 0 ? r4((totCol[i] - mediaTot) / mediaTot) : 0);
  }
  filas.push(fT);

  // Avisos de la agrupacion (capitulos sin grupo o repetidos en dos grupos)
  let filaAvisoIni = 0;
  if (agr.avisos.length > 0) {
    let vacia: (string | number)[] = [];
    for (let c = 0; c < COLS; c++) { vacia.push(""); }
    filas.push(vacia);
    filaAvisoIni = filas.length;
    for (let a = 0; a < agr.avisos.length; a++) {
      let fA: (string | number)[] = [agr.avisos[a]];
      for (let c = 1; c < COLS; c++) { fA.push(""); }
      filas.push(fA);
    }
  }

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numGrupos: G, filaTotal: filaTotal, filaAvisoIni: filaAvisoIni, numAvisos: agr.avisos.length, numCols: COLS };
}

// =====================================================================================
// HOJA 4: "Comparativa partidas" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: titulo + nombres de empresa | Fila 1: cabecera | Filas 2..: capitulos y
//   partidas en el orden original
// Columnas fijas: 0=Codigo, 1=Tipo ("Capitulo"|"Subcapitulo"|"Partida"), 2=Resumen,
//                 3=Ud, 4=Medicion. Luego 2 por oferta (P. Unit, Importe) y al final
//                 Mas barata y Δ % (desviacion maxima entre ofertas).
function hojaPartidasA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  orden: string[], info: { [c: string]: CodeInfo },
  precioMap: { [c: string]: number }[], importeMap: { [c: string]: number }[]): LayoutPartidas {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const FIX = 5;
  const colBarata = FIX + N * 2;
  const colDesv = colBarata + 1;
  const COLS = colDesv + 1;

  let cab1: (string | number)[] = ["COMPARATIVA POR PARTIDA", "", "", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); }
  cab1.push("COMPARACIÓN"); cab1.push("");
  let cab2: (string | number)[] = ["Código", "Tipo", "Resumen", "Ud", "Medición"];
  for (let i = 0; i < N; i++) { cab2.push("P. Unit"); cab2.push("Importe"); }
  cab2.push("Más barata"); cab2.push("Δ %");

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  for (let r = 0; r < orden.length; r++) {
    let cod = orden[r];
    let ci = info[cod];
    let esCap = ci.tipo !== "Partida";

    let fila: (string | number)[] = [cod, ci.tipo, ci.resumen, ci.unidad, esCap ? "" : r2(ci.medicion)];

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
  hojaResumen: string, hojaCapitulos: string, hojaAgrupado: string, hojaPartidas: string,
  nombres: string[], layR: LayoutResumen, layC: LayoutCapitulos, layG: LayoutAgrupado,
  layP: LayoutPartidas) {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;

  let pares: (string | number)[][] = [];
  pares.push(["VERSION", 2]);
  pares.push(["NUM_OFERTAS", N]);
  for (let i = 0; i < N; i++) { pares.push(["NOMBRE_" + (i + 1), nombres[i]]); }
  pares.push(["HOJA_RESUMEN", hojaResumen]);
  pares.push(["HOJA_CAPITULOS", hojaCapitulos]);
  pares.push(["HOJA_AGRUPADO", hojaAgrupado]);
  pares.push(["HOJA_PARTIDAS", hojaPartidas]);

  // --- "Resumen ofertas" ---
  pares.push(["RES_HEADER_FILAS", 1]);
  pares.push(["RES_FILA_RESUMEN_INI", layR.filaResumenIni]);
  pares.push(["RES_FILA_AVISO_INI", layR.filaAvisoIni]);
  pares.push(["RES_NUM_AVISOS", layR.numAvisos]);
  pares.push(["RES_NUM_COLS", layR.numCols]);
  pares.push(["RES_COL_EMPRESA", 0]);
  pares.push(["RES_COL_PEM", 1]);
  pares.push(["RES_COL_PEC", 2]);
  pares.push(["RES_COL_DIF_ECO_EUR", 3]);
  pares.push(["RES_COL_DIF_ECO_PCT", 4]);
  pares.push(["RES_COL_DIF_MED_EUR", 5]);
  pares.push(["RES_COL_DIF_MED_PCT", 6]);
  pares.push(["RES_COL_AVISO", 7]);

  // --- "Resumen capitulos" (sin PEM) ---
  pares.push(["CAP_HEADER_FILAS", 2]);
  pares.push(["CAP_NUM_CAPITULOS", layC.numCapitulos]);
  pares.push(["CAP_FILA_TOTAL", layC.filaTotal]);
  pares.push(["CAP_NUM_COLS", layC.numCols]);
  pares.push(["CAP_COLS_FIJAS", 3]);
  pares.push(["CAP_COL_COD", 0]);
  pares.push(["CAP_COL_NOMBRE", 1]);
  pares.push(["CAP_COL_MEDIA", 2]);
  pares.push(["CAP_COL_PRIMERA_OFERTA", 3]);
  pares.push(["CAP_COLS_POR_OFERTA", 3]);

  // --- "Resumen agrupado" ---
  pares.push(["AGR_HEADER_FILAS", 2]);
  pares.push(["AGR_NUM_GRUPOS", layG.numGrupos]);
  pares.push(["AGR_FILA_TOTAL", layG.filaTotal]);
  pares.push(["AGR_FILA_AVISO_INI", layG.filaAvisoIni]);
  pares.push(["AGR_NUM_AVISOS", layG.numAvisos]);
  pares.push(["AGR_NUM_COLS", layG.numCols]);
  pares.push(["AGR_COLS_FIJAS", 3]);
  pares.push(["AGR_COL_GRUPO", 0]);
  pares.push(["AGR_COL_CAPS", 1]);
  pares.push(["AGR_COL_MEDIA", 2]);
  pares.push(["AGR_COL_PRIMERA_OFERTA", 3]);
  pares.push(["AGR_COLS_POR_OFERTA", 3]);

  // --- "Comparativa partidas" ---
  pares.push(["PART_HEADER_FILAS", 2]);
  pares.push(["PART_NUM_FILAS", layP.numFilas]);
  pares.push(["PART_NUM_COLS", layP.numCols]);
  pares.push(["PART_COLS_FIJAS", 5]);
  pares.push(["PART_COL_CODIGO", 0]);
  pares.push(["PART_COL_TIPO", 1]);
  pares.push(["PART_COL_RESUMEN", 2]);
  pares.push(["PART_COL_UD", 3]);
  pares.push(["PART_COL_MEDICION", 4]);
  pares.push(["PART_COL_PRIMERA_OFERTA", 5]);
  pares.push(["PART_COLS_POR_OFERTA", 2]);
  pares.push(["PART_COL_MAS_BARATA", layP.colMasBarata]);
  pares.push(["PART_COL_DESV", layP.colDesv]);

  ws.getRangeByIndexes(0, 0, pares.length, 2).setValues(pares);
  ws.setVisibility(ExcelScript.SheetVisibility.hidden);
}

// =====================================================================================
// UTILIDADES
// =====================================================================================

// Compara valores y devuelve la empresa mas barata y la desviacion porcentual entre
// la mas cara y la mas barata: (max - min) / min.
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

// Borra la hoja si existe y la vuelve a crear vacia (cada ejecucion parte de cero)
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
interface CodeInfo { tipo: string; resumen: string; unidad: string; medicion: number; }
interface Cmp { emp: string; pct: number; }
interface ExcepcionCapitulo { token: string; capitulo: string; }
interface Grupo { nombre: string; capitulos: string[]; }
interface AgrupadoCalc { grupos: Grupo[]; sumas: number[][]; presencia: boolean[][]; avisos: string[]; }
interface LayoutResumen { filaResumenIni: number; filaAvisoIni: number; numAvisos: number; numCols: number; }
interface LayoutCapitulos { numCapitulos: number; filaTotal: number; numCols: number; }
interface LayoutAgrupado { numGrupos: number; filaTotal: number; filaAvisoIni: number; numAvisos: number; numCols: number; }
interface LayoutPartidas { numFilas: number; numCols: number; colMasBarata: number; colDesv: number; }
