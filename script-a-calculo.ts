// =====================================================================================
// SCRIPT A — CÁLCULO DEL COMPARADOR DE OFERTAS (solo números, sin formato)
// =====================================================================================
// Recibe `ofertasJson` desde Power Automate y genera CUATRO hojas con SOLO valores:
//   1. "Resumen ofertas"       → PEC total por oferta, desviaciones y aviso PEM/PEC
//   2. "Resumen capitulos"     → capitulos padre con su importe oficial + aviso de cuadre
//   3. "Resumen agrupado"      → capitulos padre agrupados por oficios (tabla editable)
//   4. "Comparativa partidas"  → una fila por capitulo/subcapitulo/partida, orden original
//
// Este script NO aplica colores, ni merges, ni anchos, ni paneles inmovilizados.
// De la estética se encarga el SCRIPT B (formato), que se ejecuta después.
//
// -------------------------------------------------------------------------------------
// CRITERIO MAESTRO (aplica en TODAS las hojas de importes)
// -------------------------------------------------------------------------------------
//  1. Un capitulo PADRE es una fila con tipo:"capitulo" y nivel:"padre". Su importe es
//     SIEMPRE el de su propia fila: es el valor oficial declarado por la constructora
//     y ese numero manda.
//  2. Cada partida y cada subcapitulo pertenece al ULTIMO capitulo padre que aparecio
//     por encima en el ORDEN ORIGINAL (por POSICION, no por su codigo). Los codigos
//     pueden estar mal, repetidos o sin criterio: da igual, manda el orden de origen.
//     (Por eso ya NO existe la antigua tabla de excepciones por codigo: sobra.)
//  3. TOTAL PEC de una oferta = suma de los importes de las filas de sus capitulos
//     padre. Cuadra siempre por construccion.
//  4. COMPROBACION (no manda, solo avisa): por cada capitulo padre se suman los
//     importes de sus partidas. Si esa suma no casa con el importe de la fila del
//     capitulo (tolerancia configurable), se marca el capitulo con una "X" en la
//     columna "Aviso" de "Resumen capitulos". El importe usado sigue siendo SIEMPRE
//     el de la fila del capitulo, nunca el de las partidas.
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
//   CAP_COL_COD, CAP_COL_NOMBRE, CAP_COL_MEDIA, CAP_COL_PRIMERA_OFERTA,
//   CAP_COLS_POR_OFERTA, CAP_COL_AVISO (columna "X" de cuadre, al final)
//   AGR_HEADER_FILAS, AGR_NUM_GRUPOS, AGR_FILA_TOTAL, AGR_FILA_AVISO_INI, AGR_NUM_AVISOS,
//   AGR_NUM_COLS, AGR_COLS_FIJAS, AGR_COL_GRUPO, AGR_COL_CAPS, AGR_COL_MEDIA,
//   AGR_COL_PRIMERA_OFERTA, AGR_COLS_POR_OFERTA
//   PART_HEADER_FILAS, PART_NUM_FILAS, PART_NUM_COLS, PART_COLS_FIJAS,
//   PART_COL_MARCA (columna "X" de codigos repetidos, al inicio), PART_COL_CODIGO,
//   PART_COL_TIPO, PART_COL_RESUMEN, PART_COL_UD, PART_COL_MEDICION,
//   PART_COL_PRIMERA_OFERTA, PART_COLS_POR_OFERTA, PART_COL_MAS_BARATA, PART_COL_DESV
//
// El Script B actual (v2) funciona sin cambios: lee todos los indices del contrato,
// asi que el desplazamiento de columnas por la nueva columna de marca es transparente.
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
  // Nombres de las hojas de salida (el Script B los lee del contrato)
  const HOJA_RESUMEN = "Resumen ofertas";
  const HOJA_CAPITULOS = "Resumen capitulos";
  const HOJA_AGRUPADO = "Resumen agrupado";
  const HOJA_PARTIDAS = "Comparativa partidas";
  const HOJA_CONTRATO = "_CONTRATO"; // hoja oculta de metadatos (NO renombrar sin cambiar B)

  // Umbral de anomalía PEM/PEC: si el total de una oferta está más de este porcentaje
  // POR DEBAJO de la media, se avisa (suele indicar oferta en PEM sin convertir a PEC).
  // 0.20 = 20 %. Red de seguridad, NO garantía: verificar siempre a mano.
  const UMBRAL_ANOMALIA_PEM = 0.20;

  // Tolerancia (€) de la comprobación por capítulo: |suma de partidas − importe de la
  // fila del capítulo| mayor que esto → "X" en la columna Aviso de "Resumen capitulos".
  const TOLERANCIA_CUADRE_CAPITULO = 1.0;

  // Tolerancia (€) de la validación de grupos: |suma de grupos − TOTAL PEC| mayor que
  // esto → fila de AVISO en "Resumen agrupado". Cubre céntimos de redondeo.
  const TOLERANCIA_DESCUADRE = 1.0;

  // (Los colores se configuran en el Script B, que es quien pinta.)

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
  if (!ofertas || !Array.isArray(ofertas)) {
    throw new Error("ofertasJson debe ser un ARRAY de ofertas [ { nombre, filas }, ... ] y ha llegado otro tipo de JSON. Revisa el paso que compone ofertasJson en Power Automate.");
  }
  if (ofertas.length === 0) {
    throw new Error("El JSON no contiene ninguna oferta.");
  }

  // Nombres de oferta limpios (sin extension .xlsx/.xlsm/.xls, por si llegan sucios)
  let nombres: string[] = [];
  for (let i = 0; i < ofertas.length; i++) {
    nombres.push(limpiarNombre(ofertas[i].nombre));
  }
  const N = nombres.length;

  // ===================================================================================
  // AGREGACIÓN DE DATOS (criterio maestro: por POSICIÓN, el importe del padre manda)
  // ===================================================================================
  // Para la comparativa, cada fila se identifica con una CLAVE = tipo + codigo + numero
  // de ocurrencia dentro de la oferta. Asi, si un codigo se repite dentro de una misma
  // oferta, cada aparicion conserva su propia fila en su posicion de origen, y las
  // apariciones n-esimas de cada oferta se alinean entre si.
  let orden: string[] = [];                          // claves en orden de 1a aparicion
  let seen: { [k: string]: boolean } = {};
  let info: { [k: string]: CodeInfo } = {};
  let precioMap: { [k: string]: number }[] = [];     // por oferta: clave → precio unit
  let importeMap: { [k: string]: number }[] = [];    // por oferta: clave → importe
  let dupPartida: { [cod: string]: boolean } = {};   // codigos de partida repetidos en alguna oferta

  // Los capitulos padre se identifican por CLAVE = codigo + "|" + ocurrencia, no solo
  // por codigo: asi, si un codigo de padre viene repetido dentro de una oferta, cada
  // FILA padre cuenta por si misma (manda la posicion) y su importe suma al PEC.
  let capOrden: string[] = [];                       // claves de capitulos PADRE
  let capSeen: { [c: string]: boolean } = {};
  let capCodigo: { [c: string]: string } = {};       // clave → codigo visible
  let capNombre: { [c: string]: string } = {};
  let capImporte: { [c: string]: number }[] = [];    // por oferta: clave padre → importe de SU fila
  let sumaPartidas: { [c: string]: number }[] = [];  // por oferta: clave padre → suma de sus partidas
  let totalPEC: number[] = [];                       // por oferta: suma de filas padre
  let avisosDatos: string[] = [];                    // avisos de integridad de datos

  for (let k = 0; k < N; k++) {
    let o = ofertas[k];
    let fs: Fila[] = o.filas ? o.filas : [];
    if (fs.length === 0) {
      throw new Error("La oferta \"" + nombres[k] + "\" no contiene filas. Revisa su Excel de origen.");
    }
    let pM: { [key: string]: number } = {};
    let iM: { [key: string]: number } = {};
    let cImp: { [c: string]: number } = {};
    let sPart: { [c: string]: number } = {};
    let occCap: { [c: string]: number } = {};   // ocurrencias de codigos de capitulo
    let occPar: { [c: string]: number } = {};   // ocurrencias de codigos de partida
    let occPad: { [c: string]: number } = {};   // ocurrencias de codigos de capitulo PADRE
    let padreActual = "";                        // CLAVE del ultimo padre visto (posicion)
    let huerfanas = 0;                           // importe NETO de partidas antes del 1er padre
    let huerfanasAbs = 0;                        // suma de |importe| (evita compensaciones)

    for (let j = 0; j < fs.length; j++) {
      let f = fs[j];
      let cod = f.codigo ? f.codigo.toString().trim() : "";
      let res = f.resumen ? f.resumen.toString() : "";

      if (f.tipo === "capitulo") {
        // ----- Fila de capitulo (padre o subcapitulo) -----
        let nOcc = occCap[cod] === undefined ? 0 : occCap[cod];
        occCap[cod] = nOcc + 1;
        let clave = "C|" + cod + "|" + nOcc;
        let etiqueta = f.nivel === "padre" ? "Capitulo" : "Subcapitulo";
        if (!seen[clave]) {
          seen[clave] = true;
          orden.push(clave);
          info[clave] = { codigo: cod, tipo: etiqueta, resumen: res, unidad: "", medicion: 0 };
        }
        if (iM[clave] === undefined) { iM[clave] = f.importe; }

        if (f.nivel === "padre") {
          // El importe oficial del capitulo padre es el de ESTA fila. Si el codigo se
          // repite dentro de la oferta, cada aparicion es un capitulo DISTINTO (clave
          // codigo|ocurrencia): las apariciones n-esimas se alinean entre ofertas, y
          // TODAS las filas padre suman al PEC (criterio maestro: manda la posicion).
          let nPad = occPad[cod] === undefined ? 0 : occPad[cod];
          occPad[cod] = nPad + 1;
          let capClave = cod + "|" + nPad;
          padreActual = capClave;
          if (!capSeen[capClave]) {
            capSeen[capClave] = true;
            capOrden.push(capClave);
            capCodigo[capClave] = cod;
            capNombre[capClave] = res;
          }
          if (cImp[capClave] === undefined) { cImp[capClave] = f.importe; }
          if (nPad === 1) {
            avisosDatos.push("AVISO " + nombres[k] + ": el capitulo padre " + cod + " aparece mas de una vez; cada aparicion se trata como un capitulo distinto (alineadas por orden) y TODAS sus filas suman al PEC. Revisa el Excel de origen.");
          }
        }
        // Los subcapitulos NO se suman a nada: su detalle ya esta en las partidas y el
        // total oficial lo da la fila del padre.
        continue;
      }

      // ----- Fila de partida -----
      let nOccP = occPar[cod] === undefined ? 0 : occPar[cod];
      occPar[cod] = nOccP + 1;
      if (nOccP >= 1) { dupPartida[cod] = true; } // codigo repetido dentro de ESTA oferta
      let claveP = "P|" + cod + "|" + nOccP;
      if (!seen[claveP]) {
        seen[claveP] = true;
        orden.push(claveP);
        // Nota: resumen/unidad/medicion se toman de la PRIMERA oferta donde aparece.
        info[claveP] = { codigo: cod, tipo: "Partida", resumen: res, unidad: f.unidad, medicion: f.medicion };
      }
      if (pM[claveP] === undefined) { pM[claveP] = f.precio; iM[claveP] = f.importe; }

      // Pertenencia por POSICION: la partida cuelga del ultimo padre visto por encima.
      if (padreActual !== "") {
        sPart[padreActual] = (sPart[padreActual] || 0) + f.importe;
      } else {
        huerfanas += f.importe;
        huerfanasAbs += Math.abs(f.importe); // en absoluto: un descuento negativo o dos
                                             // partidas que se compensan tambien avisan
      }
    }

    if (huerfanasAbs > TOLERANCIA_CUADRE_CAPITULO) {
      avisosDatos.push("AVISO " + nombres[k] + ": hay partidas ANTES del primer capitulo padre (neto " + r2(huerfanas) + " EUR, " + r2(huerfanasAbs) + " EUR en valor absoluto); no cuentan en el PEC de ningun capitulo. Revisa el Excel de origen.");
    }

    // TOTAL PEC de la oferta = suma de los importes de las filas de capitulos padre
    let tot = 0;
    for (let c = 0; c < capOrden.length; c++) {
      if (cImp[capOrden[c]] !== undefined) { tot += cImp[capOrden[c]]; }
    }

    precioMap.push(pM);
    importeMap.push(iM);
    capImporte.push(cImp);
    sumaPartidas.push(sPart);
    totalPEC.push(tot);
  }

  // ===================================================================================
  // COMPROBACIÓN POR CAPÍTULO (no manda, solo avisa): partidas vs fila del padre
  // ===================================================================================
  // avisoCap[cap] = nombres de las ofertas donde la suma de partidas del capitulo no
  // casa con el importe de su fila (fuera de tolerancia).
  let avisoCap: { [c: string]: string[] } = {};
  for (let c = 0; c < capOrden.length; c++) {
    let cap = capOrden[c];
    let lista: string[] = [];
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] === undefined) { continue; }
      let sp = sumaPartidas[i][cap] === undefined ? 0 : sumaPartidas[i][cap];
      if (Math.abs(sp - capImporte[i][cap]) > TOLERANCIA_CUADRE_CAPITULO) {
        lista.push(nombres[i]);
      }
    }
    avisoCap[cap] = lista;
  }

  // ===================================================================================
  // AGRUPACIÓN POR OFICIOS + VALIDACIÓN CRÍTICA (grupos vs TOTAL PEC)
  // ===================================================================================
  let agr = calcularAgrupado(capOrden, capCodigo, capImporte, N);

  // Los avisos de integridad de datos tambien se muestran en "Resumen agrupado"
  for (let a = 0; a < avisosDatos.length; a++) { agr.avisos.push(avisosDatos[a]); }

  // Doble red: ademas de la garantia estructural (todo padre cae en un grupo, aunque
  // sea en "SIN CLASIFICAR"), se comprueba numericamente que grupos = PEC por oferta.
  for (let i = 0; i < N; i++) {
    let tG = 0;
    for (let g = 0; g < agr.grupos.length; g++) { tG += agr.sumas[g][i]; }
    if (Math.abs(tG - totalPEC[i]) > TOLERANCIA_DESCUADRE) {
      agr.avisos.push("AVISO " + nombres[i] + ": la suma de grupos (" + r2(tG) + " EUR) no coincide con el TOTAL PEC (" + r2(totalPEC[i]) + " EUR). Diferencia: " + r2(tG - totalPEC[i]) + " EUR. Revisa la tabla de grupos.");
    }
  }

  // ===================================================================================
  // ESCRITURA DE LAS CUATRO HOJAS (solo valores) + CONTRATO
  // ===================================================================================
  let layR = hojaResumenA(workbook, HOJA_RESUMEN, nombres, totalPEC, UMBRAL_ANOMALIA_PEM);
  let layC = hojaCapitulosA(workbook, HOJA_CAPITULOS, nombres, capOrden, capCodigo, capNombre, capImporte, avisoCap);
  let layG = hojaAgrupadoA(workbook, HOJA_AGRUPADO, nombres, agr);
  let layP = hojaPartidasA(workbook, HOJA_PARTIDAS, nombres, orden, info, precioMap, importeMap, dupPartida);

  escribirContrato(workbook, HOJA_CONTRATO, HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_AGRUPADO,
    HOJA_PARTIDAS, nombres, layR, layC, layG, layP);
}

// =====================================================================================
// TABLA DE GRUPOS (OFICIOS) — ¡¡CRITERIO DEL USUARIO, VERIFICAR EN CADA PROYECTO!!
// =====================================================================================
// Cada grupo suma los IMPORTES DE FILA de los capitulos padre indicados. Los padres del
// proyecto que NO aparezcan aqui NO se pierden: van automaticamente al grupo
// "SIN CLASIFICAR (REVISAR)" y se genera un AVISO (asi el total siempre cuadra y el
// descuadre nunca pasa en silencio). Nota: ya no deberia aparecer un C49 padre (ahora
// es subcapitulo dentro de C13), pero si apareciera cualquier padre no listado, caera
// en SIN CLASIFICAR y en el aviso.
// Para editar: anade/quita g.push({ nombre: "...", capitulos: ["Cxx", ...] });
function tablaGrupos(): Grupo[] {
  let g: Grupo[] = [];
  g.push({ nombre: "ESTRUCTURAS", capitulos: ["C01", "C02", "C03", "C04"] });
  g.push({ nombre: "OBRA SUCIA", capitulos: ["C05", "C06", "C07", "C08"] });
  g.push({ nombre: "CARPINTERÍA Y CERRAJERÍA", capitulos: ["C09", "C10", "C11"] });
  g.push({ nombre: "PINTURAS", capitulos: ["C12"] });
  g.push({ nombre: "URBANIZACIÓN", capitulos: ["C13", "C14"] });
  g.push({ nombre: "VARIOS", capitulos: ["C15"] });
  g.push({ nombre: "INSTALACIONES", capitulos: ["C16", "C30", "C31", "C32", "C33", "C38", "C39", "C40", "C41", "C43", "C44"] });
  g.push({ nombre: "CONTROL DE CALIDAD", capitulos: ["C45"] });
  g.push({ nombre: "SEGURIDAD Y SALUD", capitulos: ["C46"] });
  g.push({ nombre: "GESTIÓN DE RESIDUOS", capitulos: ["C47"] });
  return g;
}

// =====================================================================================
// CÁLCULO DE LA AGRUPACIÓN POR OFICIOS
// =====================================================================================
// Garantias para que el aviso NUNCA falle en detectar un descuadre:
//  - Capitulo repetido en dos grupos → se cuenta SOLO en el primero + AVISO.
//  - Capitulo padre presente en las ofertas pero ausente de la tabla → grupo automatico
//    "SIN CLASIFICAR (REVISAR)" + AVISO con los codigos y el importe afectado.
function calcularAgrupado(capOrden: string[], capCodigo: { [c: string]: string },
  capImporte: { [c: string]: number }[], N: number): AgrupadoCalc {
  let base = tablaGrupos();
  let codGrupo: { [cod: string]: number } = {};   // codigo de capitulo → indice de grupo
  let avisos: string[] = [];

  for (let g = 0; g < base.length; g++) {
    let caps = base[g].capitulos;
    for (let j = 0; j < caps.length; j++) {
      if (codGrupo[caps[j]] !== undefined) {
        avisos.push("AVISO: el capitulo " + caps[j] + " aparece en los grupos \"" + base[codGrupo[caps[j]]].nombre + "\" y \"" + base[g].nombre + "\". Se contabiliza solo en el primero. Corrige la tabla de grupos.");
      } else {
        codGrupo[caps[j]] = g;
      }
    }
  }

  // Capitulos padre presentes en las ofertas que no estan en ningun grupo de la tabla.
  // capOrden contiene CLAVES (codigo|ocurrencia); el grupo se decide por el CODIGO.
  let sinGrupo: string[] = [];
  let claveGrupo: { [clave: string]: number } = {};
  for (let c = 0; c < capOrden.length; c++) {
    let clave = capOrden[c];
    let g = codGrupo[capCodigo[clave]];
    if (g === undefined) {
      sinGrupo.push(capCodigo[clave]);
      claveGrupo[clave] = base.length; // indice del grupo automatico
    } else {
      claveGrupo[clave] = g;
    }
  }

  let grupos: Grupo[] = [];
  for (let g = 0; g < base.length; g++) { grupos.push(base[g]); }
  if (sinGrupo.length > 0) {
    grupos.push({ nombre: "SIN CLASIFICAR (REVISAR)", capitulos: sinGrupo });
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
    let g = claveGrupo[cap];
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) {
        sumas[g][i] += capImporte[i][cap];
        presencia[g][i] = true;
      }
    }
  }

  // Aviso de sin clasificar con el importe afectado (el mayor entre las ofertas)
  if (sinGrupo.length > 0) {
    let idxAuto = grupos.length - 1;
    let maxImporte = 0;
    for (let i = 0; i < N; i++) {
      if (sumas[idxAuto][i] > maxImporte) { maxImporte = sumas[idxAuto][i]; }
    }
    avisos.push("AVISO: capitulos padre fuera de la tabla de grupos: " + sinGrupo.join(", ") + ". Se han sumado en \"SIN CLASIFICAR (REVISAR)\" (hasta " + r2(maxImporte) + " EUR por oferta) para que el total cuadre. Revisa la tabla de grupos del Script A.");
  }

  return { grupos: grupos, sumas: sumas, presencia: presencia, avisos: avisos };
}

// =====================================================================================
// HOJA 1: "Resumen ofertas" (solo valores)
// =====================================================================================
// Estructura (0-based):
//   Fila 0: cabecera | Filas 1..N: ofertas | N+1: blanco | N+2..N+4: economica, media,
//   horquilla. El PEC de cada oferta es la suma de sus filas de capitulos padre.
function hojaResumenA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  totalPEC: number[], umbralAnomalia: number): LayoutResumen {

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

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { filaResumenIni: filaResumenIni, numCols: COLS };
}

// =====================================================================================
// HOJA 2: "Resumen capitulos" (solo valores, SIN columna PEM)
// =====================================================================================
// Solo capitulos PADRE. El importe de cada oferta es el de la FILA del capitulo
// (criterio maestro), nunca la suma de partidas.
// Estructura (0-based):
//   Fila 0: nombres de empresa (primera columna de su bloque; B hace el merge)
//   Fila 1: cabecera | Filas 2..: capitulos padre | Ultima fila: TOTAL PEC
// Columnas: 0=Cod, 1=Capitulo, 2=Media PEC, luego 3 por oferta (Importe, Δ€, Δ%),
//           y al FINAL una columna "Aviso" con "X" donde las partidas no casan con
//           la fila del capitulo (y entre parentesis, en que ofertas).
function hojaCapitulosA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  capOrden: string[], capCodigo: { [c: string]: string }, capNombre: { [c: string]: string },
  capImporte: { [c: string]: number }[], avisoCap: { [c: string]: string[] }): LayoutCapitulos {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS_FIJAS = 3;
  const COL_AVISO = COLS_FIJAS + N * 3;
  const COLS = COL_AVISO + 1;

  let cab1: (string | number)[] = ["", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); cab1.push(""); }
  cab1.push("");
  let cab2: (string | number)[] = ["Cód.", "Capítulo", "Media PEC"];
  for (let i = 0; i < N; i++) { cab2.push("Importe PEC"); cab2.push("Δ € s/Media"); cab2.push("Δ % s/Media"); }
  cab2.push("Aviso");

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

    // En la columna Cod se muestra el CODIGO (la clave interna lleva ademas la ocurrencia)
    let fila: (string | number)[] = [capCodigo[cap] || cap, capNombre[cap] || "", r2(media)];
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
    // "X" si en alguna oferta la suma de partidas no casa con la fila del capitulo
    let lista = avisoCap[cap] ? avisoCap[cap] : [];
    fila.push(lista.length > 0 ? "X (" + lista.join(", ") + ")" : "");
    filas.push(fila);
  }

  // Fila TOTAL PEC (suma de filas de capitulos padre = PEC oficial)
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
  fT.push("");
  filas.push(fT);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numCapitulos: capOrden.length, filaTotal: filaTotal, numCols: COLS, colAviso: COL_AVISO };
}

// =====================================================================================
// HOJA 3: "Resumen agrupado" (solo valores)
// =====================================================================================
// Capitulos padre agrupados por oficios (titulos EN MAYUSCULAS). Cada grupo suma los
// importes de FILA de sus capitulos padre. La suma de todos los grupos debe ser
// identica al TOTAL PEC; cualquier padre fuera de la tabla cae en "SIN CLASIFICAR
// (REVISAR)" y en una fila de AVISO.
// Estructura (0-based):
//   Fila 0: nombres de empresa | Fila 1: cabecera | Filas 2..: un grupo por fila |
//   fila TOTAL | (si hay avisos) blanco + filas de AVISO
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

  // Fila TOTAL (= TOTAL PEC si la tabla de grupos esta bien)
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

  // Filas de AVISO (sin clasificar, duplicados, huerfanas, descuadres...)
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
// Una fila por capitulo, subcapitulo y partida, EN ORDEN ORIGINAL. Los codigos
// repetidos dentro de una misma oferta conservan cada uno su fila y su posicion; esas
// filas de partida se marcan con "X" en la primera columna (solo aviso visual).
// Estructura (0-based):
//   Fila 0: titulo + nombres de empresa | Fila 1: cabecera | Filas 2..: datos
// Columnas fijas: 0=Aviso ("X" si codigo de partida repetido), 1=Codigo,
//                 2=Tipo ("Capitulo"|"Subcapitulo"|"Partida"), 3=Resumen, 4=Ud,
//                 5=Medicion. Luego 2 por oferta (P. Unit, Importe) y al final
//                 Mas barata (por precio unitario) y Δ % (max vs min).
function hojaPartidasA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  orden: string[], info: { [k: string]: CodeInfo },
  precioMap: { [k: string]: number }[], importeMap: { [k: string]: number }[],
  dupPartida: { [cod: string]: boolean }): LayoutPartidas {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const FIX = 6;
  const colBarata = FIX + N * 2;
  const colDesv = colBarata + 1;
  const COLS = colDesv + 1;

  let cab1: (string | number)[] = ["COMPARATIVA POR PARTIDA", "", "", "", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); }
  cab1.push("COMPARACIÓN"); cab1.push("");
  let cab2: (string | number)[] = ["Aviso", "Código", "Tipo", "Resumen", "Ud", "Medición"];
  for (let i = 0; i < N; i++) { cab2.push("P. Unit"); cab2.push("Importe"); }
  cab2.push("Más barata"); cab2.push("Δ %");

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  for (let r = 0; r < orden.length; r++) {
    let clave = orden[r];
    let ci = info[clave];
    let esCap = ci.tipo !== "Partida";

    // Marca "X": solo partidas cuyo codigo se repite dentro de alguna oferta
    let marca = (!esCap && dupPartida[ci.codigo]) ? "X" : "";

    let fila: (string | number)[] = [marca, ci.codigo, ci.tipo, ci.resumen, ci.unidad, esCap ? "" : r2(ci.medicion)];

    // Valores a comparar: precios unitarios en partidas, importes en capitulos
    let vals: number[] = [];
    let idxs: number[] = [];
    for (let i = 0; i < N; i++) {
      let tienePrecio = precioMap[i][clave] !== undefined;
      let tieneImporte = importeMap[i][clave] !== undefined;
      let pu = tienePrecio ? precioMap[i][clave] : 0;
      let im = tieneImporte ? importeMap[i][clave] : 0;
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
  pares.push(["RES_FILA_AVISO_INI", 0]);   // sin filas de aviso en esta hoja
  pares.push(["RES_NUM_AVISOS", 0]);
  pares.push(["RES_NUM_COLS", layR.numCols]);
  pares.push(["RES_COL_EMPRESA", 0]);
  pares.push(["RES_COL_PEM", 1]);
  pares.push(["RES_COL_PEC", 2]);
  pares.push(["RES_COL_DIF_ECO_EUR", 3]);
  pares.push(["RES_COL_DIF_ECO_PCT", 4]);
  pares.push(["RES_COL_DIF_MED_EUR", 5]);
  pares.push(["RES_COL_DIF_MED_PCT", 6]);
  pares.push(["RES_COL_AVISO", 7]);

  // --- "Resumen capitulos" (sin PEM; con columna Aviso al final) ---
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
  pares.push(["CAP_COL_AVISO", layC.colAviso]);

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

  // --- "Comparativa partidas" (con columna de marca "X" al inicio) ---
  pares.push(["PART_HEADER_FILAS", 2]);
  pares.push(["PART_NUM_FILAS", layP.numFilas]);
  pares.push(["PART_NUM_COLS", layP.numCols]);
  pares.push(["PART_COLS_FIJAS", 6]);
  pares.push(["PART_COL_MARCA", 0]);
  pares.push(["PART_COL_CODIGO", 1]);
  pares.push(["PART_COL_TIPO", 2]);
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

// Borra la hoja si existe y la vuelve a crear vacia (cada ejecucion parte de cero).
// OJO al orden: se CREA primero la hoja nueva (nombre temporal) y se borra despues la
// antigua. Si se borrara primero y fuese la unica hoja visible del libro, la API de
// Office Scripts lanzaria un error en runtime ("cannot delete the only visible sheet").
function recrear(wb: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  let nueva = wb.addWorksheet();
  let ex = wb.getWorksheet(name);
  if (ex) { ex.delete(); }
  nueva.setName(name);
  return nueva;
}

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
interface CodeInfo { codigo: string; tipo: string; resumen: string; unidad: string; medicion: number; }
interface Cmp { emp: string; pct: number; }
interface Grupo { nombre: string; capitulos: string[]; }
interface AgrupadoCalc { grupos: Grupo[]; sumas: number[][]; presencia: boolean[][]; avisos: string[]; }
interface LayoutResumen { filaResumenIni: number; numCols: number; }
interface LayoutCapitulos { numCapitulos: number; filaTotal: number; numCols: number; colAviso: number; }
interface LayoutAgrupado { numGrupos: number; filaTotal: number; filaAvisoIni: number; numAvisos: number; numCols: number; }
interface LayoutPartidas { numFilas: number; numCols: number; colMasBarata: number; colDesv: number; }
