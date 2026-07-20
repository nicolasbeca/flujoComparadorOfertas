// =====================================================================================
// SCRIPT A — CÁLCULO DEL COMPARADOR DE OFERTAS (solo números, sin formato)
// =====================================================================================
// Recibe `ofertasJson` desde Power Automate y genera SEIS hojas con SOLO valores:
//   1. "Portada"               → resumen ejecutivo (nº ofertas, ganadora, ahorro, fecha)
//   2. "Resumen ofertas"       → PEC por oferta ORDENADO de mas barata a mas cara,
//                                ranking, avisos PEM/PEC y de oferta incompleta
//   3. "Resumen capitulos"     → capitulos padre con su importe oficial + aviso de cuadre
//   4. "Resumen agrupado"      → capitulos padre agrupados por oficios (tabla editable)
//   5. "Comparativa partidas"  → una fila por capitulo/subcapitulo/partida, orden
//                                original, con columnas Alcance y Desv. s/media
//   6. "Top desviaciones"      → las N partidas donde mas difieren las ofertas (en EUR)
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
// REDONDEO Y CUADRE (criterio unico para evitar falsos descuadres)
// -------------------------------------------------------------------------------------
// TODOS los importes y precios se redondean a 2 decimales AL LEER el JSON (r2), y todas
// las sumas se hacen sobre esos valores ya redondeados. Asi, "Resumen ofertas",
// "Resumen capitulos" y "Resumen agrupado" suman exactamente lo mismo por construccion
// y la fila de CUADRE no da falsos positivos por arrastres de decimales.
// La fila de CUADRE visible (al pie de cada hoja de importes) compara por oferta:
//   TOTAL PEC (resumen) vs suma de capitulos vs suma de grupos vs suma de partidas de
//   "Comparativa partidas". Si todo casa (tolerancia configurable) → "CUADRE OK";
//   si no → "DESCUADRE: X € — detalle". OJO: la suma de partidas puede descuadrar
//   LEGITIMAMENTE si la constructora declara en la fila del capitulo un importe que no
//   es la suma de sus partidas; ese caso tambien debe verse (nunca en silencio).
//
// -------------------------------------------------------------------------------------
// CONTRATO ENTRE SCRIPT A Y SCRIPT B (VERSION 3)
// -------------------------------------------------------------------------------------
// Este script deja una hoja OCULTA "_CONTRATO" con pares clave/valor (col A = clave,
// col B = valor). El Script B lee esa hoja y se orienta con ella sin adivinar nada.
// Todos los índices son 0-based (la primera fila/columna es 0).
//
// Claves escritas:
//   VERSION (=3), NUM_OFERTAS, NOMBRE_1..NOMBRE_n
//   HOJA_PORTADA, HOJA_RESUMEN, HOJA_CAPITULOS, HOJA_AGRUPADO, HOJA_PARTIDAS, HOJA_TOP
//
//   Portada ("Portada", 2 columnas etiqueta/valor):
//     POR_NUM_FILAS, POR_COL_ETIQUETA, POR_COL_VALOR, POR_FILA_CUADRE
//
//   "Resumen ofertas" (filas ORDENADAS por PEC ascendente; nueva col Ranking en 0):
//     RES_HEADER_FILAS, RES_FILA_RESUMEN_INI, RES_FILA_AVISO_INI, RES_NUM_AVISOS,
//     RES_NUM_COLS, RES_COL_RANKING, RES_COL_EMPRESA, RES_COL_PEM, RES_COL_PEC,
//     RES_COL_DIF_ECO_EUR, RES_COL_DIF_ECO_PCT, RES_COL_DIF_MED_EUR,
//     RES_COL_DIF_MED_PCT, RES_COL_AVISO, RES_FILA_CUADRE, RES_COL_CUADRE
//     OJO B: el orden de las filas de oferta ya NO es el de NOMBRE_1..n, es por PEC
//     ascendente; la empresa de cada fila se lee de RES_COL_EMPRESA.
//
//   "Resumen capitulos":
//     CAP_HEADER_FILAS, CAP_NUM_CAPITULOS, CAP_FILA_TOTAL, CAP_NUM_COLS, CAP_COLS_FIJAS,
//     CAP_COL_COD, CAP_COL_NOMBRE, CAP_COL_MEDIA, CAP_COL_PRIMERA_OFERTA,
//     CAP_COLS_POR_OFERTA, CAP_COL_AVISO (columna "X" de cuadre, al final),
//     CAP_FILA_CUADRE, CAP_COL_CUADRE
//
//   "Resumen agrupado":
//     AGR_HEADER_FILAS, AGR_NUM_GRUPOS, AGR_FILA_TOTAL, AGR_FILA_AVISO_INI,
//     AGR_NUM_AVISOS, AGR_NUM_COLS, AGR_COLS_FIJAS, AGR_COL_GRUPO, AGR_COL_CAPS,
//     AGR_COL_MEDIA, AGR_COL_PRIMERA_OFERTA, AGR_COLS_POR_OFERTA,
//     AGR_FILA_CUADRE, AGR_COL_CUADRE
//
//   "Comparativa partidas" (nuevas columnas Alcance y Desv. s/media al final, y fila
//   TOTAL partidas + fila CUADRE al pie):
//     PART_HEADER_FILAS, PART_NUM_FILAS, PART_NUM_COLS, PART_COLS_FIJAS,
//     PART_COL_MARCA (columna "X" de codigos repetidos, al inicio), PART_COL_CODIGO,
//     PART_COL_TIPO, PART_COL_RESUMEN, PART_COL_UD, PART_COL_MEDICION,
//     PART_COL_PRIMERA_OFERTA, PART_COLS_POR_OFERTA, PART_COL_MAS_BARATA, PART_COL_DESV,
//     PART_COL_ALCANCE ("SOLO <oferta>" si solo cotiza una),
//     PART_COL_DESV_MEDIA (texto compacto con las ofertas que se desvian > umbral de la
//     media de precio unitario de la partida, p. ej. "ACME +32% | BETA -28%"),
//     PART_FILA_TOTAL (fila "TOTAL partidas"), PART_FILA_CUADRE, PART_COL_CUADRE
//
//   "Top desviaciones" (1 fila de cabecera; 1 columna de importe por oferta):
//     TOP_HEADER_FILAS, TOP_NUM_FILAS, TOP_NUM_COLS, TOP_COL_CODIGO, TOP_COL_RESUMEN,
//     TOP_COL_PRIMERA_OFERTA, TOP_COLS_POR_OFERTA, TOP_COL_DIF_EUR, TOP_COL_DIF_PCT,
//     TOP_COL_MAS_BARATA
//
//   Celdas de CUADRE (*_FILA_CUADRE / *_COL_CUADRE): la celda contiene "CUADRE OK" o
//   un texto que empieza por "DESCUADRE:". El Script B debe pintarla en verde si
//   empieza por "CUADRE OK" y en rojo si empieza por "DESCUADRE".
//
// El Script B v2 sigue funcionando: lee todos los indices del contrato, asi que el
// desplazamiento de columnas (Ranking, Alcance, Desv. s/media) le es transparente;
// las claves nuevas solo las usa un B actualizado.
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
  const HOJA_PORTADA = "Portada";
  const HOJA_RESUMEN = "Resumen ofertas";
  const HOJA_CAPITULOS = "Resumen capitulos";
  const HOJA_AGRUPADO = "Resumen agrupado";
  const HOJA_PARTIDAS = "Comparativa partidas";
  const HOJA_TOP = "Top desviaciones";
  const HOJA_CONTRATO = "_CONTRATO"; // hoja oculta de metadatos (NO renombrar sin cambiar B)

  // Umbral de anomalía PEM/PEC: si el total de una oferta está más de este porcentaje
  // POR DEBAJO de la media de las DEMÁS, se avisa (suele indicar oferta en PEM sin
  // convertir a PEC). 0.20 = 20 %. Red de seguridad, NO garantía: verificar a mano.
  const UMBRAL_ANOMALIA_PEM = 0.20;

  // Umbral de oferta incompleta: si el nº de partidas de una oferta es inferior a este
  // porcentaje de la media de las DEMÁS ofertas, se avisa (posible export incompleto).
  // 0.5 = menos del 50 % de la media.
  const UMBRAL_OFERTA_INCOMPLETA = 0.5;

  // Nº de partidas de la hoja "Top desviaciones" (las de mayor diferencia en EUR).
  const TOP_DESVIACIONES = 20;

  // Umbral de la columna compacta "Desv. s/media" de "Comparativa partidas": solo se
  // listan las ofertas cuyo precio unitario se desvia mas de esto de la media de la
  // partida. 0.25 = 25 %.
  const UMBRAL_DESV_MEDIA = 0.25;

  // Tolerancia (€) de la comprobación por capítulo: |suma de partidas − importe de la
  // fila del capítulo| mayor que esto → "X" en la columna Aviso de "Resumen capitulos".
  const TOLERANCIA_CUADRE_CAPITULO = 1.0;

  // Tolerancia (€) de la validación de grupos: |suma de grupos − TOTAL PEC| mayor que
  // esto → fila de AVISO en "Resumen agrupado". Cubre céntimos de redondeo.
  const TOLERANCIA_DESCUADRE = 1.0;

  // Tolerancia (€) de la fila de CUADRE visible: diferencia maxima admitida entre
  // TOTAL PEC, suma de capitulos, suma de grupos y suma de partidas por oferta.
  const TOLERANCIA_CUADRE_TOTAL = 1.0;

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
  // REDONDEO: todos los importes y precios se pasan por r2 AL LEER (ver cabecera).
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
  let totalPartidas: number[] = [];                  // por oferta: suma de TODAS sus partidas
  let numPartidas: number[] = [];                    // por oferta: nº de partidas (para aviso incompleta)
  let avisosDatos: string[] = [];                    // avisos de integridad de datos
  let nombreObra = "";                               // nombre de obra si alguna oferta lo trae

  for (let k = 0; k < N; k++) {
    let o = ofertas[k];
    let fs: Fila[] = o.filas ? o.filas : [];
    if (fs.length === 0) {
      throw new Error("La oferta \"" + nombres[k] + "\" no contiene filas. Revisa su Excel de origen.");
    }
    if (nombreObra === "" && o.obra && o.obra.toString().trim().length > 0) {
      nombreObra = o.obra.toString().trim();
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
    let totPart = 0;                             // suma de todas las partidas de la oferta
    let nPart = 0;                               // nº de partidas de la oferta

    for (let j = 0; j < fs.length; j++) {
      let f = fs[j];
      let cod = f.codigo ? f.codigo.toString().trim() : "";
      let res = f.resumen ? f.resumen.toString() : "";
      // Redondeo unico al leer: todo lo que se sume despues ya va a 2 decimales.
      let imp = r2(f.importe ? f.importe : 0);
      let pre = r2(f.precio ? f.precio : 0);

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
        if (iM[clave] === undefined) { iM[clave] = imp; }

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
          if (cImp[capClave] === undefined) { cImp[capClave] = imp; }
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
      if (pM[claveP] === undefined) { pM[claveP] = pre; iM[claveP] = imp; }

      totPart += imp;
      nPart++;

      // Pertenencia por POSICION: la partida cuelga del ultimo padre visto por encima.
      if (padreActual !== "") {
        sPart[padreActual] = (sPart[padreActual] || 0) + imp;
      } else {
        huerfanas += imp;
        huerfanasAbs += Math.abs(imp); // en absoluto: un descuento negativo o dos
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
    totalPEC.push(r2(tot));
    totalPartidas.push(r2(totPart));
    // nº de partidas: se usa el numPartidas del lector si viene; si no, el recuento propio
    numPartidas.push(typeof o.numPartidas === "number" ? o.numPartidas : nPart);
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
      if (Math.abs(r2(sp) - capImporte[i][cap]) > TOLERANCIA_CUADRE_CAPITULO) {
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
    if (Math.abs(r2(tG) - totalPEC[i]) > TOLERANCIA_DESCUADRE) {
      agr.avisos.push("AVISO " + nombres[i] + ": la suma de grupos (" + r2(tG) + " EUR) no coincide con el TOTAL PEC (" + r2(totalPEC[i]) + " EUR). Diferencia: " + r2(tG - totalPEC[i]) + " EUR. Revisa la tabla de grupos.");
    }
  }

  // ===================================================================================
  // CUADRE GLOBAL VISIBLE — VALIDACIÓN CRÍTICA (nunca en silencio)
  // ===================================================================================
  // Por oferta se comparan los cuatro totales: TOTAL PEC (resumen), suma de capitulos,
  // suma de grupos y suma de partidas de "Comparativa partidas". El texto resultante
  // se escribe al pie de las cuatro hojas de importes y el Script B lo pinta.
  let textoCuadre = calcularTextoCuadre(nombres, totalPEC, capOrden, capImporte,
    agr, totalPartidas, TOLERANCIA_CUADRE_TOTAL);

  // ===================================================================================
  // ESCRITURA DE LAS SEIS HOJAS (solo valores) + CONTRATO
  // ===================================================================================
  let layR = hojaResumenA(workbook, HOJA_RESUMEN, nombres, totalPEC, numPartidas,
    UMBRAL_ANOMALIA_PEM, UMBRAL_OFERTA_INCOMPLETA, textoCuadre);
  let layC = hojaCapitulosA(workbook, HOJA_CAPITULOS, nombres, capOrden, capCodigo,
    capNombre, capImporte, avisoCap, textoCuadre);
  let layG = hojaAgrupadoA(workbook, HOJA_AGRUPADO, nombres, agr, textoCuadre);
  let layP = hojaPartidasA(workbook, HOJA_PARTIDAS, nombres, orden, info, precioMap,
    importeMap, dupPartida, UMBRAL_DESV_MEDIA, textoCuadre);
  let layT = hojaTopA(workbook, HOJA_TOP, nombres, orden, info, importeMap, TOP_DESVIACIONES);
  let layPor = hojaPortadaA(workbook, HOJA_PORTADA, nombres, totalPEC, nombreObra, textoCuadre);

  escribirContrato(workbook, HOJA_CONTRATO, HOJA_PORTADA, HOJA_RESUMEN, HOJA_CAPITULOS,
    HOJA_AGRUPADO, HOJA_PARTIDAS, HOJA_TOP, nombres, layR, layC, layG, layP, layT, layPor);
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

  // Sumas por grupo y oferta (los importes ya llegan redondeados a 2 decimales)
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
// CUADRE GLOBAL: texto de la celda de cuadre visible
// =====================================================================================
// Compara por oferta los cuatro totales (PEC, capitulos, grupos, partidas). Devuelve
// "CUADRE OK" si todo casa, o "DESCUADRE: X € — detalle" con la peor diferencia y el
// detalle por oferta (que total se separa mas del PEC oficial).
function calcularTextoCuadre(nombres: string[], totalPEC: number[], capOrden: string[],
  capImporte: { [c: string]: number }[], agr: AgrupadoCalc, totalPartidas: number[],
  tolerancia: number): string {

  const N = nombres.length;
  let detalles: string[] = [];
  let peorGlobal = 0;

  for (let i = 0; i < N; i++) {
    // Suma independiente de capitulos (lo que muestra "Resumen capitulos")
    let sCap = 0;
    for (let c = 0; c < capOrden.length; c++) {
      if (capImporte[i][capOrden[c]] !== undefined) { sCap += capImporte[i][capOrden[c]]; }
    }
    sCap = r2(sCap);
    // Suma independiente de grupos (lo que muestra "Resumen agrupado")
    let sGru = 0;
    for (let g = 0; g < agr.grupos.length; g++) { sGru += agr.sumas[g][i]; }
    sGru = r2(sGru);
    // Suma de partidas (lo que muestra la fila TOTAL de "Comparativa partidas")
    let sPar = totalPartidas[i];

    let dCap = Math.abs(sCap - totalPEC[i]);
    let dGru = Math.abs(sGru - totalPEC[i]);
    let dPar = Math.abs(sPar - totalPEC[i]);

    // La peor diferencia de esta oferta y de que hoja procede
    let peor = dCap;
    let hojaPeor = "suma de capitulos";
    let valPeor = sCap;
    if (dGru > peor) { peor = dGru; hojaPeor = "suma de grupos"; valPeor = sGru; }
    if (dPar > peor) { peor = dPar; hojaPeor = "suma de partidas"; valPeor = sPar; }

    if (peor > tolerancia) {
      if (peor > peorGlobal) { peorGlobal = peor; }
      detalles.push(nombres[i] + ": " + hojaPeor + " " + r2(valPeor) + " EUR vs PEC " + r2(totalPEC[i]) + " EUR (dif " + r2(valPeor - totalPEC[i]) + " EUR)");
    }
  }

  if (detalles.length === 0) {
    return "CUADRE OK";
  }
  return "DESCUADRE: " + r2(peorGlobal) + " € — " + detalles.join(" | ");
}

// =====================================================================================
// HOJA "Portada": resumen ejecutivo (solo valores, 2 columnas etiqueta/valor)
// =====================================================================================
// El nombre de obra se toma del campo opcional `obra` de las ofertas si viene; si no,
// se deja el hueco vacio para rellenarlo a mano.
function hojaPortadaA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  totalPEC: number[], nombreObra: string, textoCuadre: string): LayoutPortada {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;

  let mn = totalPEC[0], mx = totalPEC[0], suma = 0, idxMin = 0;
  for (let i = 0; i < N; i++) {
    let v = totalPEC[i];
    if (v < mn) { mn = v; idxMin = i; }
    if (v > mx) { mx = v; }
    suma += v;
  }
  let media = suma / N;

  let filas: (string | number)[][] = [];
  filas.push(["COMPARATIVA DE OFERTAS — RESUMEN EJECUTIVO", ""]);
  filas.push(["", ""]);
  filas.push(["Obra", nombreObra]); // hueco si no viene en los datos
  filas.push(["Fecha de generación", fechaHoy()]);
  filas.push(["Nº de ofertas comparadas", N]);
  filas.push(["Oferta más económica", nombres[idxMin]]);
  filas.push(["Importe más económico (PEC)", r2(mn)]);
  filas.push(["Media de las " + N + " ofertas (PEC)", r2(media)]);
  filas.push(["Ahorro más económica vs media (€)", r2(media - mn)]);
  filas.push(["Ahorro más económica vs media (%)", media > 0 ? r4((media - mn) / media) : 0]);
  filas.push(["Horquilla (máx − mín)", r2(mx - mn)]);
  filas.push(["", ""]);
  let filaCuadre = filas.length;
  filas.push(["Estado de cuadre", textoCuadre]);

  ws.getRangeByIndexes(0, 0, filas.length, 2).setValues(filas);
  ws.setPosition(0); // la portada, primera del libro
  return { numFilas: filas.length, filaCuadre: filaCuadre };
}

// =====================================================================================
// HOJA "Resumen ofertas" (solo valores) — ORDENADA de mas barata a mas cara
// =====================================================================================
// Estructura (0-based):
//   Fila 0: cabecera | Filas 1..N: ofertas ORDENADAS por PEC ascendente (ganadora
//   arriba, con su ranking) | N+1: blanco | N+2..N+4: economica, media, horquilla |
//   N+5: blanco | N+6: fila de CUADRE visible.
// La columna Aviso concentra: anomalia PEM/PEC (total muy por debajo de la media de
// las demas) y oferta posiblemente incompleta (numPartidas muy inferior a la media
// de las demas). Ambos umbrales en la zona de configuracion.
function hojaResumenA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  totalPEC: number[], numPartidas: number[], umbralAnomalia: number,
  umbralIncompleta: number, textoCuadre: string): LayoutResumen {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const COLS = 9;

  let mn = totalPEC[0], mx = totalPEC[0], suma = 0;
  for (let i = 0; i < N; i++) {
    let v = totalPEC[i];
    if (v < mn) { mn = v; }
    if (v > mx) { mx = v; }
    suma += v;
  }
  let media = suma / N;

  // Orden ascendente por PEC (la mas barata primero). Ordenacion por seleccion con
  // bucles clasicos (sin callbacks, prohibidos en este runtime).
  let ordenIdx = ordenAscendente(totalPEC);

  let filas: (string | number)[][] = [];
  filas.push(["Ranking", "EMPRESA", "PEM (manual)", "IMPORTE PEC", "Δ vs +econ. (€)", "Δ vs +econ. (%)", "Δ vs media (€)", "Δ vs media (%)", "Aviso"]);

  for (let p = 0; p < N; p++) {
    let i = ordenIdx[p];
    let v = totalPEC[i];
    let avisos: string[] = [];

    // Aviso PEM/PEC: total muy por debajo de la MEDIA DE LAS DEMAS ofertas.
    if (N >= 2) {
      let sumaOtras = 0;
      for (let j = 0; j < N; j++) { if (j !== i) { sumaOtras += totalPEC[j]; } }
      let mediaOtras = sumaOtras / (N - 1);
      if (mediaOtras > 0) {
        let pctBajo = (mediaOtras - v) / mediaOtras;
        if (pctBajo > umbralAnomalia) {
          let pctTxt = Math.round(pctBajo * 1000) / 10;
          avisos.push("AVISO: importe un " + pctTxt + "% por debajo de la media del resto. Posible PEM sin convertir a PEC — revisar.");
        }
      }
    }

    // Aviso de oferta incompleta: numPartidas muy inferior a la media de las demas.
    if (N >= 2) {
      let sumaP = 0;
      for (let j = 0; j < N; j++) { if (j !== i) { sumaP += numPartidas[j]; } }
      let mediaP = sumaP / (N - 1);
      if (mediaP > 0 && numPartidas[i] < mediaP * umbralIncompleta) {
        avisos.push("AVISO: solo " + numPartidas[i] + " partidas frente a una media de " + Math.round(mediaP) + " en el resto. Posible oferta incompleta o mal exportada — revisar.");
      }
    }

    filas.push([
      p + 1, // ranking: 1 = mas barata
      nombres[i],
      "", // PEM se rellena a mano
      r2(v),
      r2(v - mn),
      mn > 0 ? r4((v - mn) / mn) : 0,
      r2(v - media),
      media > 0 ? r4((v - media) / media) : 0,
      avisos.join(" | ")
    ]);
  }

  filas.push(["", "", "", "", "", "", "", "", ""]);
  let filaResumenIni = filas.length;
  filas.push(["", "Oferta más económica", "", r2(mn), "", "", "", "", ""]);
  filas.push(["", "Media " + N + " ofertas", "", r2(media), "", "", "", "", ""]);
  filas.push(["", "Horquilla (máx − mín)", "", r2(mx - mn), "", "", "", "", ""]);
  filas.push(["", "", "", "", "", "", "", "", ""]);
  let filaCuadre = filas.length;
  filas.push(["CUADRE", textoCuadre, "", "", "", "", "", "", ""]);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { filaResumenIni: filaResumenIni, filaCuadre: filaCuadre, numCols: COLS };
}

// =====================================================================================
// HOJA "Resumen capitulos" (solo valores, SIN columna PEM)
// =====================================================================================
// Solo capitulos PADRE. El importe de cada oferta es el de la FILA del capitulo
// (criterio maestro), nunca la suma de partidas.
// Estructura (0-based):
//   Fila 0: nombres de empresa (primera columna de su bloque; B hace el merge)
//   Fila 1: cabecera | Filas 2..: capitulos padre | fila TOTAL PEC | blanco | CUADRE
// Columnas: 0=Cod, 1=Capitulo, 2=Media PEC, luego 3 por oferta (Importe, Δ€, Δ%),
//           y al FINAL una columna "Aviso" con "X" donde las partidas no casan con
//           la fila del capitulo (y entre parentesis, en que ofertas).
function hojaCapitulosA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  capOrden: string[], capCodigo: { [c: string]: string }, capNombre: { [c: string]: string },
  capImporte: { [c: string]: number }[], avisoCap: { [c: string]: string[] },
  textoCuadre: string): LayoutCapitulos {

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
      if (capImporte[i][cap] !== undefined) { sumaCap += capImporte[i][cap]; cuenta++; }
    }
    let media = cuenta > 0 ? sumaCap / cuenta : 0;

    // En la columna Cod se muestra el CODIGO (la clave interna lleva ademas la ocurrencia)
    let fila: (string | number)[] = [capCodigo[cap] || cap, capNombre[cap] || "", r2(media)];
    for (let i = 0; i < N; i++) {
      if (capImporte[i][cap] !== undefined) {
        let v = capImporte[i][cap]; // ya redondeado a 2 decimales al leer
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

  // Fila de CUADRE visible
  let vacia: (string | number)[] = [];
  for (let c = 0; c < COLS; c++) { vacia.push(""); }
  filas.push(vacia);
  let filaCuadre = filas.length;
  let fQ: (string | number)[] = ["CUADRE", textoCuadre];
  for (let c = 2; c < COLS; c++) { fQ.push(""); }
  filas.push(fQ);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numCapitulos: capOrden.length, filaTotal: filaTotal, filaCuadre: filaCuadre, numCols: COLS, colAviso: COL_AVISO };
}

// =====================================================================================
// HOJA "Resumen agrupado" (solo valores)
// =====================================================================================
// Capitulos padre agrupados por oficios (titulos EN MAYUSCULAS). Cada grupo suma los
// importes de FILA de sus capitulos padre. La suma de todos los grupos debe ser
// identica al TOTAL PEC; cualquier padre fuera de la tabla cae en "SIN CLASIFICAR
// (REVISAR)" y en una fila de AVISO.
// Estructura (0-based):
//   Fila 0: nombres de empresa | Fila 1: cabecera | Filas 2..: un grupo por fila |
//   fila TOTAL | (si hay avisos) blanco + filas de AVISO | blanco | fila de CUADRE
// Columnas: 0=Grupo, 1=Capitulos que suma, 2=Media PEC, luego 3 por oferta
function hojaAgrupadoA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  agr: AgrupadoCalc, textoCuadre: string): LayoutAgrupado {

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
  let vacia: (string | number)[] = [];
  for (let c = 0; c < COLS; c++) { vacia.push(""); }
  if (agr.avisos.length > 0) {
    filas.push(vacia);
    filaAvisoIni = filas.length;
    for (let a = 0; a < agr.avisos.length; a++) {
      let fA: (string | number)[] = [agr.avisos[a]];
      for (let c = 1; c < COLS; c++) { fA.push(""); }
      filas.push(fA);
    }
  }

  // Fila de CUADRE visible
  filas.push(vacia);
  let filaCuadre = filas.length;
  let fQ: (string | number)[] = ["CUADRE", textoCuadre];
  for (let c = 2; c < COLS; c++) { fQ.push(""); }
  filas.push(fQ);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numGrupos: G, filaTotal: filaTotal, filaAvisoIni: filaAvisoIni, numAvisos: agr.avisos.length, filaCuadre: filaCuadre, numCols: COLS };
}

// =====================================================================================
// HOJA "Comparativa partidas" (solo valores)
// =====================================================================================
// Una fila por capitulo, subcapitulo y partida, EN ORDEN ORIGINAL. Los codigos
// repetidos dentro de una misma oferta conservan cada uno su fila y su posicion; esas
// filas de partida se marcan con "X" en la primera columna (solo aviso visual).
// Estructura (0-based):
//   Fila 0: titulo + nombres de empresa | Fila 1: cabecera | Filas 2..: datos |
//   fila "TOTAL partidas" | blanco | fila de CUADRE
// Columnas fijas: 0=Aviso ("X" si codigo de partida repetido), 1=Codigo,
//                 2=Tipo ("Capitulo"|"Subcapitulo"|"Partida"), 3=Resumen, 4=Ud,
//                 5=Medicion. Luego 2 por oferta (P. Unit, Importe) y al final:
//   - Mas barata (por precio unitario) y Δ % (max vs min), como antes.
//   - Alcance: "SOLO <oferta>" si la partida solo la cotiza UNA oferta y las demas la
//     dejan vacia (diferencia de alcance). No se quita la fila, solo se marca.
//   - Desv. s/media (COMPACTA): en vez de una columna % por oferta (con 6 ofertas
//     serian 6 columnas mas y la hoja quedaria ilegible), una sola columna de texto
//     que lista SOLO las ofertas cuyo precio unitario se desvia mas del umbral de la
//     media de la partida, con signo y %: "ACME +32% | BETA -28%".
function hojaPartidasA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  orden: string[], info: { [k: string]: CodeInfo },
  precioMap: { [k: string]: number }[], importeMap: { [k: string]: number }[],
  dupPartida: { [cod: string]: boolean }, umbralDesvMedia: number,
  textoCuadre: string): LayoutPartidas {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const FIX = 6;
  const colBarata = FIX + N * 2;
  const colDesv = colBarata + 1;
  const colAlcance = colDesv + 1;
  const colDesvMedia = colAlcance + 1;
  const COLS = colDesvMedia + 1;

  let cab1: (string | number)[] = ["COMPARATIVA POR PARTIDA", "", "", "", "", ""];
  for (let i = 0; i < N; i++) { cab1.push(nombres[i]); cab1.push(""); }
  cab1.push("COMPARACIÓN"); cab1.push(""); cab1.push(""); cab1.push("");
  let cab2: (string | number)[] = ["Aviso", "Código", "Tipo", "Resumen", "Ud", "Medición"];
  for (let i = 0; i < N; i++) { cab2.push("P. Unit"); cab2.push("Importe"); }
  cab2.push("Más barata"); cab2.push("Δ %"); cab2.push("Alcance"); cab2.push("Desv. s/media");

  let filas: (string | number)[][] = [];
  filas.push(cab1);
  filas.push(cab2);

  // Totales de importe por oferta (solo partidas), para la fila TOTAL y el cuadre
  let totPart: number[] = [];
  for (let i = 0; i < N; i++) { totPart.push(0); }

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
    let numCotizan = 0;    // ofertas que traen esta fila (para la columna Alcance)
    let idxUnica = -1;
    for (let i = 0; i < N; i++) {
      let tienePrecio = precioMap[i][clave] !== undefined;
      let tieneImporte = importeMap[i][clave] !== undefined;
      let pu = tienePrecio ? precioMap[i][clave] : 0;
      let im = tieneImporte ? importeMap[i][clave] : 0;
      if (tienePrecio || tieneImporte) { numCotizan++; idxUnica = i; }
      if (esCap) {
        fila.push("");
        fila.push(tieneImporte ? im : "");
        if (tieneImporte && im > 0) { vals.push(im); idxs.push(i); }
      } else {
        fila.push(tienePrecio ? pu : "");
        fila.push(tieneImporte ? im : "");
        if (tieneImporte) { totPart[i] += im; }
        if (tienePrecio && pu > 0) { vals.push(pu); idxs.push(i); }
      }
    }
    let cmp = comparar(vals, idxs, nombres);
    fila.push(cmp.emp);
    fila.push(cmp.pct);

    // Columna Alcance: partida cotizada por UNA sola oferta (diferencia de alcance)
    let alcance = "";
    if (!esCap && N >= 2 && numCotizan === 1) { alcance = "SOLO " + nombres[idxUnica]; }
    fila.push(alcance);

    // Columna compacta Desv. s/media: solo ofertas fuera del umbral (por precio unit)
    let desvTxt = "";
    if (!esCap && vals.length >= 2) {
      let sumaPu = 0;
      for (let v = 0; v < vals.length; v++) { sumaPu += vals[v]; }
      let mediaPu = sumaPu / vals.length;
      if (mediaPu > 0) {
        let tokens: string[] = [];
        for (let v = 0; v < vals.length; v++) {
          let d = (vals[v] - mediaPu) / mediaPu;
          if (Math.abs(d) > umbralDesvMedia) {
            let pct = Math.round(d * 100);
            tokens.push(nombres[idxs[v]] + " " + (pct > 0 ? "+" : "") + pct + "%");
          }
        }
        desvTxt = tokens.join(" | ");
      }
    }
    fila.push(desvTxt);

    filas.push(fila);
  }

  // Fila TOTAL partidas (suma de la columna Importe de cada oferta). Puede diferir del
  // PEC oficial si las filas de capitulo declaran otra cosa: eso lo cuenta el CUADRE.
  let filaTotal = filas.length;
  let fT: (string | number)[] = ["", "", "", "TOTAL partidas", "", ""];
  for (let i = 0; i < N; i++) { fT.push(""); fT.push(r2(totPart[i])); }
  fT.push(""); fT.push(""); fT.push(""); fT.push("");
  filas.push(fT);

  // Fila de CUADRE visible
  let vacia: (string | number)[] = [];
  for (let c = 0; c < COLS; c++) { vacia.push(""); }
  filas.push(vacia);
  let filaCuadre = filas.length;
  let fQ: (string | number)[] = ["CUADRE", textoCuadre];
  for (let c = 2; c < COLS; c++) { fQ.push(""); }
  filas.push(fQ);

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return {
    numFilas: orden.length, numCols: COLS, colMasBarata: colBarata, colDesv: colDesv,
    colAlcance: colAlcance, colDesvMedia: colDesvMedia, filaTotal: filaTotal, filaCuadre: filaCuadre
  };
}

// =====================================================================================
// HOJA "Top desviaciones": donde esta el dinero y el margen de negociacion
// =====================================================================================
// Las TOP_DESVIACIONES partidas donde mas difieren las ofertas entre si, medido en
// EUROS ABSOLUTOS (importe de la oferta mas cara menos el de la mas barata de esa
// partida), ordenadas de mayor a menor. Solo partidas cotizadas por 2 o mas ofertas.
// Estructura (0-based): Fila 0: cabecera | Filas 1..: partidas.
// Columnas: 0=Codigo, 1=Resumen, 2..2+N-1=importe de cada oferta, luego Dif €, Dif %
// (dif/min) y Mas barata (por importe).
function hojaTopA(wb: ExcelScript.Workbook, nombreHoja: string, nombres: string[],
  orden: string[], info: { [k: string]: CodeInfo },
  importeMap: { [k: string]: number }[], topN: number): LayoutTop {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;
  const colDifEur = 2 + N;
  const colDifPct = colDifEur + 1;
  const colBarata = colDifPct + 1;
  const COLS = colBarata + 1;

  // Candidatas: partidas con importe en 2+ ofertas, con su diferencia max-min
  let items: TopItem[] = [];
  for (let r = 0; r < orden.length; r++) {
    let clave = orden[r];
    if (info[clave].tipo !== "Partida") { continue; }
    let mn = 0, mx = 0, cuenta = 0, idxMin = -1;
    for (let i = 0; i < N; i++) {
      let im = importeMap[i][clave];
      if (im === undefined) { continue; }
      if (cuenta === 0 || im < mn) { mn = im; idxMin = i; }
      if (cuenta === 0 || im > mx) { mx = im; }
      cuenta++;
    }
    if (cuenta >= 2) {
      items.push({ clave: clave, dif: r2(mx - mn), min: mn, idxMin: idxMin });
    }
  }

  // Seleccion parcial de las topN mayores diferencias (sin .sort con callback:
  // prohibido en este runtime; con topN pequeño es ademas mas barato que ordenar todo)
  let limite = topN < items.length ? topN : items.length;
  for (let a = 0; a < limite; a++) {
    let m = a;
    for (let b = a + 1; b < items.length; b++) {
      if (items[b].dif > items[m].dif) { m = b; }
    }
    let t = items[a]; items[a] = items[m]; items[m] = t;
  }

  let cab: (string | number)[] = ["Código", "Resumen"];
  for (let i = 0; i < N; i++) { cab.push(nombres[i]); }
  cab.push("Dif €"); cab.push("Dif %"); cab.push("Más barata");

  let filas: (string | number)[][] = [];
  filas.push(cab);

  for (let a = 0; a < limite; a++) {
    let it = items[a];
    let ci = info[it.clave];
    let fila: (string | number)[] = [ci.codigo, ci.resumen];
    for (let i = 0; i < N; i++) {
      let im = importeMap[i][it.clave];
      fila.push(im === undefined ? "" : im);
    }
    fila.push(it.dif);
    fila.push(it.min > 0 ? r4(it.dif / it.min) : 0);
    fila.push(it.idxMin >= 0 ? nombres[it.idxMin] : "");
    filas.push(fila);
  }

  ws.getRangeByIndexes(0, 0, filas.length, COLS).setValues(filas);
  return { numFilas: limite, numCols: COLS, colDifEur: colDifEur, colDifPct: colDifPct, colMasBarata: colBarata };
}

// =====================================================================================
// HOJA OCULTA "_CONTRATO": metadatos para el Script B (VERSION 3)
// =====================================================================================
function escribirContrato(wb: ExcelScript.Workbook, nombreHoja: string,
  hojaPortada: string, hojaResumen: string, hojaCapitulos: string, hojaAgrupado: string,
  hojaPartidas: string, hojaTop: string, nombres: string[], layR: LayoutResumen,
  layC: LayoutCapitulos, layG: LayoutAgrupado, layP: LayoutPartidas, layT: LayoutTop,
  layPor: LayoutPortada) {

  let ws = recrear(wb, nombreHoja);
  const N = nombres.length;

  let pares: (string | number)[][] = [];
  pares.push(["VERSION", 3]);
  pares.push(["NUM_OFERTAS", N]);
  for (let i = 0; i < N; i++) { pares.push(["NOMBRE_" + (i + 1), nombres[i]]); }
  pares.push(["HOJA_PORTADA", hojaPortada]);
  pares.push(["HOJA_RESUMEN", hojaResumen]);
  pares.push(["HOJA_CAPITULOS", hojaCapitulos]);
  pares.push(["HOJA_AGRUPADO", hojaAgrupado]);
  pares.push(["HOJA_PARTIDAS", hojaPartidas]);
  pares.push(["HOJA_TOP", hojaTop]);

  // --- "Portada" (2 columnas etiqueta/valor) ---
  pares.push(["POR_NUM_FILAS", layPor.numFilas]);
  pares.push(["POR_COL_ETIQUETA", 0]);
  pares.push(["POR_COL_VALOR", 1]);
  pares.push(["POR_FILA_CUADRE", layPor.filaCuadre]);

  // --- "Resumen ofertas" (ordenada por PEC ascendente; Ranking en col 0) ---
  pares.push(["RES_HEADER_FILAS", 1]);
  pares.push(["RES_FILA_RESUMEN_INI", layR.filaResumenIni]);
  pares.push(["RES_FILA_AVISO_INI", 0]);   // sin filas de aviso en esta hoja
  pares.push(["RES_NUM_AVISOS", 0]);
  pares.push(["RES_NUM_COLS", layR.numCols]);
  pares.push(["RES_COL_RANKING", 0]);
  pares.push(["RES_COL_EMPRESA", 1]);
  pares.push(["RES_COL_PEM", 2]);
  pares.push(["RES_COL_PEC", 3]);
  pares.push(["RES_COL_DIF_ECO_EUR", 4]);
  pares.push(["RES_COL_DIF_ECO_PCT", 5]);
  pares.push(["RES_COL_DIF_MED_EUR", 6]);
  pares.push(["RES_COL_DIF_MED_PCT", 7]);
  pares.push(["RES_COL_AVISO", 8]);
  pares.push(["RES_FILA_CUADRE", layR.filaCuadre]);
  pares.push(["RES_COL_CUADRE", 1]);

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
  pares.push(["CAP_FILA_CUADRE", layC.filaCuadre]);
  pares.push(["CAP_COL_CUADRE", 1]);

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
  pares.push(["AGR_FILA_CUADRE", layG.filaCuadre]);
  pares.push(["AGR_COL_CUADRE", 1]);

  // --- "Comparativa partidas" (marca "X" al inicio; Alcance y Desv. s/media al final) ---
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
  pares.push(["PART_COL_ALCANCE", layP.colAlcance]);
  pares.push(["PART_COL_DESV_MEDIA", layP.colDesvMedia]);
  pares.push(["PART_FILA_TOTAL", layP.filaTotal]);
  pares.push(["PART_FILA_CUADRE", layP.filaCuadre]);
  pares.push(["PART_COL_CUADRE", 1]);

  // --- "Top desviaciones" ---
  pares.push(["TOP_HEADER_FILAS", 1]);
  pares.push(["TOP_NUM_FILAS", layT.numFilas]);
  pares.push(["TOP_NUM_COLS", layT.numCols]);
  pares.push(["TOP_COL_CODIGO", 0]);
  pares.push(["TOP_COL_RESUMEN", 1]);
  pares.push(["TOP_COL_PRIMERA_OFERTA", 2]);
  pares.push(["TOP_COLS_POR_OFERTA", 1]);
  pares.push(["TOP_COL_DIF_EUR", layT.colDifEur]);
  pares.push(["TOP_COL_DIF_PCT", layT.colDifPct]);
  pares.push(["TOP_COL_MAS_BARATA", layT.colMasBarata]);

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

// Devuelve los indices de vals ordenados de menor a mayor (ordenacion por seleccion
// con bucles clasicos; .sort con comparador esta prohibido en este runtime).
function ordenAscendente(vals: number[]): number[] {
  let idx: number[] = [];
  for (let i = 0; i < vals.length; i++) { idx.push(i); }
  for (let a = 0; a < idx.length - 1; a++) {
    let m = a;
    for (let b = a + 1; b < idx.length; b++) {
      if (vals[idx[b]] < vals[idx[m]]) { m = b; }
    }
    let t = idx[a]; idx[a] = idx[m]; idx[m] = t;
  }
  return idx;
}

// Quita extensiones .xlsx/.xlsm/.xls del nombre de la oferta, por si llegan sin limpiar
function limpiarNombre(n: string): string {
  let s = (n ? n : "").toString().trim();
  return s.replace(/\.(xlsx|xlsm|xls)$/i, "");
}

// Fecha de hoy en formato dd/mm/aaaa (para la Portada)
function fechaHoy(): string {
  let d = new Date();
  let dia = d.getDate();
  let mes = d.getMonth() + 1;
  let txtDia = dia < 10 ? "0" + dia : "" + dia;
  let txtMes = mes < 10 ? "0" + mes : "" + mes;
  return txtDia + "/" + txtMes + "/" + d.getFullYear();
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
interface Oferta {
  nombre: string;
  filas: Fila[];
  numPartidas?: number;  // lo aporta el script lector (nº de partidas reales leidas)
  obra?: string;         // nombre de obra, opcional (para la Portada)
}
interface CodeInfo { codigo: string; tipo: string; resumen: string; unidad: string; medicion: number; }
interface Cmp { emp: string; pct: number; }
interface Grupo { nombre: string; capitulos: string[]; }
interface AgrupadoCalc { grupos: Grupo[]; sumas: number[][]; presencia: boolean[][]; avisos: string[]; }
interface TopItem { clave: string; dif: number; min: number; idxMin: number; }
interface LayoutPortada { numFilas: number; filaCuadre: number; }
interface LayoutResumen { filaResumenIni: number; filaCuadre: number; numCols: number; }
interface LayoutCapitulos { numCapitulos: number; filaTotal: number; filaCuadre: number; numCols: number; colAviso: number; }
interface LayoutAgrupado { numGrupos: number; filaTotal: number; filaAvisoIni: number; numAvisos: number; filaCuadre: number; numCols: number; }
interface LayoutPartidas { numFilas: number; numCols: number; colMasBarata: number; colDesv: number; colAlcance: number; colDesvMedia: number; filaTotal: number; filaCuadre: number; }
interface LayoutTop { numFilas: number; numCols: number; colDifEur: number; colDifPct: number; colMasBarata: number; }
