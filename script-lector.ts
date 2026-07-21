// =====================================================================================
// SCRIPT LECTOR — lee un presupuesto de obra y devuelve sus filas clasificadas
// =====================================================================================
// Se ejecuta desde Power Automate sobre el Excel de cada constructora. Devuelve
// { nombre, filas, numPartidas, totalOrigen, totalEncontrado } que el flujo acumula en
// el JSON `ofertasJson` para el Script A.
//
// REGLA DE NIVEL:
//   - Una fila es capitulo PADRE si su codigo es "C" + exactamente dos digitos
//     (C01..C99, mayuscula o minuscula). SIN tope de proyecto: el script vale para
//     cualquier oferta; que no falte ni sobre un capitulo lo controla el CUADRE del
//     Script A contra el total de origen (totalOrigen), no una lista de codigos aqui.
//   - Cualquier otra fila de capitulo ("1", "2", "49", "F2_C03.01"...) es
//     SUBCAPITULO (nivel:"sub").
//   - Si la columna Nat dice "Subcapítulo" explicitamente, es sub SIEMPRE.
//   - Las partidas siguen igual (tipo:"partida", nivel:"").
// El orden original de las filas se respeta.
//
// TOTAL GENERAL DE ORIGEN: se busca al pie del Excel (recorriendo desde el final) una
// fila cuyo codigo o resumen normalizado EMPIECE por "total", descartando los
// subtotales de capitulo ("total c43", "totalc44"...). Su importe se devuelve en
// `totalOrigen` con `totalEncontrado` = true; si no aparece, 0 y false (el Script A
// avisara de que no se pudo validar). Esa fila NO se añade a `filas`.
//
// CABECERA Y COLUMNAS AUTODETECTADAS:
//   - La fila de cabecera se busca en las primeras 15 filas (una celda con
//     "Código"/"Codigo" y otra con "Nat" o "Pres"). Si no aparece, se asume la
//     fila 3 como antes y el script sigue funcionando.
//   - Las columnas se localizan por su texto de cabecera (Código, Nat, Ud, Resumen,
//     CanPres, Pres, ImpPres). Toda columna que no se encuentre por nombre cae al
//     orden fijo clasico A..G, de modo que las ofertas con el formato actual se
//     leen exactamente igual que antes.
//
// PARTIDAS VACIAS: una partida con medicion 0 Y precio 0 Y importe 0 (vacia en el
// origen de esta oferta) NO se incluye en la salida. Los capitulos y subcapitulos
// se mantienen siempre aunque esten a 0.
//
// LECTURA POR COLUMNAS ABSOLUTAS: se lee siempre desde la columna A aunque la primera
// columna usada de la hoja no sea la A (getUsedRange puede venir desplazado si A esta
// vacia y eso descuadraria los indices de columna).
// =====================================================================================

function main(workbook: ExcelScript.Workbook, nombreOferta?: string): ResultadoOferta {
  let hojas = workbook.getWorksheets();
  let hoja = hojas[0];
  for (let iH = 0; iH < hojas.length; iH++) {
    let n = hojas[iH].getName().toLowerCase();
    if (n.includes("presupuesto") || n.includes("hoja1")) { hoja = hojas[iH]; break; }
  }

  let filas: Fila[] = [];
  let numPartidas = 0;
  let totalOrigen = 0;
  let totalEncontrado = false;
  let rango = hoja.getUsedRange();
  if (rango) {
    // Leer SIEMPRE desde la columna A: si el rango usado empezara en B (columna A
    // vacia), los indices de columna leerian datos equivocados sin avisar. Se leen
    // como minimo 7 columnas (A..G) y, si la hoja usa mas, todas las usadas, para
    // que la deteccion por nombre pueda encontrar columnas desplazadas.
    let ultimaFila = rango.getRowIndex() + rango.getRowCount();
    let ultimaCol = rango.getColumnIndex() + rango.getColumnCount();
    let numCols = ultimaCol > 7 ? ultimaCol : 7;
    let datos = hoja.getRangeByIndexes(0, 0, ultimaFila, numCols).getValues();

    // --- Deteccion de la fila de cabecera (fallback: fila 3, indice 2) ---
    let filaCabecera = buscarFilaCabecera(datos);
    let primeraFilaDatos = filaCabecera + 1;

    // --- Deteccion de columnas por nombre (fallback por columna: orden fijo A..G) ---
    let cols = buscarColumnas(datos, filaCabecera);
    let cCod = cols.codigo, cNat = cols.nat, cUd = cols.unidad, cRes = cols.resumen;
    let cCan = cols.medicion, cPre = cols.precio, cImp = cols.importe;

    for (let i = primeraFilaDatos; i < datos.length; i++) {
      let f = datos[i];
      let nat = (f[cNat] || "").toString().trim().toLowerCase();
      let cod = (f[cCod] || "").toString().trim();
      let res = limpiaTexto((f[cRes] || "").toString());

      // Subcapitulo explicito en la columna Nat: sub SIEMPRE, sin mirar el codigo.
      // (Antes estas filas se perdian: solo se aceptaba "Capítulo" o "Partida".)
      if (nat === "subcapítulo" || nat === "subcapitulo") {
        filas.push({
          tipo: "capitulo", nivel: "sub",
          capitulo: cod, codigo: cod, resumen: res, unidad: "",
          medicion: aNumero(f[cCan]), precio: aNumero(f[cPre]), importe: aNumero(f[cImp])
        });
        continue;
      }

      if (nat === "capítulo" || nat === "capitulo") {
        filas.push({
          tipo: "capitulo",
          nivel: esCapituloPadre(cod) ? "padre" : "sub",
          capitulo: cod, codigo: cod, resumen: res, unidad: "",
          medicion: aNumero(f[cCan]), precio: aNumero(f[cPre]), importe: aNumero(f[cImp])
        });
        continue;
      }

      if (nat !== "partida") { continue; }

      let medicion = aNumero(f[cCan]);
      let precio = aNumero(f[cPre]);
      let importe = aNumero(f[cImp]);

      // Partida totalmente vacia en esta oferta: no se incluye en la salida.
      if (medicion === 0 && precio === 0 && importe === 0) { continue; }

      filas.push({
        tipo: "partida", nivel: "",
        capitulo: "", codigo: cod, resumen: res,
        unidad: (f[cUd] || "").toString().trim(),
        medicion: medicion, precio: precio, importe: importe
      });
      numPartidas++;
    }

    // --- TOTAL GENERAL de origen: se busca DESDE EL FINAL hacia arriba (esta al pie)
    //     y se acepta la PRIMERA coincidencia valida encontrada desde abajo. Una fila
    //     es total si su codigo o su resumen normalizado EMPIEZA por "total", salvo
    //     los subtotales de capitulo ("total c43", "totalc44"...): esos se descartan
    //     y se sigue subiendo. OJO: "total 0" SI es total general (el 0 es el marcador
    //     de obra sin nombre, no un capitulo); solo se descarta el patron c+digitos.
    //     El importe se lee de la MISMA columna de importes que usan las partidas. ---
    for (let i = datos.length - 1; i >= primeraFilaDatos; i--) {
      let f = datos[i];
      let txtCod = normaliza((f[cCod] || "").toString());
      let txtRes = normaliza((f[cRes] || "").toString());
      let candidato = "";
      if (txtCod.indexOf("total") === 0) { candidato = txtCod; }
      else if (txtRes.indexOf("total") === 0) { candidato = txtRes; }
      if (candidato === "") { continue; }

      // Subtotal de capitulo: tras "total" (sin espacios) viene "c" + digitos.
      let sinEspacios = candidato.replace(/\s+/g, "");
      if (/^totalc\d+/.test(sinEspacios)) { continue; }

      totalOrigen = aNumero(f[cImp]);
      totalEncontrado = true;
      break;
    }
  }

  let nombre = (nombreOferta && nombreOferta.trim().length > 0) ? nombreOferta.trim() : "Oferta";
  // numPartidas expone cuantas partidas reales trae esta oferta (sin contar las
  // vacias descartadas), para que el script comparador pueda detectar ofertas
  // anomalas con muchas menos partidas que las demas. totalOrigen/totalEncontrado
  // exponen el TOTAL GENERAL declarado al pie del Excel de origen, para que el
  // comparador pueda cuadrar el PEC calculado contra el total oficial.
  return { nombre: nombre, filas: filas, numPartidas: numPartidas, totalOrigen: totalOrigen, totalEncontrado: totalEncontrado };
}

// =====================================================================================
// DETECCION DE LA FILA DE CABECERA
// =====================================================================================
// Busca en las primeras MAX_FILAS_CABECERA filas una que contenga una celda con
// "Código"/"Codigo" y otra celda con "Nat" o "Pres". Devuelve el indice (base 0) de
// esa fila. Si no la encuentra, devuelve 2 (fila 3, el formato clasico) para que las
// ofertas actuales sigan funcionando sin cambios.
function buscarFilaCabecera(datos: (string | number | boolean)[][]): number {
  const MAX_FILAS_CABECERA = 15;
  const FILA_CABECERA_FALLBACK = 2; // fila 3 en Excel (indice base 0)

  let tope = datos.length < MAX_FILAS_CABECERA ? datos.length : MAX_FILAS_CABECERA;
  for (let i = 0; i < tope; i++) {
    let tieneCodigo = false;
    let tieneNatOPres = false;
    for (let j = 0; j < datos[i].length; j++) {
      let celda = normaliza((datos[i][j] || "").toString());
      if (celda === "") { continue; }
      if (celda.indexOf("codigo") >= 0) { tieneCodigo = true; }
      if (celda === "nat" || celda.indexOf("pres") >= 0) { tieneNatOPres = true; }
    }
    if (tieneCodigo && tieneNatOPres) { return i; }
  }
  return FILA_CABECERA_FALLBACK;
}

// =====================================================================================
// DETECCION DE COLUMNAS POR NOMBRE DE CABECERA
// =====================================================================================
// Recorre la fila de cabecera y asigna cada columna por su texto. El orden de las
// comprobaciones importa: "CanPres" e "ImpPres" contienen "pres", asi que se evaluan
// antes que la columna de precio ("Pres"). Cada columna se asigna solo una vez (gana
// la primera coincidencia de izquierda a derecha). Toda columna que no aparezca por
// nombre cae a su posicion fija clasica (A=Codigo, B=Nat, C=Ud, D=Resumen, E=CanPres,
// F=Pres, G=ImpPres), garantizando que las ofertas con el formato actual se leen
// exactamente igual que hasta ahora.
function buscarColumnas(datos: (string | number | boolean)[][], filaCabecera: number): IndicesColumnas {
  // Valores por defecto: el orden fijo clasico A..G.
  let cols: IndicesColumnas = { codigo: 0, nat: 1, unidad: 2, resumen: 3, medicion: 4, precio: 5, importe: 6 };
  if (filaCabecera < 0 || filaCabecera >= datos.length) { return cols; }

  // -1 = todavia no encontrada por nombre.
  let codigo = -1, nat = -1, unidad = -1, resumen = -1, medicion = -1, precio = -1, importe = -1;

  let cab = datos[filaCabecera];
  for (let j = 0; j < cab.length; j++) {
    let celda = normaliza((cab[j] || "").toString());
    if (celda === "") { continue; }

    if (codigo < 0 && celda.indexOf("codigo") >= 0) { codigo = j; continue; }
    // Medicion antes que precio: "CanPres" contiene "pres".
    if (medicion < 0 && (celda.indexOf("canpres") >= 0 || celda.indexOf("medicion") >= 0 || celda.indexOf("cantidad") >= 0)) { medicion = j; continue; }
    // Importe antes que precio: "ImpPres" contiene "pres".
    if (importe < 0 && (celda.indexOf("imppres") >= 0 || celda.indexOf("importe") >= 0)) { importe = j; continue; }
    if (precio < 0 && (celda === "pres" || celda.indexOf("precio") >= 0)) { precio = j; continue; }
    if (nat < 0 && (celda === "nat" || celda.indexOf("naturaleza") >= 0)) { nat = j; continue; }
    if (unidad < 0 && (celda === "ud" || celda === "ud." || celda.indexOf("unidad") >= 0)) { unidad = j; continue; }
    if (resumen < 0 && (celda.indexOf("resumen") >= 0 || celda.indexOf("descripcion") >= 0)) { resumen = j; continue; }
  }

  // Solo se sustituye el indice fijo si la columna se encontro por nombre.
  if (codigo >= 0) { cols.codigo = codigo; }
  if (nat >= 0) { cols.nat = nat; }
  if (unidad >= 0) { cols.unidad = unidad; }
  if (resumen >= 0) { cols.resumen = resumen; }
  if (medicion >= 0) { cols.medicion = medicion; }
  if (precio >= 0) { cols.precio = precio; }
  if (importe >= 0) { cols.importe = importe; }
  return cols;
}

// =====================================================================================
// REGLA DE CAPÍTULO PADRE
// =====================================================================================
// Padre = "C" + exactamente 2 digitos (C01..C99, mayuscula o minuscula), sin puntos ni
// nada mas. SIN tope de proyecto: el script vale para cualquier oferta. Que no falte
// ni sobre un capitulo se controla por CUADRE en el Script A (contra totalOrigen), no
// escondiendo codigos aqui.
function esCapituloPadre(cod: string): boolean {
  return /^C\d{2}$/.test(cod.trim().toUpperCase());
}

// Normaliza un texto de cabecera para compararlo: minusculas y sin acentos.
function normaliza(texto: string): string {
  let t = texto.trim().toLowerCase();
  t = t.replace(/á/g, "a").replace(/é/g, "e").replace(/í/g, "i").replace(/ó/g, "o").replace(/ú/g, "u");
  return t;
}

// Deja el resumen en una sola linea limpia: sin saltos de linea, tabuladores ni
// espacios dobles.
function limpiaTexto(texto: string): string {
  return texto.replace(/[\r\n\t]+/g, " ").replace(/ {2,}/g, " ").trim();
}

function aNumero(v: (string | number | boolean)): number {
  if (typeof v === "number") { return v; }
  if (v === null || v === undefined || v === "") { return 0; }
  let n = parseFloat(v.toString().replace(",", "."));
  return isNaN(n) ? 0 : n;
}

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
interface IndicesColumnas {
  codigo: number;
  nat: number;
  unidad: number;
  resumen: number;
  medicion: number;
  precio: number;
  importe: number;
}
interface ResultadoOferta {
  nombre: string;
  filas: Fila[];
  numPartidas: number;
  totalOrigen: number;       // TOTAL GENERAL declarado al pie del Excel de origen
  totalEncontrado: boolean;  // false si no se localizo ninguna fila de total valida
}
