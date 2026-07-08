// =====================================================================================
// SCRIPT LECTOR — lee un presupuesto de obra y devuelve sus filas clasificadas
// =====================================================================================
// Se ejecuta desde Power Automate sobre el Excel de cada constructora. Devuelve
// { nombre, filas } que el flujo acumula en el JSON `ofertasJson` para el Script A.
//
// REGLA DE NIVEL (endurecida):
//   - Una fila es capitulo PADRE solo si su codigo es "C" + dos digitos Y el numero
//     esta dentro del rango de capitulos del proyecto (C01..C47, editable en
//     esCapituloPadre). Asi un subcapitulo con codigo enganoso tipo "C49" NO se cuela
//     como padre.
//   - Cualquier otra fila de capitulo ("1", "2", "49", "C49", "F2_C03.01"...) es
//     SUBCAPITULO (nivel:"sub").
//   - Si la columna Nat dice "Subcapítulo" explicitamente, es sub SIEMPRE.
//   - Las partidas siguen igual (tipo:"partida", nivel:"").
// El orden original de las filas se respeta.
//
// LECTURA POR COLUMNAS ABSOLUTAS: se lee siempre A..G aunque la primera columna usada
// de la hoja no sea la A (getUsedRange puede venir desplazado si A esta vacia y eso
// descuadraria los indices de columna).
// =====================================================================================

function main(workbook: ExcelScript.Workbook, nombreOferta?: string): ResultadoOferta {
  let hojas = workbook.getWorksheets();
  let hoja = hojas[0];
  for (let h of hojas) {
    let n = h.getName().toLowerCase();
    if (n.includes("presupuesto") || n.includes("hoja1")) { hoja = h; break; }
  }

  let filas: Fila[] = [];
  let rango = hoja.getUsedRange();
  if (rango) {
    // Leer SIEMPRE desde la columna A: si el rango usado empezara en B (columna A
    // vacia), los indices fijos de columna leerian datos equivocados sin avisar.
    let ultimaFila = rango.getRowIndex() + rango.getRowCount();
    let datos = hoja.getRangeByIndexes(0, 0, ultimaFila, 7).getValues();
    const cCod = 0, cNat = 1, cUd = 2, cRes = 3, cCan = 4, cPre = 5, cImp = 6;

    for (let i = 0; i < datos.length; i++) {
      let f = datos[i];
      let nat = (f[cNat] || "").toString().trim().toLowerCase();
      let cod = (f[cCod] || "").toString().trim();
      let res = (f[cRes] || "").toString().trim();

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
      filas.push({
        tipo: "partida", nivel: "",
        capitulo: "", codigo: cod, resumen: res,
        unidad: (f[cUd] || "").toString().trim(),
        medicion: aNumero(f[cCan]), precio: aNumero(f[cPre]), importe: aNumero(f[cImp])
      });
    }
  }

  let nombre = (nombreOferta && nombreOferta.trim().length > 0) ? nombreOferta.trim() : "Oferta";
  return { nombre: nombre, filas: filas };
}

// =====================================================================================
// REGLA DE CAPÍTULO PADRE — CONFIG EDITABLE POR PROYECTO
// =====================================================================================
// Padre = "C" + exactamente 2 digitos (mayuscula o minuscula), sin puntos ni nada mas,
// Y ADEMAS el numero debe estar entre PADRE_MIN y PADRE_MAX.
// OJO: el filtro por rango es lo que evita el fallo recurrente: un subcapitulo con
// codigo "C49" (o C48, C50...) casa con el patron C+2digitos pero NO es un capitulo
// padre del proyecto (los padres reales van de C01 a C47). Si en otro proyecto los
// padres llegan mas alla, sube PADRE_MAX aqui.
function esCapituloPadre(cod: string): boolean {
  const PADRE_MIN = 1;
  const PADRE_MAX = 47; // ultimo capitulo padre real del proyecto (C47 = Gestion de residuos)
  let m = cod.trim().toUpperCase().match(/^C(\d{2})$/);
  if (!m) { return false; }
  let n = parseInt(m[1], 10);
  return n >= PADRE_MIN && n <= PADRE_MAX;
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
interface ResultadoOferta { nombre: string; filas: Fila[]; }
