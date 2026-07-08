// =====================================================================================
// SCRIPT LECTOR — lee un presupuesto de obra y devuelve sus filas clasificadas
// =====================================================================================
// Se ejecuta desde Power Automate sobre el Excel de cada constructora. Devuelve
// { nombre, filas } que el flujo acumula en el JSON `ofertasJson` para el Script A.
//
// REGLA DE NIVEL (corregida): una fila es capitulo PADRE solo si su codigo casa
// EXACTAMENTE con "C" + dos digitos (C01, C02, ... C47). Cualquier otra fila que venga
// como "Capítulo" en la columna Nat pero cuyo codigo NO cumpla ese patron (ej: "1",
// "2", "49", "F2_C03.01") es SUBCAPITULO (nivel:"sub"). Las partidas siguen igual
// (tipo:"partida", nivel:""). El orden original de las filas se respeta.
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
    let datos = rango.getValues();
    const cCod = 0, cNat = 1, cUd = 2, cRes = 3, cCan = 4, cPre = 5, cImp = 6;
    for (let i = 0; i < datos.length; i++) {
      let f = datos[i];
      let nat = (f[cNat] || "").toString().trim().toLowerCase();
      let cod = (f[cCod] || "").toString().trim();
      let res = (f[cRes] || "").toString().trim();

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

// Capitulo PADRE: SOLO "C" + exactamente 2 digitos (C01..C99), sin puntos ni nada mas.
// CORREGIDO: antes tambien aceptaba codigos numericos cortos ("1", "4", "49"...) como
// padre, y eso marcaba mal los subcapitulos que ponen "Capítulo" en la columna Nat.
// Esos codigos ahora son nivel:"sub".
function esCapituloPadre(cod: string): boolean {
  return /^C\d{2}$/.test(cod);
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
