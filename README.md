# Flujo comparador de ofertas — Office Scripts

Comparador de presupuestos de obra de varias constructoras, para Excel Online ejecutado
desde Power Automate ("Run script from SharePoint library"). Tres scripts en cadena:

| Script | Fichero | Qué hace | Parámetros |
|---|---|---|---|
| **Lector** | `script-lector.ts` | Lee el Excel de cada constructora y clasifica sus filas (padre solo si `/^C\d{2}$/`) | `nombreOferta` |
| **A — Cálculo** | `script-a-calculo.ts` | Genera las 4 hojas con solo valores, aplica el criterio maestro y deja la hoja oculta `_CONTRATO` | `ofertasJson` (string) |
| **B — Formato** | `script-b-formato.ts` | Aplica solo estética leyendo `_CONTRATO` (semáforos, bordes, merges, paneles). Re-ejecutable sin riesgo | (ninguno) |

## Criterio maestro (Script A)

1. **Capítulo padre** = fila `tipo:"capitulo"` y `nivel:"padre"`. Su importe es **siempre
   el de su propia fila** (valor oficial de la constructora). Ese número manda.
2. Cada partida y subcapítulo pertenece al **último capítulo padre que apareció por
   encima en el orden original** (por posición, no por código). Por eso ya no existe la
   antigua tabla de excepciones por código.
3. **TOTAL PEC** de una oferta = suma de los importes de las filas de capítulos padre.
4. **Comprobación** (solo avisa): por capítulo, la suma de sus partidas se compara con
   el importe de la fila (tolerancia 1 €); si no casa → "X" en la columna Aviso de
   "Resumen capitulos". El importe usado sigue siendo el de la fila.

## Hojas generadas

1. **Resumen ofertas** — PEC por oferta, desviaciones vs más económica y vs media, PEM
   manual, columna Aviso (anomalía PEM/PEC: >20 % por debajo de la media, configurable).
2. **Resumen capitulos** — solo padres, importe de fila por oferta, media, desviaciones,
   TOTAL PEC y columna **Aviso "X"** donde las partidas no casan con la fila del capítulo.
3. **Resumen agrupado** — oficios en MAYÚSCULAS según `tablaGrupos()` (editable):
   ESTRUCTURAS C01–C04 · OBRA SUCIA C05–C08 · CARPINTERÍA Y CERRAJERÍA C09–C11 ·
   PINTURAS C12 · URBANIZACIÓN C13+C14 · VARIOS C15 · INSTALACIONES
   C16+C30+C31+C32+C33+C38+C39+C40+C41+C43+C44 · CONTROL DE CALIDAD C45 ·
   SEGURIDAD Y SALUD C46 · GESTIÓN DE RESIDUOS C47. Los padres fuera de la tabla van a
   **SIN CLASIFICAR (REVISAR)** con fila de AVISO; se comprueba además numéricamente que
   la suma de grupos = TOTAL PEC.
4. **Comparativa partidas** — capítulos, subcapítulos y partidas en orden original.
   Columna de marca al inicio: **"X"** en partidas cuyo código se repite dentro de una
   misma oferta (las repetidas conservan su fila y posición; se alinean entre ofertas
   por número de ocurrencia).

## Contrato entre A y B

Hoja oculta `_CONTRATO` (pares clave/valor, índices 0-based) con número de ofertas,
nombres, nombres de hojas e índices de filas/columnas clave. `VERSION = 2`. El Script B
actual funciona sin cambios: todos los desplazamientos de columnas (marca al inicio de
la comparativa, Aviso al final de capítulos) viajan en el contrato. La lista completa de
claves está en la cabecera de `script-a-calculo.ts`.

## Config editable

- Script A: `UMBRAL_ANOMALIA_PEM` (20 %), `TOLERANCIA_CUADRE_CAPITULO` (1 €),
  `TOLERANCIA_DESCUADRE` (1 €), `tablaGrupos()`.
- Script B: colores, formatos de número, grosor de borde (en `encuadrar()`, no en la
  config: el linter prohíbe guardar enums de la API).

## Encadenado en Power Automate

1. "Run script" → **Lector** sobre cada Excel de oferta; el flujo acumula `ofertasJson`.
2. "Run script from SharePoint library" → **Script A** sobre el libro de salida con `ofertasJson`.
3. Tras su éxito → **Script B** sobre el mismo libro, sin parámetros.

## Restricciones del runtime respetadas

Sin `.find()`, `.reduce()`, `.map()`, `.filter()`, spread ni flechas en callbacks de
array (solo bucles `for`); sin guardar APIs/enums de ExcelScript en variables de datos;
lectura/escritura por rangos en bloque; tipado estricto sin `any`; inmovilización de
paneles por hoja con `unfreeze()` previo y rangos de la propia hoja.
