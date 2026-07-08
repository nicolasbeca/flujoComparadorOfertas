# Flujo comparador de ofertas — Office Scripts

Comparador de presupuestos de obra de varias constructoras, para Excel Online ejecutado
desde Power Automate ("Run script from SharePoint library"). La lógica está separada en
dos scripts que trabajan en cadena sobre el mismo libro de salida:

| Script | Fichero | Qué hace | Parámetros |
|---|---|---|---|
| **A — Cálculo** | `script-a-calculo.ts` | Genera las 3 hojas con solo valores y la hoja oculta `_CONTRATO` | `ofertasJson` (string) |
| **B — Formato** | `script-b-formato.ts` | Aplica solo estética leyendo `_CONTRATO`. No recalcula nada. Re-ejecutable sin riesgo | (ninguno) |

## Hojas generadas

1. **Resumen ofertas** — total PEC por oferta, desviación vs la más económica y vs la
   media (€ y %), columna PEM manual y columna **Aviso** (detección de posible oferta
   en PEM sin convertir a PEC: más de un 20 % por debajo de la media, umbral editable
   en la config del Script A).
2. **Resumen capitulos** — solo capítulos padre, importe PEC por oferta, media y
   desviaciones vs media. Fila final TOTAL PEC.
3. **Comparativa partidas** — una fila por partida/capítulo en el orden original.
   Columnas fijas: Código, Tipo, Nivel, Resumen, Ud, Medición; luego P. Unit e Importe
   por oferta; al final la oferta más barata y la desviación % máx/mín.

## Contrato entre A y B

El Script A escribe una hoja oculta `_CONTRATO` con pares clave/valor (col A = clave,
col B = valor, índices 0-based): número de ofertas, nombres, nombres de las hojas y los
índices de fila/columna clave de cada hoja. El Script B se orienta exclusivamente con
esa hoja: si A cambia el layout, basta con que actualice el contrato y B sigue
funcionando. Si `_CONTRATO` falta o está incompleta, B se detiene con un error que pide
ejecutar A primero. La lista completa de claves está en la cabecera de `script-a-calculo.ts`.

## Puntos que son criterio del usuario (verificar por proyecto)

- **Tabla de excepciones de capítulo** (`tablaExcepciones()` al principio del Script A):
  telecomunicaciones 1.x/2.x/3.x/4 → C33, aparatos sanitarios 23 → C16. Editar ahí.
- **Umbrales**: anomalía PEM (Script A, `UMBRAL_ANOMALIA_PEM`, 20 %) y desviación de
  partida (Script B, `umbralDesviacionPartida`, 100 %).
- **Colores y formatos**: zona de configuración al principio del Script B.

## Encadenado en Power Automate

1. Acción "Run script from SharePoint library" → **Script A (cálculo)** sobre el libro
   de salida, pasando `ofertasJson` con el JSON de ofertas del script lector.
2. A continuación (tras el éxito de la anterior), otra acción "Run script from
   SharePoint library" → **Script B (formato)** sobre el **mismo libro**, sin parámetros.

## Restricciones del runtime respetadas

Sin `.find()`, `.reduce()`, `.map()`, spread ni flechas en callbacks de array (solo
bucles `for`); escritura y lectura por rangos en bloque; tipado estricto sin `any`;
inmovilización de paneles por hoja con `unfreeze()` previo y rangos de la propia hoja.
