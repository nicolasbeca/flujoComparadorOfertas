# Flujo comparador de ofertas — Office Scripts

Comparador de presupuestos de obra de varias constructoras, para Excel Online ejecutado
desde Power Automate ("Run script from SharePoint library"). La lógica está separada en
dos scripts que trabajan en cadena sobre el mismo libro de salida:

| Script | Fichero | Qué hace | Parámetros |
|---|---|---|---|
| **A — Cálculo** | `script-a-calculo.ts` | Genera las 4 hojas con solo valores, valida cuadres y deja la hoja oculta `_CONTRATO` | `ofertasJson` (string) |
| **B — Formato** | `script-b-formato.ts` | Aplica solo estética leyendo `_CONTRATO` (semáforos, bordes, merges, paneles). Re-ejecutable sin riesgo | (ninguno) |

## Hojas generadas

1. **Resumen ofertas** — total PEC por oferta, desviaciones vs más económica y vs media,
   columna PEM manual, columna **Aviso** (anomalía PEM/PEC: >20 % por debajo de la media,
   umbral editable) y filas de **AVISO de descuadre** entre hojas.
2. **Resumen capitulos** — solo capítulos padre (sin columna PEM), importe por oferta,
   media y desviaciones. Fila TOTAL PEC. Semáforo verde→rojo por columna.
3. **Resumen agrupado** — capítulos agrupados por oficios según la tabla editable
   `tablaGrupos()` del Script A (Estructuras, Obra sucia, Carpintería y Cerrajería,
   Pinturas, Urbanización, Varios, Instalaciones C16+C31–C44, Cierre). Los capítulos
   fuera de la tabla van al grupo automático **"Sin clasificar (revisar)"** con AVISO,
   para que el total nunca descuadre en silencio.
4. **Comparativa partidas** — una fila por partida/capítulo en el orden original.
   Columnas: Código, **Tipo** (Capitulo/Subcapitulo/Partida), Resumen, Ud, Medición;
   luego P. Unit e Importe por oferta; al final oferta más barata y desviación %.

## Validación crítica de cuadre (Script A)

Por cada oferta se comparan cuatro totales: T1 = total PEC (partidas), T2 = suma de
capítulos, T3 = suma de grupos, T4 = suma de partidas únicas de la comparativa. Cualquier
diferencia mayor que la tolerancia (1 €, configurable) genera una fila AVISO en
"Resumen ofertas", que el Script B pinta en naranja. Capítulos duplicados en dos grupos o
fuera de la tabla generan AVISO en "Resumen agrupado".

## Contrato entre A y B

El Script A escribe la hoja oculta `_CONTRATO` (pares clave/valor, índices 0-based):
número de ofertas, nombres, nombres de hojas, índices de filas/columnas clave y posición
de las filas de aviso. El Script B se orienta exclusivamente con esa hoja y comprueba
`VERSION = 2`. Si falta algo, se detiene con un error que pide ejecutar A primero.
La lista completa de claves está en la cabecera de `script-a-calculo.ts`. B deja la hoja
oculta al terminar (no la borra, para poder re-ejecutar B sin repetir A).

## Puntos que son criterio del usuario (verificar por proyecto)

- **Tabla de grupos** (`tablaGrupos()`, Script A) y **tabla de excepciones de capítulo**
  (`tablaExcepciones()`, Script A: telecom 1.x–4 → C33, sanitarios 23 → C16).
- **Umbrales**: anomalía PEM (`UMBRAL_ANOMALIA_PEM`, 20 %) y tolerancia de descuadre
  (`TOLERANCIA_DESCUADRE`, 1 €) en el Script A.
- **Colores, formatos, grosor de bordes**: zona de configuración del Script B.

## Encadenado en Power Automate

1. "Run script from SharePoint library" → **Script A (cálculo)** sobre el libro de
   salida, pasando `ofertasJson`.
2. Tras su éxito, "Run script from SharePoint library" → **Script B (formato)** sobre el
   **mismo libro**, sin parámetros.

## Restricciones del runtime respetadas

Sin `.find()`, `.reduce()`, `.map()`, `.filter()`, spread ni flechas en callbacks de
array (solo bucles `for`); lectura/escritura por rangos en bloque; semáforos con
`addConditionalFormat` (color scale) por columna; tipado estricto sin `any`;
inmovilización de paneles por hoja con `unfreeze()` previo y rangos de la propia hoja.
