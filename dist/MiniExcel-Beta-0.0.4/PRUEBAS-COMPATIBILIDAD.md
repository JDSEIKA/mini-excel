# Pruebas con archivos reales y LibreOffice

Fecha: 2026-11-11

## Qué se probó

- **28 archivos reales** de `apache/poi/test-data/spreadsheet`: `.xls`, `.xlsx`,
  strict OOXML, libros vacíos, fechas 1900/1904, estilos, comentarios,
  hipervínculos, tablas, fórmulas, textos largos, hojas múltiples, formatos
  numéricos y nombres Unicode.
- **4 archivos malformados**: `.xls` truncado, `.xlsx` truncado, bytes aleatorios
  con extensión `.xls` y `.xlsx`.
- **LibreOffice 26.8.2.2**, instalado con `winget`, para convertir los mismos
  archivos y comparar valores/fórmulas con `openpyxl`.
- **Prueba de rendimiento sintética**: 200.000 celdas numéricas en una hoja.

## Resultado del corpus

```text
Archivos:                 28
Carga correcta:           28 / 28
Guardado correcto:        27 / 28
Guardado bloqueado:        1 / 28  (SimpleStrict.xlsx, OOXML strict)
Errores de carga:          0
Cierres del proceso:       0
Tiempo total de carga:   972 ms
Tiempo total de guardado: 876 ms
Malformados con error
controlado:                4 / 4
```

El guardado conservador de un `.xlsx` strict sigue bloqueado para no mezclar namespaces;
la copia OOXML strict se genera desde Archivo y se probó en Excel.

## Correcciones que salieron de este corpus

| Archivo | Fallo encontrado | Estado |
|---|---|---|
| `StringContinueRecords.xls` | Texto de 24.732 caracteres con `STRING`/`CONTINUE`: aparecía un NUL y el guardado fallaba | Corregido; texto igual a `xlrd` |
| `Formatting.xls` | Formatos de fecha con sufijo `;@` se clasificaban mal | Corregido; fechas y formatos |
| `SimpleWithFormula.xls` | El `STRING` cacheado borraba la fórmula | Corregido; fórmula y valor |
| `shared_formulas.xls` | Fórmulas compartidas simples no se reconstruían | Corregido; 40 fórmulas y valores |
| `StringContinueRecords.xls` | Nombres definidos globales y referencias `ptgName` no se resolvían | Corregido; 29 nombres |
| `StringContinueRecords.xls` | `FIXED` no existía | Corregido |
| `SimpleStrict.xlsx` | Relaciones `r:id` con namespace strict no se leían | Corregido; carga 19 celdas |
| `simple-monthly-budget.xlsx` | Nombres definidos con tablas (`tblExpenses[Amount]`) daban `#VALUE!` | Corregido; totales correctos |
| `NumberFormatTests.xlsx` | `TEXT` devuelve coma decimal (es-CO); LibreOffice devuelve punto | Diferencia regional, no corrupción |
| `FormulaEvalTestData.xls` | Fórmulas complejas con tokens/arrays no soportados | Pendiente; 522 celdas sin equivalente |

## Funciones modernas añadidas

Ahora están en el motor y tienen pruebas propias (`ModernFunctionsProbe.cs`):

- `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`.
- `SUBTOTAL` (1–11 y 101–111, con filas ocultas para las variantes 101–111).
- `TIME`, `TIMEVALUE`.
- `XLOOKUP`, `XMATCH` (exacto y búsqueda aproximada básica).
- `TEXTJOIN`.
- `SWITCH`.
- `DAYS`, `DATEDIF`, `NETWORKDAYS`, `WORKDAY`, `YEARFRAC`.
- `RANK.EQ`, `PERCENTILE.INC`, `QUARTILE.INC`.
- `CEILING.MATH`, `FLOOR.MATH`.
- `AGGREGATE`, `TEXTBEFORE`, `TEXTAFTER`.
- `ISNA`, `ISERR`.
- Alias `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`.

Resultado:

```text
ModernFunctions: 33 correctas, 0 fallidas
```

## Avisos al abrir

`CompatibilityScanner` detecta funciones no soportadas en las fórmulas y
`XlsxIO`/`XlsIO` detectan gráficos, imágenes, tablas, tablas dinámicas, macros y
el modo de solo lectura de `.xls`. Al abrir, la interfaz muestra un aviso con:

- los alcances de esta versión (hoja de cálculo básica y ultraligera);
- las funciones no calculadas;
- los objetos que pueden no mostrarse o perderse;
- la recomendación de guardar una copia.

Prueba dedicada:

```text
Warnings: 4 correctas, 0 fallidas
```

## Impresión

`WorkbookPrintDocument` paginan selección, hoja activa, hojas marcadas o todo el
libro. `PrintOptionsDialog` permite elegir orientación, papel, margen, cuadrícula,
encabezados y ajuste a una página o una página de ancho; la vista previa usa el
mismo documento que la impresora.

La paginación descuenta el espacio de encabezados, admite orientación horizontal
y mantiene juntos los encabezados de columna con la tabla. Las fuentes y los
bordes se reducen junto con la escala de ajuste, y el texto que no cabe se recorta
con puntos suspensivos dentro de su celda en lugar de desbordarse.

La hoja, los márgenes y la tabla se convierten a la misma unidad (puntos) antes
de calcular el ajuste. Así una selección grande —incluso con filas o columnas
vacías fuera de los datos— cabe completa en una página vertical u horizontal,
en vez de quedar recortada por la derecha o por abajo.

La vista previa propia muestra todas las páginas sobre un panel desplazable:
la rueda del mouse mueve la vista verticalmente y el control de zoom permite
ampliar o reducir sin perder la página actual.

```text
Print: 5 correctas, 0 fallidas
```

## Gráficos

Se leen `xl/charts/*.xml` y los anclajes de `xl/drawings/*.xml`; los tipos
barras/columnas, líneas, área, circular y dispersión se dibujan en la
cuadrícula y en la vista previa de impresión. Los gráficos nuevos se escriben
como ChartML con referencias a las celdas, y Excel los abre como gráficos
editables. Los gráficos importados que no se modifican se conservan. En la
cuadrícula se pueden seleccionar, mover y redimensionar con el mouse; el cambio
se marca como pendiente de guardar. También se eliminan con `Supr` o desde el
menú contextual, sin borrar las celdas.

```text
Charts: 7 correctas, 0 fallidas
Features (validación, CRB v6 y strict): 18 correctas, 0 fallidas
GridLogic (freeze/autofiltro/validación): 45 correctas, 0 fallidas
```

Se probaron además siete plantillas `.xlsx` de gráficos descargadas de
[`xlsx-chart`](https://unpkg.com/xlsx-chart@0.3.0/template/), no solo archivos
generados por Mini Excel: columnas, líneas, áreas, circular, dispersión y radar.
La comparación visual con LibreOffice muestra diferencias de estilo y ejes,
pero los archivos se abren y los gráficos se conservan.

## Buscar y reemplazar

`Buscar` ahora incluye **Reemplazar todo**, con opciones de mayúsculas,
coincidencia completa y búsqueda en fórmulas. No altera celdas numéricas.

```text
Replace: 5 correctas, 0 fallidas
```

## Nombres definidos y fórmulas compartidas `.xls`

- Se leen registros `NAME`, incluidos los nombres globales y locales.
- Se resuelven los `ixti` de `EXTERNSHEET` para que `'Param'!$C$3` apunte a la
  hoja correcta.
- Se resuelven tokens `ptgName`/`ptgNameX` dentro de fórmulas y de otros nombres.
- Se reconstruyen fórmulas compartidas simples mediante `SHRFMLA` y `ptgExp`,
  desplazando la fórmula al rango compartido.

Prueba dedicada:

```text
XlsNamesShared: 7 correctas, 0 fallidas
```

`shared_formulas.xls` ahora coincide con LibreOffice en valores y fórmulas.
`SharedFormulaTest.xls` conserva los valores, pero sus fórmulas compartidas
normalizadas (`ptgRefN`) ahora se reconstruyen; los valores coinciden, pero
quedan 12 diferencias de texto de fórmula por paréntesis/referencias absolutas.

## Comparación con LibreOffice

Se convirtieron los 28 archivos a `.xlsx` con LibreOffice y se compararon valores
y fórmulas.

- **23 archivos** coinciden en valores sin diferencias.
- **`NumberFormatTests.xlsx`**: 12 diferencias de `TEXT`; son de configuración
  regional (`12,34` frente a `12.34`), no de lectura.
- **`FormulaEvalTestData.xls`**: 12 diferencias de valor y 522 celdas sin
  equivalente; son fórmulas complejas con tokens/arrays no soportados.
- **`StringContinueRecords.xls`**: 4.824 fórmulas reconstruidas, 55 errores de
  evaluación, 12 diferencias de valor/fecha y 517 celdas sin equivalente. Los
  nombres, `VLOOKUP` y las fórmulas compartidas simples se resuelven; quedan
  fórmulas dinámicas y referencias complejas.
- **`SharedFormulaTest.xls`**: diferencias de valor 0; quedan 12 diferencias de
  texto de fórmula por paréntesis y referencias absolutas.
- **`SimpleStrict.xlsx`**: se abre; el guardado conservador se bloquea, pero se puede
  generar una copia strict desde Archivo.

La diferencia de `TEXT` es esperada: Mini Excel está en español y LibreOffice en
esta prueba usó el punto decimal.

## Rendimiento

| Caso | Carga | Guardado | Observación |
|---|---:|---:|---|
| `StringContinueRecords.xls` (1,18 MB, 16.320 celdas) | 0,6-0,7 s | 0,2-0,4 s | Textos largos, 4.824 fórmulas |
| `FormulaEvalTestData.xls` (4.488 celdas, 1.212 fórmulas) | 80-143 ms | 78-105 ms | Fórmulas complejas |
| `Tables.xlsx` (3.544 celdas, 73 combinaciones) | 34-52 ms | 49-60 ms | Tablas y formato |
| Corpus completo, 28 archivos | 1,0 s | 0,9 s | |

Prueba sintética de 200.000 celdas:

```text
construcción: 246 ms
guardado:     991 ms
carga:       1.368 ms
archivo:     1,30 MB
```

El recálculo sigue siendo completo: un libro con muchas fórmulas puede tardar más
al editar.

## Errores controlados

Los cuatro archivos malformados devolvieron `FileFormatException` con mensaje,
sin cerrar el proceso ni crear un archivo de salida:

```text
random.xls   -> El archivo no es un libro .xls válido.
random.xlsx  -> El archivo .xlsx no es un paquete válido.
trunc.xls    -> El archivo no es un libro .xls válido.
trunc.xlsx   -> El archivo .xlsx no es un paquete válido.
```

## ¿Se puede publicar?

**Todavía no como estable 1.0.** Sí como beta/vista previa con las limitaciones
escritas. La diferencia con la revisión anterior es que los nombres definidos
`.xls` y las fórmulas compartidas simples ya funcionan; el bloqueo restante es
más pequeño pero real.

Antes de una publicación estable faltan:

1. **Fórmulas compartidas complejas `.xls`**: las simples y normalizadas ya se
   reconstruyen; quedan matrices y tokens complejos.
2. **Fórmulas complejas/arrays `.xls`**: `FormulaEvalTestData.xls` deja 522
   celdas sin equivalente.
3. **Guardado strict OOXML**: el guardado conservador se bloquea; la copia strict
   se genera mediante reconstrucción y se validó su apertura en Excel.
4. **Objetos que no son celdas**: imágenes, macros, comentarios y algunos formatos
   condicionales de `.xls` no se conservan al convertir.
5. **Funciones dinámicas**: `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`,
   `LAMBDA` y derrame `#` siguen fuera de alcance.
6. **Máscaras `TEXT` y más pruebas de estrés**: falta un corpus de cientos de MB y
   fórmulas encadenadas.

## Recomendación

- **Publicar ahora:** beta, con `LEEME` que explique que `.xls` es de solo lectura
  y que los objetos avanzados pueden perderse.
- **Publicar como estable:** después de cerrar los puntos 1–3 y ampliar el corpus.

## Reproducir

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

El repositorio conserva los fixtures necesarios en `tests/bin` y `tests/corpus`.
Las pruebas que no forman parte del build se retiraron para mantener el proyecto
organizado.

