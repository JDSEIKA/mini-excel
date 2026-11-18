# Compatibilidad con Excel: funciones básicas

Comparación de Mini Excel con Excel en el subconjunto de funciones que usa una hoja de
cálculo normal. La referencia son las reglas documentadas de Excel y 47 casos ejecutados
en `tests/BasicCompatProbe.cs`.

Leyenda:

- **Alta**: el caso típico coincide con Excel.
- **Parcial**: el caso básico funciona, pero faltan variantes, máscaras o casos límite.
- **No soportada**: Excel la tiene y Mini Excel no.

## Resultado general

| Área | Compatibilidad | Comentario |
|---|---|---|
| Operadores aritméticos y comparaciones | Alta | `+ - * / ^ % &`, paréntesis, comparaciones |
| Suma, promedio, mínimo, máximo | Alta | Rangos, números, texto ignorado en referencias |
| Conteos (`COUNT`, `COUNTA`, `COUNTBLANK`) | Alta | Corregidos errores y texto vacío |
| Condicionales (`IF`, `IFS`, `IFERROR`, `IFNA`) | Alta | Evaluación perezosa de ramas |
| Lógica (`AND`, `OR`, `NOT`) | Alta | Celdas vacías se ignoran |
| Redondeo y matemáticas básicas | Alta | `ROUND`, `INT`, `TRUNC`, `MOD`, etc. |
| Texto básico | Alta | `LEN`, `LEFT`, `MID`, `TRIM`, `PROPER`, etc. |
| Fechas y horas básicas | Alta | Serial Excel, `DATE`, `HOUR`, `WEEKDAY`, etc. |
| Búsquedas básicas | Alta | `VLOOKUP`, `INDEX`, `MATCH` exactos/aproximados |
| Funciones condicionales agregadas | Alta/Parcial | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`; faltan casos límite |
| Formato de texto | Parcial | `TEXT` cubre máscaras simples; no todas las máscaras de Excel |
| Estadística | Alta/Parcial | `STDEV`, `VAR`, alias `.S`/`.P`, rango/percentiles; faltan distribuciones avanzadas |
| Funciones modernas/dinámicas | Parcial | `XLOOKUP`, `SUMIFS`, `TEXTJOIN`, `AGGREGATE`, fechas de red y estadística; faltan arrays dinámicos |

## Funciones verificadas

Las siguientes 47 comprobaciones pasan con los valores esperados de Excel:

| Función / caso | Resultado en Mini Excel |
|---|---:|
| `SUM(1;2;3)` | 6 |
| `AVERAGE(1;2;3)` | 2 |
| `MIN(3;1;2)` | 1 |
| `MAX(3;1;2)` | 3 |
| `COUNT(A1:A7)` con números, texto, error y fórmula `""` | 3 |
| `COUNTA(A1:A7)` | 6 |
| `COUNTBLANK(A1:A7)` | 2 |
| `SUM(A1:A7)` con un error dentro | `#¡DIV/0!` |
| `ROUND(2,5;0)` | 3 |
| `ROUND(-2,5;0)` | -3 |
| `ROUNDUP(2,1;0)` | 3 |
| `ROUNDDOWN(2,9;0)` | 2 |
| `INT(-2,1)` | -3 |
| `TRUNC(-2,9)` | -2 |
| `MOD(-3;2)` | 1 |
| `ABS(-5)` | 5 |
| `SQRT(9)` | 3 |
| `POWER(2;3)` | 8 |
| `IF(1=1;10;20)` | 10 |
| `IFERROR(1/0;99)` | 99 |
| `AND(VERDADERO;1)` | VERDADERO |
| `OR(FALSO;0)` | FALSO |
| `NOT(FALSO)` | VERDADERO |
| `LEN("Hola")` | 4 |
| `LEFT("Hola";2)` | Ho |
| `RIGHT("Hola";2)` | la |
| `MID("Hola";2;2)` | ol |
| `UPPER("hola")` | HOLA |
| `TRIM(" a  b ")` | a b |
| `PROPER("hola mundo")` | Hola Mundo |
| `FIND("l";"Hola")` | 3 |
| `SEARCH("L";"Hola")` | 3 |
| `DATE(2026;1;2)` | 46024 |
| `YEAR`, `MONTH`, `DAY` de esa fecha | 2026, 1, 2 |
| `HOUR(0,5)` | 12 |
| `WEEKDAY(02/01/2026)` | 6 (viernes) |
| `VLOOKUP(2;H1:I4;2;FALSO)` | B |
| `MATCH(3;H1:H4;0)` | 3 |
| `INDEX(I1:I4;2)` | B |
| `COUNTIF(H1:H4;">2")` | 2 |
| `SUMIF(H1:H4;">2")` | 7 |
| `AVERAGEIF(H1:H4;">2")` | 3,5 |
| `SUMPRODUCT(H1:H4;H1:H4)` | 30 |

## Comparación por familia

### 1. Aritmética y operadores

| Excel | Mini Excel | Estado |
|---|---|---|
| `+ - * /` | Sí | Alta |
| `^` | Sí | Alta; precedencia corregida |
| `%` | Sí | Alta |
| `&` | Sí | Alta |
| `= <> < > <= >=` | Sí | Alta |
| Referencias `A1`, `$A$1`, rangos `A1:B2` | Sí | Alta |
| Referencias 3D `Hoja1:Hoja3!A1` | No | Parcial |
| Constantes de matriz `{1;2;3}` | No | No soportada |
| `#` / derrame dinámico | No | No soportada |

### 2. Agregaciones

| Función | Estado | Nota |
|---|---|---|
| `SUM` | Alta | Ignora texto en rangos; propaga errores |
| `AVERAGE` | Alta | `#DIV/0!` si no hay números |
| `MIN`, `MAX` | Alta | Devuelve 0 sin números |
| `COUNT` | Alta | Cuenta números; ignora texto y errores de referencias |
| `COUNTA` | Alta | Cuenta error, texto y fórmula `""` como no vacíos |
| `COUNTBLANK` | Alta | Cuenta vacío y fórmula que devuelve texto vacío |
| `MEDIAN` | Alta | Básico |
| `STDEV`, `STDEVP`, `VAR`, `VARP` | Alta | También alias `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P` |
| `RANK`, `RANK.EQ`, `PERCENTILE.INC`, `QUARTILE.INC` | Parcial | Casos básicos sin matrices |
| `SUMIF`, `COUNTIF`, `AVERAGEIF` | Parcial | Criterios básicos y comodines |
| `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` | Alta | Varios criterios y rangos del mismo tamaño |
| `MAXIFS`, `MINIFS` | Alta | Varios criterios |
| `SUBTOTAL` | Parcial | 1–11 y 101–111; oculta filas en las variantes 101–111 |
| `AGGREGATE` | No | No soportada |

### 3. Condicionales y lógica

| Función | Estado | Nota |
|---|---|---|
| `IF` | Alta | Ramas perezosas |
| `IFS` | Alta | Pares condición/resultado |
| `IFERROR` | Alta | |
| `IFNA` | Alta | |
| `AND`, `OR`, `NOT` | Alta | Vacíos ignorados |
| `SWITCH` | Alta | Pares valor/resultado y valor por defecto |
| `ISNA`, `ISERR` | Alta | `ISREF` no soportada; `ISERROR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISEVEN`, `ISODD` sí existen |

### 4. Matemáticas y redondeo

| Función | Estado | Nota |
|---|---|---|
| `ROUND`, `ROUNDUP`, `ROUNDDOWN` | Alta | Redondeo alejado de cero |
| `FIXED` | Alta | Decimales y separador de miles; usado por libros `.xls` |
| `INT`, `TRUNC` | Alta | Negativos correctos |
| `MOD` | Alta | Signo del divisor |
| `ABS`, `SIGN`, `SQRT`, `POWER` | Alta | |
| `EXP`, `LN`, `LOG`, `LOG10`, `PI` | Alta | |
| `MROUND` | Parcial | Casos de signo |
| `CEILING`, `FLOOR` | Parcial | También `CEILING.MATH`/`FLOOR.MATH` básicos |
| `RAND`, `RANDBETWEEN` | Alta | No deterministas; rango enorme devuelve `#¡NUM!` |

### 5. Texto

| Función | Estado | Nota |
|---|---|---|
| `LEN`, `LEFT`, `RIGHT`, `MID` | Alta | Longitud en unidades UTF-16 |
| `UPPER`, `LOWER`, `PROPER` | Parcial | Depende de la cultura |
| `TRIM` | Alta | Solo espacios ASCII |
| `CONCATENATE`, `CONCAT` | Alta | Básico |
| `SUBSTITUTE`, `REPT` | Alta | |
| `FIND`, `SEARCH` | Alta | |
| `EXACT` | Alta | |
| `CHAR`, `CODE` | Alta | |
| `TEXT` | Parcial | Máscaras simples de número/fecha/moneda; no todo el lenguaje de Excel |
| `VALUE` | Parcial | Depende de la configuración regional |
| `TEXTJOIN` | Alta | Delimitador, ignorar vacíos y rangos |

### 6. Fechas y horas

| Función | Estado | Nota |
|---|---|---|
| `DATE`, `TODAY`, `NOW` | Alta | Serial 1900 |
| `YEAR`, `MONTH`, `DAY` | Alta | |
| `HOUR`, `MINUTE`, `SECOND` | Alta | Acepta solo hora (`0,5`) |
| `WEEKDAY` | Alta | Tipos 1, 2 y 3 |
| `DATEVALUE` | Parcial | Formatos españoles/ISO básicos |
| `EDATE`, `EOMONTH` | Alta | |
| `TIME`, `TIMEVALUE` | Alta | Horas/minutos/segundos y texto de hora |
| `DAYS`, `DATEDIF`, `NETWORKDAYS`, `WORKDAY`, `YEARFRAC` | Alta/Parcial | Casos básicos; `YEARFRAC` simplificado |
| Año bisiesto 1900 | Parcial | No reproduce el "29/02/1900" ficticio de Excel |

### 7. Búsquedas

| Función | Estado | Nota |
|---|---|---|
| `VLOOKUP`, `HLOOKUP` | Alta | Exacta y aproximada básica |
| `INDEX` | Alta | |
| `MATCH` | Alta | Tipos 0, 1 y -1 básicos |
| `LOOKUP` | No | No soportada |
| `XLOOKUP`, `XMATCH` | Parcial | Exacto y aproximado básico; sin comodines ni búsqueda binaria real |
| `OFFSET`, `INDIRECT` | No | No soportadas |
| `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE` | No | No soportadas |

### 8. Información y referencia

| Función | Estado |
|---|---|
| `ROW`, `COLUMN`, `ROWS`, `COLUMNS` | Alta |
| `CHOOSE` | Alta |
| `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR` | Alta |
| `ISEVEN`, `ISODD` | Alta |
| `NA` | Alta |
| `TYPE`, `N`, `T`, `CELL`, `INFO` | No soportadas |

## Diferencias que pueden afectar resultados

1. **Separadores**: las funciones usan `;` como separador principal. Con comas dentro de
   una función y sin `;`, la coma se trata como separador de argumentos; para decimal
   con coma hay que escribir `;` en la fórmula o usar punto decimal.
2. **Cultura**: mayúsculas/minúsculas, `PROPER`, `VALUE`, `TEXT` y comparaciones usan la
   configuración regional.
3. **Fechas**: se usa el serial 1900 de Excel, pero no el día ficticio 29/02/1900.
4. **Formato**: `TEXT` no implementa todas las máscaras de Excel; los formatos de celda
   personalizados también son un subconjunto.
5. **Arrays dinámicos**: no hay derrame, `#`, `FILTER`, `SORT`, `UNIQUE` ni `SEQUENCE`.
6. **Funciones modernas**: `XLOOKUP`, `SUMIFS`, `COUNTIFS`, `SWITCH` y `TEXTJOIN`
   están soportadas en casos básicos; faltan `FILTER`, `SORT`, `UNIQUE`,
   `SEQUENCE`, `LET`, `LAMBDA` y el derrame dinámico. `XLOOKUP` no implementa
   comodines ni búsqueda binaria real.
7. **Límites de seguridad**: fórmulas con más de 512 dependencias o rangos de más de
   2.000.000 celdas devuelven error para no bloquear la aplicación.

## Conclusión

Para **funciones básicas de una hoja de cálculo doméstica o de oficina** —sumar, promediar,
contar, condicionales, redondeo, texto, fechas y búsquedas— la compatibilidad es alta en
los casos típicos: **47/47 casos de referencia pasan**.

No es un reemplazo completo de Excel. Las diferencias principales están en las funciones
dinámicas (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`, `LAMBDA`), las máscaras
completas de `TEXT`, los arrays de `.xls` y algunos casos límite de fechas/cultura.
Para una planilla de uso general, el subconjunto implementado cubre la mayoría de las
operaciones cotidianas; para macros, tablas dinámicas, arrays dinámicos o compatibilidad
fórmula a fórmula con Excel, todavía no.
