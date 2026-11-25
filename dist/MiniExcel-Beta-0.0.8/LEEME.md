# Mini Excel Beta 0.0.8

Hoja de cÃ¡lculo nativa para Windows, parte de la idea **Mini Office**: rÃ¡pida, pequeÃ±a y capaz de
abrir y guardar los `.xlsx` que envÃ­an otras suites.

> **Estado: vista previa (beta), no versiÃ³n estable 1.0.** Lee `PRUEBAS-COMPATIBILIDAD.md`.
> La lectura de `.xls` es de solo lectura y puede perder fÃ³rmulas compartidas, nombres
> definidos, imÃ¡genes, macros o comentarios. Los `.xlsx` strict se abren y se
> pueden exportar como copia OOXML strict desde el menÃº Archivo.

- **Ejecutable de 462 KB**, sin instalador, sin Electron y sin dependencias de terceros.
- **Avisos al abrir** documentos con funciones no soportadas, grÃ¡ficos, imÃ¡genes,
  tablas dinÃ¡micas o macros, explicando los alcances de esta versiÃ³n bÃ¡sica.
- **ImpresiÃ³n con vista previa** (`Ctrl+P`): selecciÃ³n, hoja activa, varias hojas o todo
  el libro; orientaciÃ³n, papel, mÃ¡rgenes, cuadrÃ­cula, encabezados y ajuste a pÃ¡gina.
  La vista previa se desplaza con la rueda del mouse.
- **GrÃ¡ficos**: columnas, barras, lÃ­neas, Ã¡rea, circular y dispersiÃ³n; lectura y
  conservaciÃ³n de grÃ¡ficos de Excel; escritura de ChartML editable en `.xlsx`; tambiÃ©n
  se imprimen y se guardan en `.crb`. Se pueden seleccionar, mover, redimensionar
  y eliminar con `Supr` o el menÃº contextual.
- **Funciones modernas y nombres `.xls`**: `SUMIFS`, `COUNTIFS`, `XLOOKUP`,
  `TEXTJOIN`, `SUBTOTAL`, `TIME`, `SWITCH`; lectura de nombres definidos y fÃ³rmulas
  compartidas simples de libros Excel 97-2003.
- **Arranque en unos 0,25â€“0,37 s** en Windows 10 x64 con .NET Framework 4.8.
- **`.xlsx` como formato principal**, con conservaciÃ³n de imÃ¡genes, grÃ¡ficos, tablas, formato
  condicional, validaciÃ³n, comentarios, combinaciones, nombres definidos e hipervÃ­nculos.
- **Abre `.xls` de Excel 97-2003** (BIFF8), incluidas varias hojas, fÃ³rmulas comunes, estilos,
  combinaciones, anchos, altos, filas/columnas ocultas y paneles. Al guardar se convierte a
  `.xlsx`; la escritura directa `.xls` todavÃ­a no estÃ¡ disponible.
- **Formatos contables de Excel**: el botÃ³n `$` coloca el sÃ­mbolo a la izquierda, el importe a la
  derecha, negativos entre parÃ©ntesis y decimales exactos. TambiÃ©n lee `â‚¬`, `Â£`, `Â¥` y `S/`.
- Aumentar o quitar decimales funciona sobre formatos personalizados y contables creados por
  Excel, sin errores y conservando el formato al guardar.
- **Autosuma** respeta los subtotales anteriores y, si hay una selecciÃ³n, suma exactamente ese
  rango sin arrastrar operaciones de otras zonas.
- Interfaz mÃ¡s cercana a Excel: encabezados oscuros, botones `âœ•`/`âœ“` en la barra de fÃ³rmulas y
  fÃ³rmulas visibles en espaÃ±ol (`=SUMA`, `=SI`).
- TambiÃ©n guarda el formato nativo `.crb` â€”con combinaciones, altos, ocultos, paneles,
  hipervÃ­nculos, `Date1904` y nombres definidosâ€” y exporta `.csv`.

## Empezar

1. Ejecuta `MiniExcel.exe`; no requiere instalaciÃ³n.
2. Escribe en la celda y pulsa `Enter`, o usa la barra de fÃ³rmulas.
3. Prueba `=SUMA(A1:A10)`, `=PROMEDIO(B2:B8)` o `=SI(A1>0;"SÃ­";"No")`.
4. Abre `Ejemplo.xlsx` para ver formato, tabla, imagen y grÃ¡fico creados por otra suite, o
   `Ejemplo.xls` para probar un libro de Excel 97-2003.

Desde la lÃ­nea de comandos:

```bat
MiniExcel.exe "C:\ruta\libro.xlsx"
```

## Compatibilidad XLSX

Al guardar un `.xlsx` abierto, Mini Excel copia el paquete original y parchea solo las partes que
cambian. AsÃ­ conserva lo que no edita:

- Hojas, celdas, fÃ³rmulas, estilos, anchos, altos de fila y filas/columnas ocultas.
- Celdas combinadas, nombres definidos, hipervÃ­nculos y fechas del sistema 1904.
- ImÃ¡genes, grÃ¡ficos, dibujos, tablas, autofiltros, formato condicional, validaciÃ³n de datos y
  comentarios.
- Nombres definidos y fÃ³rmulas de tabla estructurada como `=SUMA(TablaVentas[Importe])`.
- Colores de tema, colores indexados y formatos personalizados como `#,##0.00 "kg"`.

Si el archivo original cambiÃ³ fuera de la aplicaciÃ³n entre abrir y guardar, se regenera el paquete
y las partes avanzadas pueden perderse.

## Comandos principales

| Comando | CÃ³mo |
|---|---|
| Editar | Doble clic, `F2` o escribir |
| Confirmar / cancelar | `Enter` / `Esc` |
| Copiar / cortar / pegar | `Ctrl+C`, `Ctrl+X`, `Ctrl+V` |
| Deshacer / rehacer | `Ctrl+Z`, `Ctrl+Y` |
| Ir a celda | `Ctrl+G` |
| Buscar | `Ctrl+F` |
| Formato de celdas | `Ctrl+1` |
| Negrita / cursiva / subrayado | `Ctrl+B`, `Ctrl+I`, `Ctrl+U` |
| Autosuma | `Alt+=` |
| Combinar / separar celdas | Grupo **AlineaciÃ³n** o menÃº contextual |
| Insertar / eliminar filas o columnas | Grupo **Celdas** o menÃº contextual |
| Recalcular | `F9` |

## LÃ­mites

- Muestra, imprime, mueve, recorta y elimina imÃ¡genes; los cambios se guardan.
- Los grÃ¡ficos bÃ¡sicos se muestran, imprimen y editan en posiciÃ³n/tamaÃ±o.
- No edita tablas, formato condicional, validaciÃ³n ni comentarios desde la interfaz; los conserva.
- Aplica visualmente los paneles inmovilizados.
- No admite macros/VBA, complementos, tablas dinÃ¡micas, colaboraciÃ³n ni archivos con contraseÃ±a.
- Las fÃ³rmulas que el motor no entiende se conservan para Excel; en Mini Excel pueden mostrar error.
- Al insertar o eliminar filas/columnas no se reajustan los rangos internos de tablas, formato
  condicional, validaciÃ³n, comentarios ni nombres definidos.

LÃ­mites de hoja: 1.048.576 filas Ã— 16.384 columnas, 32.767 caracteres por celda.

## Licencia

MIT. Consulta `LICENSE`.






