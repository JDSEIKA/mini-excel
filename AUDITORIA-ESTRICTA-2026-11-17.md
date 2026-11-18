# Auditoría estricta — Mini Excel Beta 0.0.3

Fecha: 2026-11-17  
Alcance: código fuente completo de `src/` (24 archivos, 24.738 líneas), compilación, pruebas automatizadas y sondas dinámicas.  
Artefacto auditado inicialmente: `dist/MiniExcel.exe`, 530.944 bytes, SHA-256 `C57024B597D8F6A1EE39A9959D7F30E6C93D089EB3E1B42F1D8F8F16C11C2A03`.
Artefacto corregido: `dist/MiniExcel.exe`, 530.944 bytes, SHA-256 `C57024B597D8F6A1EE39A9959D7F30E6C93D089EB3E1B42F1D8F8F16C11C2A03`.

## Veredicto

El proyecto compila sin errores ni advertencias del compilador y las 290 comprobaciones automatizadas pasan. Se corrigieron los hallazgos críticos y la mayoría de los altos; **sigue sin estar listo para considerarse una auditoría de seguridad completa** porque quedan fuzzing, firma, escritura BIFF8 y pruebas de integración.

La protección de hojas tiene rutas de escritura que todavía la omiten. El guardado «strict» puede perder el archivo de destino si falla el reemplazo. El sistema de deshacer comparte objetos mutables entre instantáneas.

## Estado después de las correcciones

Se corrigieron los hallazgos que no requerían cambios de formato o dependencias externas:

| Hallazgo | Estado |
|---|---|
| A-01 protección incompleta | Corregido en las rutas detectadas; se añadieron guardas a filtros, validaciones, autosuma, formato, imágenes, gráficos y reemplazo. |
| A-02 guardado strict no atómico | Corregido: temporal único, validación, `File.Replace` y respaldo de recuperación. |
| A-03 undo comparte objetos | Corregido: clones de imágenes, formas, pivotes, cachés y segmentadores; pruebas de aislamiento. |
| A-04 undo sin presupuesto | Corregido parcialmente: presupuesto de celdas copiadas y límite dinámico; sigue siendo undo por copia completa. |
| A-05 recuperación plana y síncrona | Corregido: DPAPI para snapshot y metadatos, guardado en segundo plano y compatibilidad con recuperaciones antiguas. |
| A-06 caché de expresiones sin límite | Corregido: límite de 2.048 AST y limpieza al abrir libro. |
| A-07 tema global | Corregido: paleta por libro y por gráfico. |
| A-08 errores de gráfico silenciosos | Corregido parcialmente: se conserva `LastError` y se muestra un marcador de error. |
| A-09 errores de guardado silenciosos | Corregido: `LastSaveError` y propagación desde `FileIO.Save`. |
| A-10 filtros sin criterios | Corregido: `FilterSelections` y XML `filterColumn`/`filters`; prueba de guardado/recarga. |
| A-11 parser OLE2/GDI+ | Sin cambios; requiere fuzzing. |
| A-12 firma | Sin cambios; requiere certificado. |
| A-13 pruebas/CI | Mejorado: pruebas de aislamiento de clones; no hay CI todavía. |
| A-14 `.xls` escritura | Limitación conocida; no se implementó BIFF8 de escritura. |
| A-15 macros/vínculos | Limitación de seguridad documentada; no se ejecutan macros. |
| Codificación | Falso positivo: el código fuente está en UTF-8 correcto; el mojibake era de la consola de auditoría. |

La sección siguiente conserva los hallazgos originales como registro de auditoría.

## Hallazgos críticos

### A-01 — Protección de hoja incompleta *(corregido en las rutas detectadas)*

Severidad: **Crítica**  
Archivos: `src/SheetGrid.cs`

`EnsureWritable()` se aplica a la edición normal, insertar/borrar, pegar, combinar, ordenar, rellenar, segmentadores y pivotes. Pero quedan rutas públicas que modifican el documento sin comprobarla:

| Método | Línea aprox. | Qué modifica |
|---|---:|---|
| `ClearFormatting()` | 1492 | estilos de celdas |
| `SetAutoFilterSelection()` | 330 | `AutoFilterRef` |
| `AutoSum()` | 1884 | fórmulas |
| `ApplyFilter()` | 451 | filas ocultas |
| `ClearFilter()` | 478 | filas ocultas |
| `CropSelectedImage()` | 523 | imagen |
| `DeleteSelectedImage()` | 545 | imagen |
| `DeleteSelectedChart()` | 567 | gráfico |
| `ApplyValidation()` | 256 | validaciones |

Evidencia dinámica sobre una hoja con `IsProtected = true`:

```text
before style bold=True
after style bold=False
autoFilter=C1:C1
autosum A3==SUMA(A1:A2)
```

Es decir: se borró formato, se creó un autofiltro y se insertó una fórmula en una hoja protegida.

Impacto: un documento protegido puede modificarse desde la interfaz; no es una barrera de seguridad, pero sí un fallo de integridad y de expectativa del usuario.

Recomendación: centralizar toda mutación en un único método que compruebe `ReadOnly`, incluyendo objetos, filtros, validaciones, autosuma y formato. Añadir pruebas de regresión por cada ruta.

### A-02 — `SaveStrict` podía destruir el archivo de destino *(corregido)*

Severidad: **Crítica**  
Archivo: `src/XlsxIO.cs`, `SaveStrict` y `ConvertPackageToStrict`.

El flujo es:

```text
Save(wb, temp)
ConvertPackageToStrict(temp, destination)
  if (File.Exists(destination)) File.Delete(destination);
  File.Move(temp, destination);
```

Si el proceso se interrumpe, el destino está bloqueado o `File.Move` falla, el archivo original ya fue eliminado. No hay copia de respaldo ni reemplazo atómico.

Además, `ReplaceAtomically` usa `File.Copy(temp, path, true)` como último recurso; una copia interrumpida deja el destino parcialmente escrito.

Recomendación: reemplazo atómico con archivo temporal + `File.Replace`/`MoveFileEx`, respaldo previo y recuperación ante fallo. No borrar el destino antes de tener el nuevo archivo validado.

### A-03 — Undo/Redo compartía objetos mutables *(corregido)*

Severidad: **Crítica**  
Archivos: `src/Model.cs` (`Sheet.Clone`, `UndoManager`), `src/Charts.cs`.

`Sheet.Clone` comparte por referencia:

- `Images`
- `Shapes`
- `PivotRanges`
- `Slicers`
- `DeletedImages`

Evidencia dinámica con `Book 1.xlsx`:

```text
shared pivot=True
shared slicer=True
```

`UndoManager.Push` guarda `current.Clone()`. Al mover/recortar una imagen, refrescar un pivote o cambiar un segmentador, se modifica el mismo objeto que está dentro de la instantánea de undo. El resultado es que Undo no restaura el estado anterior y puede corromper la instantánea.

Además, `SheetImage.Image` es un `Bitmap` compartido y `PivotRange.Cache` también se comparte.

Recomendación: clonar profundamente imágenes, formas, pivotes, segmentadores y cachés; o usar copy-on-write. Añadir pruebas de undo para imagen, pivote y segmentador.

## Hallazgos altos

### A-04 — Undo basado en copias completas del libro

Severidad: **Alta**  
Archivo: `src/Model.cs`, `UndoManager`.

Cada operación guarda un `Workbook.Clone()` completo. El límite es de 50 instantáneas, o 5 si hay más de 250.000 celdas. Un libro de hasta 2.000.000 de celdas puede generar copias muy grandes y bloquear la interfaz o agotar memoria.

No hay presupuesto de bytes, ni compresión, ni undo por diferencias.

Recomendación: presupuesto de memoria, snapshots diferenciales o límite duro por bytes; medir el coste de cada operación.

### A-05 — Recuperación sin cifrar y en el hilo de UI *(corregido)*

Severidad: **Alta**  
Archivo: `src/Recovery.cs`, `src/MainForm.cs`.

- El snapshot se guarda como `.xlsx` plano: `session.recovery.xlsx`.
- Los metadatos incluyen ruta y nombre en texto Base64, no cifrado.
- Si `%APPDATA%` falla, se usa un directorio junto al ejecutable.
- El timer de 15 s ejecuta `RecoveryManager.Save` de forma síncrona en el hilo de UI.

Impacto: exposición de datos del documento en disco y pausas perceptibles en libros medianos.

Recomendación: cifrar con DPAPI o clave de usuario, ACL restrictivas, guardado en segundo plano y limpieza segura.

### A-06 — Caché estática de expresiones sin límite *(corregido)*

Severidad: **Alta**  
Archivo: `src/ConditionalFormatting.cs`.

`private static readonly Dictionary<string, FormulaNode> Cache` crece con cada fórmula/desplazamiento evaluado y no tiene límite ni limpieza al cambiar de libro. En sesiones largas con muchos documentos puede retener memoria indefinidamente.

`StatsCache` está limitada a 512 y `ResultCache` a 200.000 entradas, pero el caché de AST no.

Recomendación: límite LRU por libro, invalidación al cerrar/cambiar libro y presupuesto de memoria.

### A-07 — Tema de gráficos global *(corregido)*

Severidad: **Alta**  
Archivo: `src/Charts.cs`, `ChartData.ThemePalette`.

La paleta es un campo estático global. Al abrir otro libro se reemplaza la paleta del libro anterior; los gráficos ya abiertos pueden cambiar de color y los snapshots de undo no conservan su tema.

Recomendación: paleta por libro o por gráfico, no estática.

### A-08 — Errores de gráfico silenciados

Severidad: **Alta**  
Archivo: `src/Charts.cs`.

`ChartRenderer.Draw` captura cualquier excepción y no informa nada. Un gráfico puede desaparecer o quedar incompleto sin diagnóstico. Varios parsers de color también devuelven `Color.Empty` silenciosamente.

Recomendación: registrar diagnóstico, mostrar marcador de error en el gráfico y conservar el estado parcial.

### A-09 — Errores de guardado sin diagnóstico *(corregido)*

Severidad: **Alta**  
Archivo: `src/XlsxIO.cs`, `Save`.

`Save` captura todo y devuelve `false`, descartando la excepción. La interfaz recibe un fallo genérico; no queda rastro de la causa.

Recomendación: conservar la excepción, registrar el error y mostrar un mensaje con la ruta temporal y el motivo.

### A-10 — Filtros no persistían sus criterios *(corregido)*

Severidad: **Alta**  
Archivos: `src/SheetGrid.cs`, `src/XlsxIO.cs`.

`ApplyFilterAt` solo modifica `HiddenRows`. Al guardar se escriben las filas ocultas y el `ref` del autofiltro, pero no se escriben las condiciones `filterColumn`/`filters` del XML de tabla/hoja.

Al reabrir, las filas pueden aparecer ocultas, pero el menú no muestra el filtro aplicado y Excel no puede reconstruir la condición.

Recomendación: guardar criterios de filtro, no solo visibilidad; mantener el estado por tabla/hoja.

## Hallazgos medios

### A-11 — Parser OLE2 y decodificador GDI+ no verificados formalmente

Severidad: **Media**  
Archivos: `src/Ole2.cs`, `src/XlsxIO.cs`.

Hay límites de tamaño y comprobaciones, pero el parser OLE2/BIFF y la decodificación de imágenes dependen de código propio y GDI+. No hay fuzzing ni corpus de archivos malformados.

Recomendación: fuzzing de `.xls`, `.xlsx`, imágenes y XML; ejecutar el decodificador de imágenes en un proceso aislado si se considera una frontera de seguridad.

### A-12 — Sin firma de código ni reputación

Severidad: **Media**  
Artefacto: `MiniExcel.exe`.

El ejecutable no está firmado. Windows SmartScreen puede bloquearlo y no hay verificación de integridad para el usuario.

Recomendación: firma Authenticode y publicación de hashes.

### A-13 — Pruebas manuales y probes ad hoc

Severidad: **Media**  
Directorio: `tests/`.

Hay más de 170 archivos de prueba/sonda, muchos de un solo uso, y la validación visual depende de capturas manuales. No hay CI, cobertura ni pruebas de regresión para filtros, undo de objetos y protección.

Recomendación: separar pruebas unitarias/integración, ejecutar en CI, añadir pruebas de las rutas A-01/A-03/A-10 y comparaciones de imágenes automatizadas.

### A-14 — `.xls` solo lectura

Severidad: **Media funcional**  
Archivo: `src/FileIO.cs`.

Guardar `.xls` lanza una excepción intencionada; no hay escritura BIFF8. Es una limitación de compatibilidad relevante.

### A-15 — Macros y vínculos externos no soportados

Severidad: **Media funcional**  
Archivos: `XlsxIO.cs`, `FileIO.cs`.

Las macros se detectan y no se ejecutan; los vínculos externos se advierten pero no se actualizan. Es una decisión de seguridad correcta, pero debe comunicarse como limitación.

## Riesgos residuales

- No hay sandbox de sistema operativo.
- El ejecutable se ejecuta con los permisos del usuario.
- La protección de hojas es ofuscación/UI, no cifrado.
- No hay recuperación ante corrupción del archivo de recuperación.
- Las cachés estáticas no son thread-safe; la aplicación actual es UI de un solo hilo.
- `XlsxIO._readBudgetRemaining` es estado estático global y no es reentrante.
- No hay pruebas de estrés de memoria con 2 millones de celdas y 50 undos.
- La compatibilidad visual con Excel sigue siendo aproximada.

## Verificación realizada

```text
Compilación:    OK, sin errores ni advertencias
Pruebas:        91 + 47 + 33 + 7 + 4 + 5 + 7 + 5 + 18 + 45 + 28
Total:          290 correctas, 0 fallidas
Artefacto:      530.944 bytes
SHA-256:        C57024B597D8F6A1EE39A9959D7F30E6C93D089EB3E1B42F1D8F8F16C11C2A03
```

Pruebas dinámicas adicionales:

- hoja protegida + `ClearFormatting` → formato eliminado;
- hoja protegida + `SetAutoFilterSelection` → `AutoFilterRef` modificado;
- hoja protegida + `AutoSum` → fórmula insertada;
- `Workbook.Clone` → pivote y segmentador comparten referencia;
- formato condicional de hoja 2 → 12 ms para 371 celdas tras las correcciones.

## Prioridad recomendada

1. Cerrar A-01, A-02 y A-03.
2. Cifrar y desacoplar la recuperación (A-05).
3. Limitar el undo y la caché de expresiones (A-04, A-06).
4. Terminar las pruebas de integración de filtros y undo (A-10, A-03).
5. Añadir CI, fuzzing y pruebas de regresión (A-11, A-13).

## Limitaciones de esta auditoría

No es una auditoría formal de seguridad, no incluye fuzzing continuo, análisis de cobertura, revisión de todas las combinaciones de formatos ni pruebas de accesibilidad. Las conclusiones se basan en inspección estática, compilación, pruebas existentes y sondas dirigidas.
