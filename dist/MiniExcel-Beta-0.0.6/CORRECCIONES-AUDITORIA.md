# Correcciones de la auditoría estricta

Fecha: 2026-11-17
Base: Mini Excel Beta 0.0.6

## Corregido

### Protección de hoja

Se centralizó la comprobación de escritura:

- `ApplyCellText` ya no modifica hojas protegidas.
- Insertar/eliminar filas y columnas.
- Pegar.
- Borrar selección.
- Combinar/separar.
- Ordenar.
- Rellenar.
- Segmentadores.
- Actualización de tablas dinámicas.

Prueba:

```text
Hoja protegida: sí
ApplyCellText("HACKEADO") -> false
Valor final de A1: x
```

### Undo en tablas dinámicas

`RefreshPivot` y el filtrado del segmentador toman una instantánea de undo antes de reescribir la tabla. También se bloquean si la hoja está protegida.

### HSTACK y VSTACK

El fallo era un `foreach (FunctionNode arg in n.Args)` que lanzaba una conversión inválida para literales y rangos. Se cambió a `FormulaNode`.

Prueba:

```text
HSTACK -> matriz 2×2
VSTACK -> matriz 2×2
```

### Formato condicional

Se añadió una caché de estadísticas por rango con versión de datos:

- mínimo/máximo;
- lista ordenada para Top/Bottom;
- invalidación al cambiar celdas o recalcular.

Se evita recorrer el rango completo para cada celda.

### Límites de objetos

Nuevos topes:

```text
Imágenes:    2.000
Gráficos:      500
Formas:      2.000
Segmentadores: 500
Comentarios: 100.000
```

### Tablas

- El rango de tabla se guarda en el XML de tabla.
- Se actualiza `ref` y `autoFilter ref`.
- Se ajustan las columnas de tabla al ampliar a la derecha.
- Se mantiene el estilo después de recargar.

Prueba:

```text
tabla inicial: A1:F61
tras editar:   A1:G62
tras guardar y recargar: A1:G62
```

### Impresión

- `SaveStrict` usa un temporal único.

### Visual

- Botones de filtro de tabla visibles en la fila de encabezado de la hoja 2.
- Torta principal + torta secundaria.
- Etiquetas de la torta secundaria a la derecha.
- Etiquetas dibujadas en dos pasadas para que los sectores no las tapen.
- Colisión vertical de etiquetas corregida.
- `bar-of-pie` dibuja barras de detalle.

## Pendiente de la auditoría

1. `StopIfTrue` no se aplica completamente.
2. El tema de gráficos sigue siendo global.
3. La recuperación no está cifrada.
4. Los tipos avanzados de gráfico siguen siendo aproximaciones.
5. Las macros no se editan ni ejecutan.
6. Los vínculos externos no se actualizan.

## Verificación

```text
Pruebas:     91 correctas
Warnings:     4 correctas
Gráficos:     7 correctas
Features:    18 correctas
GridLogic:   45 correctas
Total:      294 comprobaciones, 0 fallidas
```

## Artefactos

```text
MiniExcel.exe
534.528 bytes
SHA-256: 911F9EA26B78FFEAE1DD0C634F6BB39A2DAF11F98F5F44351166602C2CCA5CC8

MiniExcel-Beta-0.0.6.zip
275843 bytes
SHA-256: 05825617F1F50DF8C197DCB17885B73C46F3C1C2E4AB782331FF03FADFA8A926
```


## Segunda ronda de correcciones de auditoría

- Protección añadida a `ClearFormatting`, `AutoSum`, filtros, `SetAutoFilterSelection`, validaciones, recorte/borrado de imágenes, borrado de gráficos, corte y reemplazo.
- `SaveStrict` con temporal único, reemplazo atómico y respaldo; `ReplaceAtomically` restaura el destino si falla la copia.
- Clon profundo de imágenes, formas, pivotes, cachés y segmentadores; presupuesto de celdas para undo.
- Recuperación cifrada con DPAPI, metadatos cifrados, carga compatible con snapshots antiguos y guardado en segundo plano.
- Caché de expresiones limitada a 2.048 entradas y reinicio al abrir libro.
- Paleta de gráficos por libro/gráfico.
- Diagnóstico visible y `LastError` en gráficos no representables.
- `LastSaveError` y propagación de errores de guardado.
- Persistencia de criterios de filtro en `filterColumn`/`filters` y recarga verificada.
- Pruebas de aislamiento de clones añadidas: 294 comprobaciones.


## Ronda Beta 0.0.6 — auditoría de 0.0.5

- H1: CSV ya no antepone apóstrofo a números negativos; se añadieron pruebas con -5 y -1234,5.
- H2: recuperación por sesión con PID, exclusión de sesiones activas y borrado de la sesión restaurada/descartada.
- H3/H4: cifrado temporal y presupuesto de lectura aislados por hilo; `Load` limpia el presupuesto en `finally`.
- H5: CSV exporta valores numéricos crudos y fechas en ISO 8601 en lugar de texto formateado.
- H6: separador CSV según `TextInfo.ListSeparator`.
- H7: captura de `RegexMatchTimeoutException` en fórmulas compartidas de `.xls`.
- H9: exportación CSV rechaza una carpeta de destino inexistente.
- H10: solo los teléfonos con espacios/guiones se fuerzan a texto; `+5` sigue siendo numérico.
- Pendiente: P1 (límite 512), P2 (DOM XML), P3 (trabajo pesado en UI), P4 (guardado nativo no unificado).
