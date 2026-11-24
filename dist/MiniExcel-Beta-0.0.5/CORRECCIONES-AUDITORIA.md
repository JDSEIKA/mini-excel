# Correcciones de la auditoría estricta

Fecha: 2026-11-17
Base: Mini Excel Beta 0.0.5

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
Total:      292 comprobaciones, 0 fallidas
```

## Artefactos

```text
MiniExcel.exe
531.456 bytes
SHA-256: BD95C43557F76FC204FEC8031D0A828A5C94F33998A37FE19669E7CB240ABD97

MiniExcel-Beta-0.0.5.zip
273588 bytes
SHA-256: 7CC7B09417A65D52E78E89F2FB112EEECA1EA4354394FAB46D0A54229F7EF95D
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
- Pruebas de aislamiento de clones añadidas: 292 comprobaciones.
