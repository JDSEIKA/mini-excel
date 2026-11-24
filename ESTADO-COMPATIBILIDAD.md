# Estado de desarrollo y compatibilidad con Excel

Fecha: 2026-11-18  
Versión: Mini Excel Beta 0.0.5  
Artefacto: `dist/MiniExcel.exe`

## Estado de desarrollo

| Indicador | Estado |
|---|---|
| Fase | Beta 0.0.5 |
| Compilación | Correcta, sin errores ni warnings |
| Pruebas automatizadas | 292 correctas, 0 fallidas |
| Tamaño del proyecto | 70 archivos, ~3,9 MB |
| Código fuente | 24 archivos `.cs`, ~24.738 líneas |
| Dependencias de ejecución | .NET Framework 4.8, WinForms, GDI+ |
| Escritura `.xlsx` | Sí |
| Lectura `.xls` | Sí |
| Escritura `.xls` | No; se guarda como `.xlsx` |
| Macros/VBA | Detectadas, nunca ejecutadas |
| Firma Authenticode | Pendiente |

## Porcentaje estimado de compatibilidad

No existe un porcentaje oficial único: Excel tiene miles de funciones, objetos y reglas de presentación. La siguiente estimación es **ponderada por uso típico de oficina**, no una certificación.

| Área | Peso | Cobertura estimada | Aporte |
|---|---:|---:|---:|
| Fórmulas y cálculo básico | 20% | 85% | 17,0 |
| Lectura/escritura XLSX | 15% | 78% | 11,7 |
| Edición y cuadrícula | 10% | 72% | 7,2 |
| Formatos y estilos | 8% | 70% | 5,6 |
| Tablas, filtros y validación | 8% | 65% | 5,2 |
| Impresión y paginación | 5% | 65% | 3,3 |
| Gráficos | 7% | 45% | 3,2 |
| Tablas dinámicas y segmentadores | 6% | 35% | 2,1 |
| Lectura `.xls` | 5% | 55% | 2,8 |
| Escritura `.xls` | 5% | 0% | 0,0 |
| Macros, automatización y vínculos | 4% | 10% | 0,4 |
| Seguridad y recuperación | 3% | 80% | 2,4 |
| Casos límite e interoperabilidad | 4% | 35% | 1,4 |
| **Total ponderado** | **100%** |  | **≈ 62%** |

### Lectura recomendada

- **Uso doméstico/oficina básica:** aproximadamente **88–92%** de los casos habituales.
- **Uso profesional con Excel como referencia:** aproximadamente **60–65%**.
- **Sustituto completo de Excel:** aproximadamente **35–45%** en escenarios avanzados.

El porcentaje global estimado es **≈ 62%**, con un rango razonable de **60–65%**.

## Qué está cubierto

- Apertura y guardado de `.xlsx` preservando partes avanzadas.
- Lectura de `.xls` BIFF8.
- Cuadrícula, edición, selección, copiar/pegar, undo/redo.
- Fórmulas básicas, funciones modernas seleccionadas, arrays dinámicos básicos.
- Formatos de número, estilos, bordes, combinaciones, anchos y altos.
- Tablas, autofiltros, validaciones, formato condicional.
- Gráficos de columnas, barras, líneas, áreas, tortas, dispersión, radar y combinados.
- Torta doble, `bar-of-pie` aproximado, imágenes y formas.
- Tablas dinámicas básicas y segmentadores.
- Impresión, encabezados, pies, títulos y saltos.
- Recuperación cifrada y guardado atómico.

## Qué falta o es aproximado

- Escritura `.xls` BIFF8.
- Macros/VBA y automatización.
- Actualización de vínculos externos.
- Tablas dinámicas avanzadas: agrupaciones, campos calculados, estilos exactos.
- Fórmulas no soportadas y diferencias de redondeo/cultura.
- Todas las máscaras de formato de Excel.
- Tipos avanzados de gráfico con equivalencia visual exacta.
- Impresión y paginación pixel-perfect.
- Firma Authenticode, CI y fuzzing continuo.

## Verificación

```text
Motor:                    91 correctas
Compatibilidad básica:    47 correctas
Funciones modernas:       33 correctas
.xls nombres/fórmulas:     7 correctas
Avisos:                    4 correctas
Impresión:                 5 correctas
Gráficos:                  7 correctas
Reemplazo:                 5 correctas
Features:                 18 correctas
Cuadrícula:               45 correctas
Regresión:                30 correctas
Total:                   292 correctas, 0 fallidas
```

## Conclusión

El proyecto está en estado **beta funcional**, adecuado para hojas de cálculo sencillas y documentos de oficina, con una compatibilidad global estimada de **≈ 62%** frente a Excel y una compatibilidad de **≈ 90%** en operaciones básicas.

No debe presentarse todavía como sustituto universal de Excel.
