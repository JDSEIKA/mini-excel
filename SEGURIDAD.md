# Seguridad al abrir archivos

Fecha: 2026-11-17
Alcance: lector `.xlsx`, `.xls`, `.crb` y CSV de Mini Excel Beta 0.0.8.

## Modelo de amenaza

Mini Excel trata cada archivo como **no confiable**. Los riesgos considerados son:

- ZIP bombs y paquetes con partes descomprimidas enormes;
- XML con DTD, entidades o documentos gigantes;
- rutas `../` dentro del ZIP;
- hojas, celdas, cadenas o gráficos en cantidades abusivas;
- imágenes con dimensiones descompuestas enormes;
- macros/VBA y objetos incrustados;
- enlaces externos y protocolos peligrosos;
- fórmulas maliciosas o excesivamente largas;
- contenedores OLE2/BIFF truncados o con cadenas circulares.

## Contramedidas implementadas

### Paquete ZIP

| Control | Valor |
|---|---:|
| Archivo comprimido máximo | 120 MB |
| Entradas máximas | 10.000 |
| Parte individual descomprimida | 60 MB |
| Total descomprimido declarado | 256 MB |
| Presupuesto real leído por carga | 256 MB |
| Ratio de compresión máximo | 500:1 |

- Se rechazan rutas absolutas y segmentos `../`.
- El presupuesto real se descuenta mientras se leen las partes, aunque el ZIP declare tamaños falsos.
- Una parte con compresión no válida o ratio excesivo se rechaza antes de descomprimirla.

### XML

- `DtdProcessing = Prohibit`.
- `XmlResolver = null`.
- `MaxCharactersFromEntities = 0`.
- `MaxCharactersInDocument = 64M`.
- Las entidades externas y los bombardeos de entidades quedan bloqueados.

### Imágenes

- Tamaño máximo: **20 MB** por imagen.
- Dimensiones máximas: **12.000 × 12.000**.
- Píxeles máximos: **25 millones**.
- Antes de `Image.FromStream` se valida la cabecera PNG, JPEG, GIF o BMP.
- Formatos vectoriales o desconocidos se omiten.
- La misma validación se aplica a imágenes de `.crb` y `.xlsx`.

### Macros y código

- Las macros se **detectan y avisan**.
- No existe ejecución de VBA, macros, complementos, DDE ni `Process.Start` desde fórmulas.
- Los objetos incrustados no se cargan ni ejecutan.
- Al guardar, el contenido no ejecutado se conserva o se descarta según el modo de guardado; nunca se ejecuta.

### Enlaces

- Solo se abren `http`, `https` y `mailto` mediante el navegador/cliente del sistema.
- `file:`, `javascript:`, `shell:` y otros esquemas se rechazan.
- Los vínculos externos del libro no se descargan al abrir.
- Se avisa de que no se actualizan.

### Fórmulas y OLE2

- Límite de fórmula: 8.192 caracteres.
- Profundidad de evaluación limitada.
- Lector OLE2 con límites de archivo, flujo, sectores, cadenas y detección de cadenas circulares.
- Libros `.xls` se abren en modo compatible de solo lectura.

### Escritura y recuperación

- Guardado atómico mediante archivo temporal y reemplazo; si el reemplazo falla se conserva el original.
- La recuperación usa nombres por sesión, DPAPI y, cuando el volumen lo permite, EFS para el temporal.
- No se aceptan metadatos ni snapshots de recuperación en texto plano.
- Si EFS no está disponible, el temporal existe brevemente antes de cifrarse y borrarse; no se considera una protección completa contra acceso físico.

## Límites de seguridad

- La aplicación **no está firmada** con Authenticode.
- **No hay sandbox del sistema operativo**: se ejecuta con los permisos del usuario.
- Los límites de lectura reducen el riesgo de archivos malformados, pero no convierten a Mini Excel en una barrera para abrir archivos hostiles.
- Para documentos no confiables, use una máquina virtual o un usuario aislado y mantenga antivirus/SmartScreen activos.

## Evidencia de pruebas

Se generaron archivos de prueba locales:

```text
security-bomb.xlsx       -> rechazado: compresión sospechosa
security-dtd.xlsx        -> rechazado: DTD prohibida
security-traversal.xlsx  -> rechazado: ruta no válida
security-macro.xlsx      -> abierto con 5 avisos; macros no ejecutadas
security-image.xlsx      -> cabecera de 20.000 × 20.000 rechazada
```

Comportamiento esperado:

```text
FAIL security-bomb.xlsx     "compresión sospechosa"
FAIL security-dtd.xlsx      "DTD está prohibida"
FAIL security-traversal.xlsx "ruta no válida"
OK   security-macro.xlsx    aviso de macros, sin ejecución
```

Los archivos de prueba no forman parte del paquete distribuible.

## Limitaciones de seguridad

- El proceso no está en un sandbox del sistema operativo.
- No se escanea con antivirus ni se verifica firma digital.
- Las imágenes se validan por cabecera y tamaño, pero GDI+ sigue siendo el decodificador final.
- El lector `.xls` no es un parser formalmente verificado.
- Los archivos cifrados no se abren.
- No hay aislamiento de red: la aplicación no descarga enlaces, pero el sistema operativo puede abrirlos si el usuario pulsa un hipervínculo permitido.
- La mejor defensa sigue siendo abrir archivos de origen desconocido en una cuenta o máquina aislada.

## Recomendaciones

1. Abrir archivos no confiables con un usuario sin privilegios.
2. Mantener Windows y .NET Framework actualizados.
3. No habilitar macros manualmente: Mini Excel no las ejecuta.
4. Revisar los avisos de apertura antes de guardar.
5. Usar copia de seguridad antes de abrir libros de origen desconocido.
6. Si se distribuye la aplicación, firmar el ejecutable y publicar hashes SHA-256.
