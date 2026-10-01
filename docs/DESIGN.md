# Dirección de diseño

Un repositorio para la aplicación Generar RH, con su demo en la raíz.

## Generar RH
Qué es: genera documentos Word por lote usando Excel y una plantilla. Para quién: equipos administrativos y clientes de automatización. Objetivo: probar la selección, revisión y descarga antes de instalar Python.

Colores: tinta #19334A (texto/navegación), niebla #EEF3F6 (fondo), papel #FFFFFF (documento), petróleo #087F75 (acción), gris #526675 (secundario), ámbar #9A5B00 (incidencias).
Tipografía: Segoe UI para controles; Georgia para el documento, evocando la salida Word.
Layout: escritorio administrativo alineado a la izquierda, con tabla editable y vista previa de papel.

```text
marca                         demo / código
título                       descargar
[registros seleccionables] | [documento]
[cargar Excel/Word]         | [campos]
```
Principio: el documento final dirige la experiencia. Sin estadísticas de ahorro inventadas ni generación simulada mediante temporizadores.

## Contratos y salida visible
Mantener tinta, papel y petróleo. Priorizar capacidades reales de Python sobre el rótulo de demo. Contrato A4 de dos hojas con cláusulas numeradas y firmas; navegación de hojas en vista previa y documentos generados visibles tras descargar el lote. Una plantilla personalizada conserva su Word: su vista es un resumen de campos, no una reproducción de maquetación.
