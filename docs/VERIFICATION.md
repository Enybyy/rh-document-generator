# Verificación

Revisión del 1 de octubre de 2026.

- 30 pruebas Python: Word con etiquetas partidas, estilo, tablas anidadas, encabezados y pies, selección, nombres repetidos, DNI con ceros, cruce ambiguo, fecha/importe, archivos corruptos y API. Incluye importes con separadores ambiguos rechazados y límite de salida.
- Extracción de serie, fecha y total verificada también en un PDF de texto ficticio real (`tests/fixtures/receipt.pdf`). Los recibos de otros diseños y los escaneos no se probaron.
- Playwright verificó edición, selección/desselección, descarga ZIP real, carga CSV con filas inválidas, interfaz de 360 px, carga XLSX/DOCX local y descarga ZIP de servidor y bloqueo al cambiar archivos después de revisarlos.
- Se abrieron ambos ZIP descargados y todos los DOCX internos con `python-docx`.
- Cero errores JavaScript en el recorrido. Capturas de escritorio y móvil inspeccionadas visualmente.

No se realizó una prueba con documentos reales de una empresa ni una revisión jurídica de su contenido. La maquetación personalizada depende de la plantilla Word del usuario; la interfaz muestra datos, no una representación exacta de esa plantilla. Publicación remota se comprueba después del push.
