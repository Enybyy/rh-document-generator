# Comparación de las automatizaciones

Revisión de código del 1 de octubre de 2026.

| Proyecto anterior | Funciones observadas | Problemas | Decisión |
|---|---|---|---|
| Reemplazar formato | Lee Excel, cruza personal, reemplaza Word, empaqueta ZIP | Etiquetas partidas no se reemplazan, almacenamiento compartido, colisiones de nombre, conversión dependiente de Word | Base funcional elegida y reconstruida |
| Automatizar RH web | Cruza nombres, descarga PDFs de Drive, extrae serie/fecha/total, exporta Excel | Rutas/credenciales y sesiones ligadas al entorno, limpieza global de temporales, propósito diferente al generador Word | Extraer localmente con `receipts.py`; retirar acoplamiento a Drive |
| Contract Automation System | Django, modelos, carga DOCX/XLSX y extracción de etiquetas/columnas | No existe endpoint de generación; pruebas vacías y portada simula descarga | Retirado, consolidado en Generar RH |

## DNI

`dni-identity-validator` tiene una interfaz individual y por lotes, mapeo y función de servidor. Es una mejor base que los scripts `VERIFICAR_DNI` de la suite, ligados a archivos y rutas concretas. Se conserva como repositorio independiente, con datos ficticios explícitos, validación de formato, comparación completa de nombres y exportación. No se conservan afirmaciones de verificación estatal que la demo no realiza.

## Estructura final

Un proyecto por repositorio: `rh-document-generator`, `dni-identity-validator`, `keyboard-event-lab`, `web-scraping-selenium-pipeline` y `pos-sales-management-system`. El repositorio de contratos retirado conserva un aviso y enlace, con su historial Git disponible; no contiene otra copia activa del generador.
