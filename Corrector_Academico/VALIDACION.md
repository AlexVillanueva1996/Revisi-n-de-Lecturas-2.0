# Validación de entrega — 29 de octubre de 2026

Se desarrolló y verificó en Linux con Python 3.12.14, python-docx 1.2.0,
lxml 6.1.1, Pillow 12.3.0, PyMuPDF 1.26.6, Graphviz 2.42.4 y LibreOffice.
No se proporcionaron lecturas reales. Todos los documentos utilizados son sintéticos.

## Comprobaciones ejecutadas

- Instalación de `requirements.txt` en un entorno virtual nuevo: correcta.
- `pip check`: sin dependencias rotas.
- Ejecución completa y repetible de `scripts/preparar_nube.sh`: correcta.
- `.venv/bin/python -m unittest discover -s tests -v`: **39 pruebas aprobadas**, sin fallos ni pruebas omitidas.
- El registro de la última verificación del entorno se guarda en `TEMPORALES/verificacion_entorno.log` (archivo local generado).
- Dos documentos sintéticos procesados de principio a fin, con seis capítulos cada uno: Word válidos, reportes JSON/Markdown, PDF real y vistas PNG.
- Páginas finales de la demo: 11 y 10. Abstract en página 2, Resumen en página 3, capítulo I en página 4. Los capítulos posteriores y las referencias comienzan en páginas independientes; el capítulo I con objetos ocupa dos páginas sin perder contenido.
- Comparación de integridad: texto, párrafos, citas, cifras, tablas con celdas combinadas, medios originales por hash, hipervínculos, campos, ecuaciones, referencias y encabezados conservados.
- Flowchart horizontal renderizado en ambas orientaciones: se seleccionó vertical; nodos, decisiones, etiquetas y conexiones comprobados; texto efectivo de aproximadamente 11 pt verificado otra vez en el PDF.
- Pruebas de diagramas con ramificaciones, subgrupos, etiquetas largas, sintaxis desconocida, bloques incompletos y diagramas que no caben con texto legible. Los casos no seguros conservan el código original.
- Abstract de varias páginas: conserva el texto y desplaza las secciones siguientes sin reducir el cuerpo tipográfico.
- Manejo de lecturas incompletas, archivos duplicados y archivos dañados: incidencias registradas, sin aprobaciones falsas y continuando con el resto.
- Paquetes ZIP abiertos y comprobados con `testzip`; solo contienen los Word expresamente seleccionados. El ZIP de pendientes se crea por separado cuando se habilita.
- Revisión visual adicional de la portada, la página con tabla y diagrama, y la imagen original de la demo: diagrama vertical legible; imagen sintética original con texto pequeño que debe permanecer pendiente de revisión humana.

## Resultado real de la demo

Ambas lecturas quedan en **REQUIERE REVISIÓN MANUAL**, con integridad correcta y sin errores geométricos de PDF. No hay aprobados: falta revisión lingüística completa con LanguageTool local y hay figuras que requieren juicio visual/semántico. Por ello `Lecturas_Corregidas.zip` contiene cero documentos; al habilitar la opción correspondiente, `Lecturas_Pendientes_Revision.zip` contiene los dos Word generados.

## No verificado o fuera del alcance de la automatización segura

- Ejecución mediante doble clic del `.bat` en un Windows real.
- Automatización Microsoft Word COM y creación de RAR mediante WinRAR.
- Un servidor real de LanguageTool español/inglés: no estaba instalado ni configurado. Se implementó el adaptador local; no se afirma haber validado su backend lingüístico.
- Revisión de las doce lecturas reales, que no se han proporcionado.
- Semántica completa, exactitud de las citas, correspondencia de fuentes y OCR de figuras preexistentes: requieren revisión académica humana.
- Sintaxis Mermaid fuera del subconjunto documentado y diagramas ASCII ambiguos: se conservan para revisión, sin sustituirlos por figuras que podrían omitir información.

La configuración del entorno en la nube guarda `install_script`, `start_skill` y la ubicación del checkout. Guardar el borrador no publica el entorno ni sube el código a GitHub. El repositorio de partida no tenía commits; los archivos del proyecto están disponibles en el espacio de trabajo y en el ZIP de entrega.

## Actualización: instalador completo 2.1

Se añadió un ejecutable Windows real, compilado con NSIS 3.11 desde Linux, y un ZIP que contiene ese `.exe`. Es un **instalador con descarga automática**, no un paquete offline con todos los motores embebidos.

- Instala o detecta Python 3.12 x64, LibreOffice, Graphviz y Java compatible mediante WinGet; si falta WinGet, intenta su reparación oficial.
- Instala todas las bibliotecas Python, incluyendo language-tool-python 3.4.0 y pywin32 en Windows.
- Descarga LanguageTool Java 6.6 mediante HTTPS y comprueba SHA256 fijo obtenido del manifiesto público Scoop `languagetool-java`. Un checksum incorrecto detiene la instalación.
- Solo crea el marcador de éxito tras probar el servidor local en ambos idiomas y un flujo Word/PDF/diagramas/integridad con documentos sintéticos en el equipo de destino.
- Crea accesos directos, conserva configuración y entradas al reinstalar y conserva documentos al desinstalar. Inicia y cierra su propio servidor local durante cada ejecución.
- **45 pruebas Python aprobadas**, sin omisiones, tras añadir seis pruebas de arranque, protección de loopback, marcadores y fallos. Las pruebas de ciclo de vida usan procesos/respuestas controlados; no certifican un LanguageTool real.
- Script PowerShell analizado y pruebas de funciones ejecutadas con PowerShell 7.4.6 oficial, cuyo hash se verificó contra la publicación oficial. Se comprobó rechazo de URL sin TLS, checksum incorrecto y propagación de códigos de error.
- Instalador compilado y reconocido como ejecutable PE/NSIS. Las bibliotecas del adaptador lingüístico se instalaron también en el entorno virtual Linux.

**Sigue pendiente la ejecución real del `.exe` en Windows**, incluyendo WinGet, UAC, instalaciones de terceros y el backend LanguageTool. Las pruebas de sintaxis y compilación no sustituyen esa verificación. El propio instalador ejecuta la comprobación funcional en el destino y no declara éxito si falla.
