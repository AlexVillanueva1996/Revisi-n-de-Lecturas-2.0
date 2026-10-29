# Corrector Académico de Word

Procesador **local** de lecturas `.docx`, con lanzador de Windows, auditoría de cambios, formato académico, diagramas y comprobación por PDF. Admite cualquier número de lecturas. Nunca sobrescribe los originales.

## Instalación automática en Windows

Use **Corrector_Academico_Instalador.exe** (o extraiga el ZIP del mismo nombre y abra ese ejecutable). Es un instalador con descarga automática para **Windows 10/11 de 64 bits** (Windows 10 1809 o posterior). El instalador está compilado, pero todavía **no se ha ejecutado en un Windows real** durante esta entrega en Linux.

1. Abra el instalador, seleccione la carpeta y pulse Instalar. Necesita Internet durante esta preparación; puede descargar varios cientos de MB. Windows puede pedir permisos para instalar las herramientas.
2. El asistente prepara **Python 3.12 x64, LibreOffice, Graphviz, Java Temurin y LanguageTool Java 6.6 local**, además de todas las bibliotecas Python, incluyendo el adaptador lingüístico y pywin32. Si falta WinGet, intenta instalar o reparar el administrador oficial de Microsoft.
3. Comprueba los ejecutables, las bibliotecas y el checksum de LanguageTool. Inicia un servidor solo local, prueba español e inglés y procesa dos documentos sintéticos para comprobar Word, PDF, diagramas e integridad. Solo guarda `instalacion.json` si pasan estas comprobaciones. Ante un fallo muestra el error y guarda `instalacion.log`.
4. Se crean accesos directos en el escritorio y el menú Inicio. Abra el programa, coloque el índice en `ENTRADAS\INDICE\Indice_Curso.docx` y los Word en `ENTRADAS\LECTURAS`. No necesita introducir comandos.
5. Las siguientes ejecuciones inician y detienen automáticamente su propio LanguageTool local. No requieren conexión para procesar documentos; nunca se envían a la nube.

Si trabaja con el código fuente en lugar del `.exe`, haga doble clic en **INSTALAR_TODO.bat**. También **INICIAR_CORRECCION.bat** detecta una instalación ausente y ofrece el mismo proceso automático. No se limita a descargar bibliotecas: configura los motores externos. Si una política de la organización impide instalar programas o ejecutar scripts locales, la instalación se detiene y registra el motivo; no desactiva esas políticas.

La instalación es por usuario en `%LOCALAPPDATA%\CorrectorAcademico` de forma predeterminada. Las dependencias externas pueden requerir permisos del sistema. Se conservan `config.json`, ENTRADAS y SALIDAS al reinstalar. El desinstalador conserva los documentos y la configuración; no desinstala las herramientas compartidas con otros programas. ZIP funciona sin WinRAR: RAR es opcional y no se instala una licencia comercial. Tampoco se instala Mermaid CLI, porque este programa renderiza mediante Graphviz.

El código de salida del procesamiento `2` significa revisión manual pendiente; `1` indica un error; `0` indica comprobaciones sin incidencias. Una instalación correcta no equivale a que todas las lecturas estén académicamente aprobadas.

## Qué hace y cómo protege el contenido

- Descubre recursivamente las lecturas, ignora bloqueos `~$`, ordena números de archivo naturalmente y detecta duplicados por SHA-256. Continúa con el resto si un Word está dañado.
- Extrae el índice oficial del curso; no depende de los doce títulos del curso inicial ni inventa un índice si falta. Compara número, título y vocabulario del contenido. Una correspondencia dudosa conserva el título y exige revisión manual. Es una **heurística de correspondencia**, no una comprobación semántica infalible.
- Capitaliza encabezados, títulos e índice: palabras principales con mayúscula; conectores interiores en minúscula; siglas y nombres especiales configurables. No aplica este estilo a párrafos de cuerpo. Las citas dentro de encabezados se conservan y se señalan.
- Reconoce capítulos romanos y subcapítulos decimales explícitos; comprueba orden, duplicados y padres. Traslada un índice final reconocido a la portada y sincroniza sus entradas con encabezados reales. Conserva entradas sobrantes y las señala en lugar de inventar capítulos. No calcula números de página para el índice: contiene títulos, sin números de página potencialmente desactualizados.
- Reordena bloques completos de Abstract, Resumen Ejecutivo, capítulos y Referencias; no resume su contenido. Separa las secciones mediante `page_break_before`. Portada, Abstract, Resumen y capítulo I empiezan normalmente en páginas 1, 2, 3 y 4. Si una sección necesita más páginas, mantiene el texto y desplaza las siguientes.
- Aplica A4, Arial 11, interlineado 1.5 y cuerpo justificado, con títulos jerárquicos. Conserva los énfasis internos, hipervínculos, campos, fórmulas, superíndices y subíndices. El formato interno de objetos protegidos prevalece sobre uniformar la fuente. No altera encabezados/pies originales.
- Corrige espacios duplicados en texto no protegido y aplica únicamente el diccionario de reemplazos explícitamente autorizado en `config.json`. Sugiere repeticiones y cambios lingüísticos que necesitan criterio editorial. No elimina palabras para “mejorar” un argumento.
- Conserva íntegramente párrafos con citas, cifras, ecuaciones y referencias, y todos los datos de tablas. La protección es deliberadamente conservadora: algunos errores tipográficos se dejan para revisión humana.
- Mantiene tablas editables, combina sus anchos proporcionalmente, repite la cabecera y evita dividir filas cuando caben. Enumera títulos y conserva notas. Las celdas combinadas se conservan. Tablas anidadas y filas más altas que una página requieren revisión.
- Valida los archivos de imagen con Pillow, calcula resolución efectiva, ajusta dimensiones a márgenes y une títulos con objetos. Si falta un título, propone uno a partir del encabezado contextual y solicita confirmación semántica. No afirma reconocer automáticamente el contenido de una imagen.
- Antes de cambiar el documento guarda su estructura y XML en `TEMPORALES`. Después compara **cada párrafo**, las tablas y datos, los medios originales por hash, relaciones, hipervínculos, campos, ecuaciones y las partes de encabezados, pies, notas y comentarios. Reabre el Word guardado y repite la comparación. Si algo se pierde, descarta el candidato de entrega y conserva una copia del original en `ERRORES`.

## Revisión ortográfica y gramatical local

**El instalador configura LanguageTool automáticamente.** Al abrir `INICIAR_CORRECCION.bat`, el lanzador elige un puerto libre en loopback, inicia el motor Java, espera su disponibilidad y conecta el corrector. Al terminar, cierra únicamente el proceso que inició. No abre el servidor a la red, no se usa la API pública y no se descargan modelos durante la revisión de documentos.

El Abstract se revisa con `en-US`; el Resumen y los capítulos con `es`. Las sugerencias que podrían alterar una afirmación se dejan para revisión humana. Puede agregar errores inequívocos conocidos a `correccion.reemplazos_seguros`, por ejemplo `{"planficación": "planificación"}`. No se aplican a citas ni referencias.

Para desarrollo en Linux o instalaciones manuales, `python -m src.main` sigue funcionando sin Java: en ese caso deja explícitamente pendiente la revisión lingüística completa si no se configura un servidor. También admite un servidor local existente en `herramientas.language_tool_url` (por ejemplo `http://127.0.0.1:8081`). La instalación automática no sobrescribe esa configuración: usa una copia en memoria con su servidor efímero al ejecutar `src.arranque`.

El backend LanguageTool real no se ha ejecutado aquí. Se han verificado la lógica del lanzador y su manejo de fallos; la prueba real del backend forma parte de la comprobación que el instalador ejecuta en el Windows de destino antes de marcarse listo.

## Mermaid y legibilidad

Se admiten bloques con fences ` ```mermaid ` / ` ``` ` en párrafos o con saltos de línea dentro de un párrafo, y flowcharts sin fences cuando el bloque es inequívoco. El motor usado es **Graphviz**, completamente local. No necesita navegador, Chromium ni Mermaid CLI.

El parser verifica un subconjunto estricto de Mermaid: `flowchart`/`graph` con LR, RL, TB, TD o BT; nodos identificados; rectángulos, decisiones y elipses; cadenas; conexiones normales, gruesas, discontinuas y sin flecha; etiquetas de conexiones y subgrupos anidados. Conserva textos, dirección, estilos de aristas, decisiones y pertenencia a subgrupos. Compara nodos y aristas con el SVG realmente generado. Archiva el código, el grafo JSON, DOT, SVG, PNG y las medidas.

Compara la disposición original y vertical; calcula el tamaño de letra **después de escalar a la anchura y altura útiles de A4**. Solo inserta PNG a 300 DPI si el texto queda al menos a 10 pt. Verifica nuevamente la escala de la imagen en el PDF final. Si es demasiado grande o la sintaxis no se reconoce con seguridad, conserva el código y explica el motivo. No divide automáticamente diagramas de decisiones ni elimina nodos.

**Límites:** secuencias, Gantt, HTML complejo, estilos/directivas Mermaid, enlaces `click`, iconos y sintaxis fuera del subconjunto se conservan para revisión. `mmdc` se detecta si existe, pero **no se utiliza como sustitución sin una validación estructural equivalente**. No es un renderizador universal de Mermaid. Los diagramas Mermaid dentro de tablas/cuadros y los diagramas ASCII no inequívocos deben revisarse manualmente; no se intenta inferir sus relaciones. Aunque se verifiquen estructura y escala, flechas, cruces, recortes originales y correspondencia académica de figuras requieren examen humano de las vistas previas.

## Resultados e informes

Cada ejecución tiene un identificador único; los resultados anteriores no se mezclan ni se sobrescriben:

```text
SALIDAS/
  WORD/<ejecucion>/Lectura 01 - Revisada.docx
  WORD/<ejecucion>/PENDIENTES/Lectura 02 - Revisada.docx
  INFORMES/<ejecucion>/Lectura_01.json y .md
  INFORMES/<ejecucion>/Informe_Global.json y .md
  INFORMES/<ejecucion>/procesamiento.log
  VISTAS_PREVIAS/<ejecucion>/Lectura 01/ (PDF y PNG por página)
  COMPRIMIDO/<ejecucion>/Lecturas_Corregidas.zip
TEMPORALES/<ejecucion>/ (instantáneas y auditoría, no se comprimen)
```

La numeración de salida sigue el orden de archivos detectados; el número de lección oficial se registra por separado. El nombre de salida no constituye una asignación académica.

Los informes incluyen títulos y estructura, cambios antes/después, sugerencias, tablas, figuras, diagramas, páginas originales/finales, resultados de integridad, comprobaciones geométricas, incidencias y archivos generados.

Estados:

- **APROBADO**: pasaron las verificaciones configuradas y no quedaron incidencias detectadas. No equivale a certificación semántica humana.
- **APROBADO CON OBSERVACIONES**: categoría reservada para futuras observaciones no bloqueantes; la implementación actual trata las dudas detectadas como revisión manual.
- **REQUIERE REVISIÓN MANUAL**: resultado útil en `PENDIENTES`; faltan comprobaciones o hay dudas académicas, lingüísticas, estructurales o visuales.
- **ERROR DE PROCESAMIENTO**: no se pudo completar el flujo o falló integridad. No se incluye el candidato en los entregables aprobados.

`Lecturas_Corregidas.zip` incluye solo documentos aprobados, nunca PDFs, código intermedio ni informes. Puede quedar **vacío si ninguno es aprobable**, lo cual se indica con el contador del resumen. Para obtener además un paquete claramente marcado con los documentos pendientes, active `salida.incluir_pendientes_en_zip`; se crea `Lecturas_Pendientes_Revision.zip` separado. Si `rar`/WinRAR está instalado, también intenta un RAR de los aprobados y verifica el archivo. RAR es opcional, ZIP no depende de WinRAR.

La comprobación visual usa un PDF real: detecta páginas vacías, inicios de sección, texto e imágenes fuera de márgenes y posibles superposiciones. Ignora correctamente los encabezados y pies conocidos. Genera PNG de todas las páginas. Reintenta tras retirar saltos manuales redundantes si hay errores. No puede certificar automáticamente todas las propiedades visuales o semánticas de una figura rasterizada; esas limitaciones quedan explícitas en el informe.

## Cambiar opciones o usar otro curso

Edite `config.json` con un editor de texto y respete la sintaxis JSON (sin comentarios ni comas finales). Las claves y tipos se validan; los valores inválidos detienen el proceso antes de modificar documentos. Puede ajustar márgenes, fuente, interlineado, glosario de siglas/nombres, inclusión de subcapítulos en índice, herramientas, tiempos máximos, resolución y paquetes. Las reglas de integridad y protección de citas no se pueden desactivar.

Para otro curso, sustituya el índice y las lecturas de **ENTRADAS**. No cambie el código ni el listado de títulos en Python. No coloque documentos revisados nuevamente entre las entradas si no desea procesarlos como nuevas lecturas.

Si una herramienta no está en PATH, escriba su ruta completa en `herramientas.soffice` o `herramientas.dot`, con barras `/` o barras dobles `\\` en JSON. Los `.bat` y `.cmd` no son ejecutables válidos para estas rutas; use el `.exe` real.

## Uso por terminal y pruebas

Desde esta carpeta:

```bash
python -m venv .venv
# Windows: .venv\Scripts\python. En Linux/macOS: .venv/bin/python
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m src.main --diagnostico
.venv/bin/python -m src.main
.venv/bin/python -m unittest discover -s tests -v
```

Ejemplo **sintético**, sin tocar las entradas reales:

```bash
python -m src.main --crear-ejemplo
python -m src.main --raiz TEMPORALES/demo
```

La generación se niega a sobrescribir una demo existente. Incluye seis capítulos, citas, una tabla con cifras, un hipervínculo, dos figuras (una regenerada desde Mermaid) y un segundo documento sencillo. No son las doce lecturas reales.

Puede usar `--raiz "otra carpeta"` para separar cursos y `--config "otra_config.json"` para otra configuración. `--abrir-salidas` abre el explorador únicamente en Windows.

## Problemas frecuentes

| Mensaje | Acción |
|---|---|
| Falta índice o no contiene títulos | Guarde el índice en `.docx` con entradas `Lección 01: Título`; evite índices que sean solo una imagen. |
| No hay lecturas | Coloque los `.docx` en `ENTRADAS/LECTURAS`; cierre Word para evitar archivos bloqueados. |
| Falta LibreOffice/Word | Vuelva a ejecutar `INSTALAR_TODO.bat` y consulte `instalacion.log`. El resultado no se aprueba sin renderizado. |
| Falta `dot` | Ejecute `INSTALAR_TODO.bat` para reparar Graphviz. El código Mermaid se conserva mientras falte. |
| Revisión lingüística pendiente | Ejecute `INSTALAR_TODO.bat` para reparar la instalación; consulte `TEMPORALES/arranque/languagetool.log`. |
| ZIP vacío | No hubo documentos aprobados; lea los informes o habilite el ZIP separado de pendientes. |
| Diagrama ilegible/complejo | Revise su vista previa y código archivado. No elimine nodos ni reduzca la tipografía por debajo del mínimo. |
| Correspondencia dudosa | Corrija el índice, número o título de entrada tras revisar su contenido. No se asigna por nombre de archivo solamente. |
| Error de integridad | No use el candidato de `TEMPORALES`; el original se conservó. Revise el informe antes de cualquier transformación adicional. |

Los documentos con macros, paquetes ZIP anómalos, objetos complejos o relaciones externas de imágenes/plantillas no se ejecutan ni se descargan. Para esas relaciones se omite el renderizado y se solicita revisión. No hay módulo de IA ni transmisión de documentos implementados.

## Construir el instalador (desarrolladores)

Con Python y NSIS 3 instalados, ejecute `python instalador/construir.py --salida /ruta/Corrector_Academico_Instalador.exe`. Genera el `.exe` y un `.zip` con ese ejecutable. El payload excluye dependencias instaladas, documentos del usuario, cachés y resultados. La configuración de fábrica se incluye como `config.predeterminada.json` y solo se copia sobre `config.json` cuando este no existe. El bootstrap contiene un checksum SHA256 fijo de LanguageTool y conserva la comprobación de paquetes de WinGet; no usa opciones para omitir hashes o TLS.

`tests/validar_instalador.ps1` verifica sintaxis PowerShell, manifiesto, caché, rechazo de hashes incorrectos y códigos de error. Las pruebas de Python comprueban la lógica de arranque usando procesos/respuestas controlados; no sustituyen una instalación completa real en Windows. El `.exe` no tiene firma de editor comercial.
