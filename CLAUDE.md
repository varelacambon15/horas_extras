# Contexto del proyecto

Programa de escritorio (Python + Tkinter) que genera el informe de horas extra de
cada quincena a partir de los reportes de nómina de las obras.

- Empresa en República Dominicana: **todas las cantidades son pesos dominicanos (RD$)**,
  nunca euros.
- Obras con código de 5 dígitos (p. ej. 26235, 26337, 26227, 25484). El código se
  saca del nombre del fichero de cada reporte de obra.
- Quincenas: 1Q y 2Q de cada mes.

## Reportes de entrada

- **Varias Obras** (`.xlsx`, emitido por el programa de nómina): fila 1 con fechas
  ("Viernes | 14-08-2026"), fila 2 con cabeceras; por cada día un bloque
  ADMIN, AJUSTE, OBRA 1, EXTRA 1, OBRA 2, EXTRA 2. Columnas finales:
  TOTAL DIAS TRABAJADOS, TOTAL HORAS EXTRA, COSTE ADMIN, COSTE AJUSTE,
  ATRASOS/OTROS, IMPORTE TOTAL, PRESTAMO, OTRAS DEDUCCIONES, TOTAL A PAGAR, OBSERVACIONES.
- **Reportes de obra** (`.xlsm`, los envían desde obra): hoja `Reporte`, cabecera en
  la fila 4. Por día: ADMIN / AJUSTE / EXTRA con la fecha en la cabecera
  (a veces con día o mes de un dígito: `12/7/2026`). Columnas finales:
  TOTAL DIAS ADMIN, TOTAL DIAS AJUSTE, TOTAL H. EXTRAS, COSTE DIAS, COSTE H. EXTRAS,
  ATRASOS/OTROS, COSTE TOTAL, DESCUENTOS, PRESTAMO, TOTAL A PAGAR, OBSERVACIONES.
  Las celdas son fórmulas; se leen los valores cacheados.

## Informe de salida

El formato de referencia es la hoja 26235 que preparó el usuario (2Q septiembre 2026),
corregida según los correos con Jaume Ordinas (Director de Proyectos), que pide un
documento claro, "sin dados ni loterías": que se vea de dónde sale cada cifra y
con pocas operaciones.

Cada hoja de obra: Trabajador, Ajustero, Supervisor, Oficio, Sal. Dia, Coste/Hora,
una columna por día con horas extra, y después (en este orden):

- Dias quincena trabajados: días REALES (un día marcado 2 cuenta como 1).
- Domingos / Festivos: días pagados doble; se pagan como ordinario.
  En los .xlsm vienen como 2 en ADMIN/AJUSTE; en Varias Obras se deducen del
  importe de COSTE ADMIN (no vienen marcados por día). Solo sale si hay alguno.
- Coste dias / ajuste: coste ORDINARIO, trasladado de la plantilla (no fórmula).
  .xlsm: COSTE DIAS. Varias Obras: COSTE ADMIN − horas extra + COSTE AJUSTE.
- Total H. Extra.
- % H. Extra = Coste H. Extra / (Coste dias / ajuste + Coste H. Extra).
- Coste H. Extra.
- Dias atrasados, Horas extras atrasadas: CANTIDADES (días / horas) que se rellenan a mano.
- Otros: ATRASOS/OTROS trasladado de la plantilla.
- Total atrasos = Dias atrasados × Sal. Dia + Horas extras atrasadas × Coste/Hora + Otros.
- Coste total quincena = Coste dias / ajuste + Coste H. Extra + Total atrasos
  (lo que se le carga a la obra; coincide con COSTE TOTAL / IMPORTE TOTAL de la plantilla).
- comentarios = OBSERVACIONES.

- Cabeceras en naranja: Coste dias / ajuste, Coste H. Extra, Total atrasos,
  Coste total quincena (y sus equivalentes en el Resumen).
- Horas y días siempre sin decimales.
- Aparecen TODOS los trabajadores del reporte (con o sin horas extra / atrasos).
- Varias Obras lleva además Técnico Obra y Obra.
- Gratificación (GRATIFICACIONES del .xlsm) justo después de Sal. Dia, solo si alguien la tiene;
  explica por qué el Coste/Hora es mayor que Sal. Dia / 8.
- Resumen: Obra | Trabajadores | Trab. con HE | Coste dias quincena (ordinario de
  TODOS los trabajadores) | Coste H. Extra | Total H. Extra | % s/Total
  (= HE / (Dias + HE)) | Coste atrasos | Coste Total (Dias + HE + Atrasos).
  Enlazado por fórmula con los totales de cada hoja.
- Alertas configurables desde el botón Ajustes (por defecto: >5 h en un día y 20 %).
- Hojas: todos los días de la quincena, con día de la semana ("Vie 18/09"); domingos y
  festivos en gris (domingo, festivo fijo de RD o algún "2" ese día en el .xlsm).
- Solo las cabeceras naranjas (totales) llevan nota explicativa (NOTAS / NOTAS_RESUMEN).
  Sin hoja de Leyenda ni de Pendiente de pago (descartadas por el usuario).
- Resumen con título de la quincena (p. ej. "2Q Septiembre 2026"), sin subtítulo; la
  quincena es la del mes del día central del periodo (1Q si día <= 15).
- Nombre del fichero generado: Reporte_Horas_Extras_Atrasos_<quincena>.xlsx
  (los reportes se leen antes de pedir dónde guardar).

## Interfaz

- Los archivos se pueden añadir con los botones o arrastrándolos (tkinterdnd2).
  Soltados sobre la ventana se clasifican por nombre ("VARIAS" → Varias Obras);
  soltados sobre una sección van a esa sección.
- tkinterdnd2 se instala solo con pip si falta (solo ejecutando con Python, no en
  el .exe). Si no está disponible, el programa funciona igual con los botones.
- El .exe se genera con PyInstaller (`resource_path` usa `_MEIPASS`); para incluir
  el arrastrar y soltar hay que añadir `--collect-all tkinterdnd2`.

## Probar

- Sin interfaz: simular el módulo `tkinter` y llamar a `read_varias_obras` /
  `read_obra_xlsm` y `build_excel` directamente.
- Con interfaz: el Python por defecto del contenedor no trae Tkinter; usar
  `/usr/bin/python3.12` (tras `apt-get install python3-tk xvfb`) en un venv con
  pandas, openpyxl y tkinterdnd2, y ejecutar con `xvfb-run`.
