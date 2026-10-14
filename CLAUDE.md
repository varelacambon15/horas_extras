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
una columna por día con horas extra, y después:

- Dias quincena trabajados: días REALES (un día marcado 2 cuenta como 1).
- Dias dobles: días pagados doble (festivo/domingo); se pagan como ordinario.
  En los .xlsm vienen como 2 en ADMIN/AJUSTE; en Varias Obras se deducen del
  importe de COSTE ADMIN (no vienen marcados por día).
- Coste dias / ajuste: coste ORDINARIO, trasladado de la plantilla (no fórmula).
  .xlsm: COSTE DIAS. Varias Obras: COSTE ADMIN − horas extra + COSTE AJUSTE.
- Coste H. Extra, Total H. Extra, % H. Extra (= Coste H. Extra / Coste total quincena).
- Coste total quincena = Coste dias / ajuste + Coste extra total
  (lo que se le carga a la obra de mano de obra; coincide con COSTE TOTAL / IMPORTE TOTAL).
- Dias atrasados, Horas extras atrasadas: enteros, informativos, deducidos de
  OBSERVACIONES; en ámbar si no cuadran con el importe.
- Atrasos/otros: trasladado de la plantilla (ATRASOS/OTROS).
- Coste extra total = Coste H. Extra + Atrasos/otros (+ nocturno y festivo en Varias Obras).
- comentarios = OBSERVACIONES.

- Horas y días siempre sin decimales.
- Solo aparecen trabajadores con horas extra o con atrasos.
- Varias Obras lleva además Técnico Obra, Obra, Coste nocturno y Dia festivo.
- El Resumen enlaza por fórmula con los totales de cada hoja.
- Alertas configurables desde el botón Ajustes (por defecto: >5 h en un día y 20 %).
- Nombre del fichero generado: Reporte_Horas_Extras_Atrasos.xlsx

## Probar sin interfaz

En el contenedor no hay Tkinter: para probar `build_excel` hay que simular el módulo
`tkinter` y llamar a `read_varias_obras` / `read_obra_xlsm` directamente.
