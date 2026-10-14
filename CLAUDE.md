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

El formato de referencia es la hoja 26235 que preparó el usuario (2Q septiembre 2026).
Cada hoja de obra: Trabajador, Ajustero, Supervisor, Oficio, Sal. Dia, Coste/Hora,
una columna por día con horas extra, Dias quincena trabajados, Coste dias / ajuste
(= Coste total quincena − Coste H. Extra), Coste H. Extra, Total H. Extra, % H. Extra,
Coste total quincena, Dias atrasados, Horas extras atrasadas, Atrasos/otros
(= días × Sal. Dia + horas × Coste/Hora), coste total (= Coste H. Extra + Atrasos/otros),
comentarios.

- Solo aparecen trabajadores con horas extra o con atrasos.
- Varias Obras lleva además Técnico Obra, Obra, Coste nocturno y Dia festivo.
- El Resumen enlaza por fórmula con los totales de cada hoja.
- Alertas configurables desde el botón Ajustes (por defecto: >5 h en un día y 20 %).

## Probar sin interfaz

En el contenedor no hay Tkinter: para probar `build_excel` hay que simular el módulo
`tkinter` y llamar a `read_varias_obras` / `read_obra_xlsm` directamente.
