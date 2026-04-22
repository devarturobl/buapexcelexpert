# Microsoft Excel Expert (MO-211)
```text
 ██████╗ ██████╗  ██████╗      ██╗███████╗ ██████╗████████╗     ██╗
 ██╔══██╗██╔══██╗██╔═══██╗     ██║██╔════╝██╔════╝╚══██╔══╝    ███║
 ██████╔╝██████╔╝██║   ██║     ██║█████╗  ██║        ██║       ╚██║
 ██╔═══╝ ██╔══██╗██║   ██║██   ██║██╔══╝  ██║        ██║        ██║
 ██║     ██║  ██║╚██████╔╝╚█████╔╝███████╗╚██████╗   ██║        ██║
 ╚═╝     ╚═╝  ╚═╝ ╚═════╝  ╚════╝ ╚══════╝ ╚═════╝   ╚═╝        ╚═╝
 ```
📋 Student Study Guide: Project 1
Instrucciones: Este proyecto consta de 28 tareas basadas en los objetivos del examen MO-211. Utiliza los archivos de recursos proporcionados y marca cada tarea conforme la completes.

🛠 Recursos
Project1_datafile.xlsx (Carpeta Project_Files)

DiscountCodes.xlsx (Carpeta Project_Files)

🚀 Project 1 Tasks
1. Gestión de Datos y Formatos
[ ] 2.1.1 Relleno rápido (Flash Fill): En la hoja Salespeople, extrae el ID de 4 dígitos (Columna A) y el nombre completo (Nombre Apellido, Columna B) de la lista en la columna C.

[ ] 2.2.1 Formatos de número personalizados: En la hoja Salespeople, aplica un formato a A2:A16 para que los valores tengan el prefijo "ID-" (ej. ID-1234).

[ ] 2.2.1 Formatos de fecha: En la hoja Sales, formatea la columna Month para que muestre solo el nombre completo del mes (ej. Junio).

[ ] 2.2.2 Validación de datos: En la hoja Sales, restringe la columna Discount number a números enteros entre 0 y 3. Configura una alerta de "Advertencia" titulada "Invalid entry" y un mensaje de entrada titulado "Discount number".

2. Funciones de Búsqueda y Lógica
[ ] 3.2.1 Búsqueda (XLOOKUP/VLOOKUP): En la hoja Sales, usa una función para traer el nombre completo del vendedor desde la hoja Salespeople usando el Salesperson ID.

[ ] 3.4.6 Ordenar datos (SORTBY): En la hoja Products, genera una lista en F2 que ordene el rango A2:D14 ascendente por Category y luego por Product name.

[ ] 3.2.1 Búsqueda (MATCH/INDEX): En la hoja Sales, combina MATCH e INDEX para obtener los nombres de productos de la hoja Products basándote en el Product ID.

[ ] 3.1.1 Lógica (SWITCH): En la hoja Sales, columna Discount code, usa SWITCH para convertir el número de descuento: 0 → "no discount", 1 → "A", 2 → "B", 3 → "C".

[ ] 3.1.1 Lógica Anidada (IF + HLOOKUP): En la hoja Sales, obtén el porcentaje de descuento del libro DiscountCodes.xlsx. Si es "no discount", devuelve 0.

3. Auditoría y Análisis de Datos
[ ] 3.5.3/3.5.4 Auditoría de fórmulas: En la hoja Sales, usa la comprobación de errores en O4. Rastrea precedentes, cambia el valor causante por "2" y elimina las flechas.

[ ] 3.5.2 Ventana de Inspección (Watch Window): En la hoja Sales, agrega las celdas de totales de la fila 309 a la ventana de inspección. Cambia valores en L2:L4 y observa el impacto.

[ ] 2.1.3 Datos Aleatorios (RANDARRAY): En la hoja Market Stall Projections, genera una matriz de números enteros aleatorios entre 15 y 50 en el rango B3:D15.

4. Tablas y Gráficos Dinámicos
[ ] 4.2.1 Crear PivotTable: En la hoja Honey Orders (celda A1), crea una tabla dinámica llamada "HoneySales" usando los datos de Sales.

[ ] 4.2.2/4.2.4 Configurar PivotTable: Mueve Product name a Columnas, Sale date a Filas (solo meses/trimestres). Formatea como Moneda ($) y renombra a "Total of All Orders".

[ ] 4.2.3 Segmentadores (Slicers): Crea un segmentador por Product name y filtra por: Buckwheat, Clover y Wildflower.

[ ] 4.3.1/4.3.2 PivotCharts: Crea un gráfico de columnas agrupadas. Filtra para mostrar solo Junio, Julio y Agosto.

[ ] 4.1.2 Gráficos Avanzados: Crea un Histograma en la hoja Honey Sold para la columna % of stock sold. Añade etiquetas de datos al centro.

5. Fórmulas Avanzadas y Condicionales
[ ] 3.1.1 Función IFS: En la hoja Honey Sold (celda D1), usa IFS para categorizar el promedio de stock: <50%, 50% or more, 65% or more.

[ ] 3.1.1 Lógica (OR/AND): En la hoja Customers Report, marca con "Yes" si el tamaño de caja es "Large" O la frecuencia es "Weekly".

[ ] 3.1.1 Lógica (NOT): En la hoja Customers Report, marca con "Yes" si el tamaño NO es "Large" Y la frecuencia NO es "Monthly".

[ ] 3.1.1 COUNTIFS: En la celda H3 de Customers Report, cuenta cuántos clientes piden cajas "Small" con frecuencia semanal o quincenal.

[ ] 2.3.2 Formato Condicional (Fórmula): En Customers Report, aplica relleno verde claro a los nombres si cumplen la condición de pedidos "Large" o "Weekly".

[ ] 3.4.5 Función FILTER + LET: En Customer Counties, usa LET y FILTER para mostrar datos filtrados por el condado en G1. Si no hay rep asignado, muestra un guion.

[ ] 2.3.2 Formato Condicional (Fechas): En Market Report, sombrea de amarillo las filas donde la fecha sea fin de semana (Sábado o Domingo).

Nota: Consulta el Learning Directory si necesitas ayuda paso a paso para cualquier dominio específico. 