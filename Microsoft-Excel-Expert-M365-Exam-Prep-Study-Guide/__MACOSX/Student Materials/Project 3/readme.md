# Microsoft Excel Expert (MO-211)
```text
 ██████╗ ██████╗  ██████╗      ██╗███████╗ ██████╗████████╗    ██████╗ 
 ██╔══██╗██╔══██╗██╔═══██╗     ██║██╔════╝██╔════╝╚══██╔══╝    ╚════██╗
 ██████╔╝██████╔╝██║   ██║     ██║█████╗  ██║        ██║        █████╔╝
 ██╔═══╝ ██╔══██╗██║   ██║██   ██║██╔══╝  ██║        ██║        ╚═══██╗
 ██║     ██║  ██║╚██████╔╝╚█████╔╝███████╗╚██████╗   ██║       ██████╔╝
 ╚═╝     ╚═╝  ╚═╝ ╚═════╝  ╚════╝ ╚══════╝ ╚═════╝   ╚═╝       ╚═════╝
 ```
 📋 Student Study Guide: Project 3
Instrucciones: Este proyecto final consta de 32 tareas de nivel experto. Cubre automatización, funciones financieras complejas y gestión avanzada de libros. Marca las tareas conforme domines cada objetivo.

🛠 Recursos
Project3_datafile.xlsx

FarmBox_Sales_2019-2021.xlsx

ProjectedSales.xlsx & ImportMacro.xlsm

FarmersMarket_Registration.xlsx & 5-Year_Projection.xlsx

🚀 Project 3 Tasks
1. Funciones de Fecha, Hora y Búsqueda
[ ] 3.3.1/3.3.2 Fechas Laborales: En OrderArrival, calcula la fecha de llegada esperada usando WORKDAY con TODAY. Incluye los días de envío y el rango de feriados PostalHolidays.

[ ] 3.2.1 XLOOKUP con Manejo de Errores: En la hoja Sales, usa XLOOKUP para traer la categoría de producto de OrderArrival. Si no existe el ID, debe devolver una celda vacía ("").

[ ] 3.3.2 WEEKDAY: En FarmersMarket_Registration, usa una función anidada para devolver el nombre del día (ej. Sábado) basado en la fecha de apertura y el rango DaysData.

[ ] 3.3.2 WORKDAY (Restar días): Calcula la fecha límite de registro restando 25 días laborables a la fecha de apertura, considerando el rango Holidays.

2. Tablas y Gráficos Dinámicos Avanzados
[ ] 4.2.5 Campo Calculado: En la tabla SalesDiscounts, añade el campo "5% discount" (Total menos 5%). Aplica formato de Contabilidad sin decimales.

[ ] 4.1.1/4.1.2 Gráfico Combinado: Convierte el SalesDiscounts PivotChart a un gráfico combinado: "Sales with existing discounts" como Línea con Marcadores y "5% discount" como Área en Eje Secundario. Muévelo a una hoja llamada "Sales comparison".

[ ] 4.3.4 Drill Down: Filtra el gráfico ByCounty por "Whatcom County" y expande (drill down) hasta el nivel de Ciudad.

3. Fórmulas Financieras y Análisis What-if
[ ] 3.4.3 Función NPER: Calcula el plazo del préstamo en meses en la hoja Financing. El pago se realiza al inicio del periodo (Type = 1).

[ ] 3.4.4 Función PMT: Calcula la cuota mensual del préstamo para el puesto del mercado en la celda B14.

[ ] 3.4.2 Goal Seek: Ajusta el monto del préstamo en Financing para que el pago mensual en B14 sea exactamente ($500).

[ ] 3.4.2 Escenarios (Scenario Manager): Crea los escenarios "Current", "10% increase" y "10% decrease" para las ventas de agosto. Genera un Resumen de Escenario usando E7 como celda de resultado.

4. Auditoría, Consolidación y Visualización
[ ] 2.3.1/2.3.3 Formato Condicional Top %: En Sales by City, borra formatos previos en C4:C21. Aplica formato al Top 25%: Negrita y color Verde Oscuro (Énfasis 6, Oscuro 50%).

[ ] 2.2.2 Validación con Mensajes: Crea una lista desplegable en A28 (Sales by City). Configura el mensaje de entrada "Select a city..." y un mensaje de error de estilo "Stop".

[ ] 3.1.1 SUMIFS/AVERAGEIFS Dinámico: En B28, suma el total de ventas de productos de miel (ID inicia con "H") para la ciudad en A28. Si está vacío, muestra "No city". Repite para el promedio en C28.

[ ] 4.1.2 Gráfico de Cascada (Waterfall): En Revenue & Costs, crea el gráfico y define "Total Revenue", "Total Costs" y "Profit" como totales (Set as total).

[ ] 3.4.1 Consolidación: En FarmBox_Sales, usa la herramienta Consolidar en la hoja Summary para unir los datos de 2019, 2020 y 2021 con vínculos al origen.

5. Macros (VBA) y Configuración del Entorno
[ ] 1.2.4 Opciones de Cálculo: Activa cálculo Automático excepto en Tablas. Prueba el cálculo manual (F9) en Projected Sales y vuelve a Automático.

[ ] 1.2.4 Cálculos Iterativos: En 5-Year_Projection, activa el cálculo manual y las Iteraciones (máximo 4). Resuelve la referencia circular en C2.

[ ] 3.6.1/3.6.2 Grabar Macro: En ProjectedSales.xlsm, graba la macro "MoneyTotal" que aplique formato Contabilidad, Negrita e Cursiva.

[ ] 3.6.3 Editar Macro: Abre el Editor de Visual Basic (VBA) y modifica "MoneyTotal" para quitar el formato de cursiva. Ejecútala en N3:N15.

[ ] 1.1.1 Copiar Módulos VBA: Copia el módulo modColumnFormatting de ImportMacro.xlsm a ProjectedSales.xlsm. Ejecuta la macro ColumnTotals en el rango B16:N16.

6. Gestión de Versiones y Seguridad
[ ] 1.2.3 Estructura del Libro: Protege la estructura del libro FarmBox_Sales_2019-2021 (sin contraseña) para evitar que se muevan o eliminen hojas.

[ ] 1.1.4 Historial de Versiones: Configura el AutoRecover a 1 minuto. Sube FarmersMarket_Registration a OneDrive, activa el historial de versiones, borra una celda, recupera el dato de una versión anterior y comparte el link de visualización.

# Certificación: Al completar estos tres proyectos, habrás cubierto el 100% de los dominios del examen MO-211.