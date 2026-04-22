# Microsoft Excel Expert (MO-211)
```text
 ██████╗ ██████╗  ██████╗      ██╗███████╗ ██████╗████████╗    ██████╗ 
 ██╔══██╗██╔══██╗██╔═══██╗     ██║██╔════╝██╔════╝╚══██╔══╝    ╚════██╗
 ██████╔╝██████╔╝██║   ██║     ██║█████╗  ██║        ██║        █████╔╝
 ██╔═══╝ ██╔══██╗██║   ██║██   ██║██╔══╝  ██║        ██║       ██╔═══╝ 
 ██║     ██║  ██║╚██████╔╝╚█████╔╝███████╗╚██████╗   ██║       ███████╗
 ╚═╝     ╚═╝  ╚═╝ ╚═════╝  ╚════╝ ╚══════╝ ╚═════╝   ╚═╝       ╚══════╝
 ```
 📋 Student Study Guide: Project 2
Instrucciones: Este proyecto contiene 32 tareas avanzadas. Utiliza los archivos de recursos y marca las casillas conforme progreses en la configuración de la arquitectura de datos.

🛠 Recursos
Project2_datafile.xlsx

Products_List.xlsx

Employee_Information.xlsx (Para Power Query)

FarmersMarket_Report.xlsx

🚀 Project 2 Tasks
1. Gestión de Datos y Power Query
[ ] 2.1.1 Flash Fill Avanzado: En la hoja Employees, extrae el ID (4 dígitos) y crea el nombre completo en formato: Nombre I. Apellido (ej. Liane B. Cormier). Asegura mayúsculas correctas.

[ ] 2.2.5 Eliminar Duplicados: Limpia el rango A2:C23 en la hoja Employees.

[ ] 1.1.2 Power Query Editor: Crea una nueva consulta vinculada a Employee_Information.xlsx. Omite la primera fila y columnas/filas vacías. Usa los encabezados de columna proporcionados.

2. Funciones de Búsqueda y Lógica Avanzada
[ ] 3.2.1 XLOOKUP Multi-columna: En EmployeeRates_Benefits, usa una sola función para traer Nombre y Apellido desde EmployeeAddressPhone. Si no existe, devuelve "not found".
[ ] 3.1.1 COUNTIFS Dinámico: En B23 (EmployeeRates_Benefits), cuenta el personal de "Delivery" con sueldo $\ge 19$/hr usando referencias de celda.
[ ] 3.1.1 MAXIFS: En B27, encuentra el sueldo máximo para "Administration" con "Benefit Package C".
[ ] 3.1.1 Lógica Anidada (IF + AND + NOT): En H4:H18, devuelve "Yes" si el rol es "Delivery" Y NO tiene el paquete de beneficios C.
[ ] 3.1.1 Lógica (IF + OR): En la columna I, devuelve "Yes" si el paquete es A o B. De lo contrario, deja la celda vacía.

3. Series, Auditoría y Formato Condicional
[ ] 2.1.2 Series Lineales (Ventas): En Sales and Bonuses, genera una serie en la columna K (Paso: 75, Inicio: 300, Fin: 1650).

[ ] 2.1.2 Series Lineales (Bonos): En la columna L, genera una serie (Paso: 0.25%, Inicio: 1.5%, Fin: 6%).

[ ] 3.2.1 Búsqueda de Rango: En C3:C24, busca el porcentaje de bono correspondiente al monto de venta en la tabla creada en la columna L.

[ ] 3.5.1/3.5.3 Auditoría: Rastrea precedentes en H3, dependientes en D7 y corrige el error en E7.

[ ] 2.3.3 Gestión de Reglas: Reordena las reglas de formato condicional en la columna Sales para que los valores > 700 permanezcan con texto negro.

[ ] 2.3.2 Formato Condicional (Fórmula): Aplica Negrita al nombre del empleado si su venta es mayor al promedio en H2.

4. Análisis con Tablas y Gráficos Dinámicos
[ ] 4.2.2 Jerarquías en PivotTable: En Sales Quantity, añade Product category sobre Product name, colapsa la categoría y filtra por "Honey" y "Jams and Jellies".

[ ] 4.3.1/4.3.4 PivotChart de Pastel: Crea el gráfico, quita título/leyenda, añade etiquetas externas con Nombre y Valor. Muestra el detalle de "Wildflower honey" en una nueva hoja.

[ ] 4.2.4 Agrupación Numérica: En Qty by City, agrupa el campo Quantity de 1 a 10 en grupos de 5.

[ ] 4.2.2 Configuración de Impresión: En Summer Sales, agrupa por meses y configura las etiquetas de fila para que se repitan en cada página impresa.

[ ] 4.2.6 Cálculos de Valor: Añade Order total por segunda vez y configúralo como "% of Row Total" con 1 decimal y nombre "% of order".

[ ] 4.2.3 Slicers con Lógica: Crea segmentadores para City y Sale Date. Oculta elementos sin datos y excluye las ciudades "Bow" y "Conway".

5. Visualización Avanzada y Formatos Custom
[ ] 4.1.2 Gráfico de Embudo (Funnel): En Product in stock, crea el gráfico en orden cronológico (Semana 1 arriba).

[ ] 4.1.2 Gráfico Solar (Sunburst): En Sales by Month, agrupa por mes (Jun-Ago). Aplica Estilo 6 y título "Summer Farmers Market Sales".

[ ] 2.2.1 Formato de Día: En Sales by Day, aplica formato personalizado para mostrar solo el nombre del día (ej. Lunes).

[ ] 2.2.1 Formato Contable Custom: En B2:B29, aplica: $ #,##0 para que el cero se vea como $ 0 y el signo de dólar tenga un espacio inicial.

[ ] 4.1.2 Gráfico de Cajas y Bigotes: Crea el gráfico incluyendo media, marcadores, puntos internos y valores atípicos (Cálculo de cuartil exclusivo).

6. Estructura, Subtotales y Seguridad
[ ] 2.2.4 Esquemas y Subtotales: En Baker Heights Sales, inserta subtotales (Suma) de Cantidad y Total por producto. Colapsa para ver solo subtotales y oculta columnas innecesarias.

[ ] 3.3.1 Tiempo Real y Formato: En F1 de FarmersMarket_Report, inserta NOW() y formatea como: 4 Jan 2023 4:35 PM.

[ ] 1.2.2 Protección de Rangos: En la hoja Sales, permite editar solo E4:E21 con el nombre "QuantitySold" y contraseña "password". Protege la hoja completa.

[ ] 1.2.1 Cifrado del Libro: Protege el archivo FarmersMarket_Report.xlsx con la contraseña "password".

Tip de Experto: Al usar Power Query, asegúrate de que la ruta de los archivos de origen sea correcta en tu entorno local para evitar errores de conexión.