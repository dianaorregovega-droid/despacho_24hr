# despacho_24hr
Despacho 24 horas sin red Colombia con y sin baterías para el 15 de noviembre de 2027

## Insumos
En el Excel “input_despacho24hv16” se encuentran los insumos que recibe el modelo para la construcción del escenario base y escenario 1.)  “Baterías que no reducen limitación de inyección de renovables”. En el excel “input_despacho24hv16_sinrestriccionesfinas” se encuentran los inusmo que recibe el modelo para la construcción del escenario 2.) “Baterías que reducen limitación de inyección de renovables”
En el excel “input_despacho24hv16_sinrestriccionesfinas” se encuentran los insumos que recibe el modelo para la simulación del escenario 2.) Baterías que reducen limitación de inyección de renovables”.
En ambos documentos se encuentran las siguientes hojas: 

Master_area: Esta hoja describe el código de cada una de las 7 áreas y define si ésta es un área del sistema eléctrico colombiano o corresponde a un enlace de interconexión con otro país. Para el caso de esta simulación no se considera la coordinación con Ecuador. 

Nombre Campo	Tipo de dato	Descripción
id_area	int	Para cada área del sistema se define un código -índice del Set A (áreas) 
name_area	Str 	Nombre del área (no entra a restricciones). Útil para reportes
type_area	Str 	Internal: Área interna del sistema eléctrico
External: Interconexión con otro país para el caso de análisis de importación y exportación o despacho coordinado. 

Master_estacion_hidrologica: Define el código de cada una de las 58 estaciones hidrológicas, asociándolas al embalse correspondiente. Esta asociación permite relacionar los caudales recibidos con el embalse asociado a cada planta.

Nombre Campo	Tipo de dato	Descripción
Id_estacionhidrologica	Int	Número de la estación hidrológica
Name_estacionhidrologica	Str 	Nombre de la estación hidrológica (no entra a restricciones). Útil para reportes
Id_embalse	Str 	Id embalse asociado a cada estación hidrológica. Índice del Set RES. Con esto se puede definir a que estación hidrológica pertenece cada embalse y, por lo tanto, que caudal recibe el embalse.  

Caudalesm3s: Describe los caudales asociados a cada estación hidrológica, previstos para el 15 de noviembre de 2027 mediante la aplicación de caudal histórico 2014-2016

Nombre Campo	Tipo de dato	Descripción
id_estacionhidrologica	Int	Número de estación hidrológica
Caudal_m3s	Float	Caudal promedio en m³/s. Parámetro Q[EH, w]. Decimal con coma. 

Hydro_with_reservoir_base: Describe los parámetros relacionados con 20 hidroeléctricas con embalse: 

Nombre Campo	Tipo de dato	Descripción
Id_planta	Str	Índice del set HRES
Name_planta	Str	Nombre de la planta hidroeléctrica con embalse
Id_area	Int	Índice del área en la que se encuentra la planta de generación
Id_estacionhidrologica	Int	Número de la estación hidrológica
Id_embalse	Str	Id del embalse asociado a la hidroeléctrica
Pmax_MW	Float	Potencia máxima de la planta definida en MW
Prod_MW_per_m3s	Int	Factor de producción promedio [MW/m3/s]
Qmin_m3s	Int	Caudal turbinable mínimo [m3/s]
Qmax_m3s	Int	Caudal turbinable máximo [m3/s]
Vmin_hm3	Int	Volumen mínimo del embalse [Hm3]
Vmax_hm3	Int	Volumen máximo del embalse [Hm3]
Vinic_hm3	Int	Volumen inicial del embalse [Hm3]. Esto se calcula como volumen mínimo + [condición inicial en pu *(volumen máximo – volumen mínimo)]
O&M_USD$_per_MWh	Int	Costo de Operación y Mantenimiento [USD/MWh]


Hydro_ror: Describe los parámetros relacionados con 32 plantas hidroeléctricas sin embalse. 5 de estas plantas equivalen a 137 hidroeléctricas sin embalse que se encontraban en operación en 2025 en cada área: 

Nombre Campo	Tipo de dato	Descripción
Id_planta	Str	Índice del set HROR
Name_planta	Str	Nombre de la planta sin embalse
Id_area	Int	Índice del área en la que se encuentra la planta de generación
Pmax_MW	Float	Potencia máxima de la planta definida en MW. Esta potencia ya limita al valor mínimo entre capacidad efectiva neta y caudal recibido* factor de producción promedio
Prod_MW_per_m3s	Int	Factor de producción promedio [MW/m3/s]
O&M_USD$_per_MWh	Int	Costo de Operación y Mantenimiento [USD/MWh]

Profile_ror_by_area: Se asume que el perfil horario respecto de la potencia máxima depende del área, por lo cual, se crea una tabla con el perfil en términos de fracción respecto de potencia máxima para cada una de las horas.

Nombre Campo	Tipo de dato	Descripción
Hour_in_day	int	Hora del día
Id_area	int	Índice del área en la que se encuentra la planta de generación
Profile	float	% de generación en la hora de la semana respecto de la potencia máxima. Se tomó como base el promedio de generación en 2025 de las hidroeléctricas sin embalse que componen el área correspondiente.

Termica_commitment_base: Describe los parámetros relacionados con 40 termoeléctricas

Nombre Campo	Tipo de dato	Descripción
Id_planta	Str	Índice del set TC (térmicas con commitment)
Name_planta	Str	Nombre de la planta de generación termoeléctrica
Id_area	Int	Índice del área en el que se encuentra la termoeléctrica
Pmin_MW	Float	Potencia mínima de la planta [MW] – Para este ejercicio se eliminan los requerimientos mínimos por lo que Pmin=0. 
Pmax_MW	Float	Potencia máxima de la planta definida en [MW]
Fuel code	Int	Código del combustible utilizado por la termoeléctrica
Heat_Rate	Float	Consumo específico [MBTU/MW]
O&M_USD$_per_MWh	float	Costo variable operación y mantenimiento [USD/MWh]
Costo_transporte	Int	Costo de transporte [USD/MBTU] se coloca el valor en cero, considerando que el costo del combustible tomado para este análisis ya incluye el costo de transporte. 
startup_cost_kUSD$	Int	Costo de arranque [Kusd]. Al ser un despacho de 24 horas, no se utiliza este parámetro

Termica_no_commitment_base: Describe los parámetros relacionados con 40 termoeléctricas

Nombre Campo	Tipo de dato	Descripción
Id_planta	Str	Índice del set TNC (térmicas no commitment)
Name_planta	Str	Nombre de la planta de generación termoeléctrica
Id_area	Int	Índice del área en el que se encuentra la termoeléctrica
Fuel code	Int	Código del combustible utilizado por la termoeléctrica
Pmin_MW	Float	Potencia mínima de la planta [MW] –Pmin=0. 
Pmax_MW	Float	Potencia máxima de la planta definida en [MW]
O&M_USD$_per_MWh	Float	Costo variable operación y mantenimiento [USD/MWh]
Heat_Rate	Float	Consumo específico [MBTU/MW]

Costo_combustible:

Nombre Campo	Tipo de dato	Descripción
Fuel code	Int	Código del combustible utilizado por la termoeléctrica
Fuel_name	Float	Nombre del combustible utilizado por la termoeléctrica. 
Fuel_unit	Float	Unidad del combustible utilizado por la termoeléctrica [por defecto se utiliza MBTU
Fuel_cost	Float	Costo de combustible [USD/MBTU]

Parametro_renovable: Describe los parámetros relacionados con 349 plantas de generación renovable eólica y solar agrupadas en 128 plantas.

Nombre Campo	Tipo de dato	Descripción
Id_planta	Str	Índice del set REN (renovables)
Name_planta	Str	Nombre de la planta de generación eólica o solar
Id_area	Int	Índice del área en el que se encuentra la planta de generación
Pmax_MW	Float	Potencia máxima de la planta definida en [MW]
Tech	Str	Tipo de tecnología [Solar / Eólica]
Heat_Rate	Float	Consumo específico [MBTU/MW]
O&M_USD$_per_MWh	float	Costo variable operación y mantenimiento [USD/MWh]

Profile_renovable: En el caso de la energía eólica, se asume un perfil horario respecto de la potencia máxima para cada área el cual se calculó en función del promedio de generación de las plantas existentes en cada área. En relación con la tecnología eólica, se calculó un único perfil para la tecnología con base en la generación de 2025 de las plantas de generación

Nombre Campo	Tipo de dato	Descripción
Hour_in_day	int	Hora del día
Id_area	int	Índice del área en la que se encuentra la planta de generación
Tech	Str	Tipo de tecnología [Solar / Eólica]
Shape_Frac	Float	% de generación en la hora de la semana respecto de la potencia máxima [valor entre 0 y 1]. 

Limites_areas: Describe los límites de intercambio entre áreas de forma que se cumpla la ecuación: LB[MW] <= intercambio neto <= UB [MW]. Si el intercambio es positivo, el área está importando y, si el intercambio es negativo el área, está exportando. Se toman los límites actuales con el fin de no considerar el escenario más conservador. 

Nombre Campo	Tipo de dato	Descripción
Id_area	Int	Índice del área en la que se encuentra la planta de generación
LB_MW	Int	Límite inferior de intercambio que se corresponde con el límite de exportación del área [MW]
UB_MW	Int	Límite superior de intercambio que se corresponde con el límite de importación del área [MW]

Demanda: Contiene el registro de las demandas horarias previstas para cada área.

Nombre Campo	Tipo de dato	Descripción
Hour_in_day	int	Hora del día
Id_area	int	Índice del área 
Demanda	Float	Demanda para cada una de las horas [MW]

## 2. Descripción general del script

El script está diseñado para la simulación del despacho económico de un solo día con Pyomo, leyendo un LIBRO EXCEL (múltiples hojas).

### 2.1. Definiciones e importaciones: 

Adicional a la preparación del entorno Python para poder hacer uso de la licencia académica de gurobi, el script inicia con la importación de módulos fundamentales de la biblioteca de Python. El módulo re permite buscar y limpiar texto, typing se usa para indicar el tipo de datos que se espera manejar, pandas para leer archivos de Excel y manejar tablas y numpy para operaciones numéricas, manejo de arrays, conversiones y valores nulos (NaN).
Luego se realiza las importaciones asociadas al modelo. ConcreteModel es donde se define el problema de optimización indicando los conjuntos - Set, variables - Var, restricciones – Constraint, función objetivo – Objective - que en nuestro caso corresponde a minimizar el costo del despacho. Suffix lo usamos para poder extraer la variable dual asociada al costo marginal. 
 

### 2.2. Configuración:
Se carga el Excel que contiene los insumos del modelo y se nombran las hojas que contiene el Excel.
 
### 2.3. Limpieza y conversión europea:
La función def_clean_columns permite limpiar los nombres de las columnas de cada una de las hojas eliminando todo lo que no sea letras, números y guiones bajos y pasando todo a minúscula. Se reemplaza los nombres originales de las columnas que trae el Excel por los nombres limpios. Ejemplo:  PMax(MW) -> pmax_mw.
La función def_to_numeric_eu convierte a formato de punto como separador décimal. Ejemplo: 1.234,56 -> 1234.56. Por su parte, la función coerce_numeric_columns convierte columnas object (texto) a numérico con el fin de arregla cualquier número que esté aun mal formateado, sin ajustar las columnas skip_cols. 
La función def_postprocess_types revisa todas las hojas posibles que podría significar hora y los convierte a Int64. 
 

### 2.4. Carga Excel:
La función def_normalize_sheet_name convierte el nombre de la hoja a texto y lo deja sin espacios al inicio ni al final. La función def load_all_tables_from_excel crea un diccionario donde la clave es el nombre normalizado de la hoja y el valor el nombre real de la misma.
 
Con lo anterior, se lee el Excel y se aplica las funciones previamente mencionadas para guardar la hoja limpia usando el nombre normalizado como clave.  
 
### 2.5. Funciones auxiliares:

	Def pick_first_existing recibe el DataFrame y lista de nombres posibles de columna y devuelve el primer nombre que encuentra o None si no encontró ninguna. 
	Def get_global_panalties obtiene las penalizaciones del DataFrame de parámetros globales y evuelve un diccionario. Si Excel no tiene los datos, define los valores por default para que el modelo no depende del Excel. 
	Def first_area_exchange_limits devuelve un diccionario de los límites de intercambio por área (lb, lu). 
	Def build_inflow_by_reservoir recibe el Data Frame de estación hidrológica que relaciona estación con el embalse y el Data Frame de caudales por estación hidrológica. Dvuelve el caudal recibido por embalse. 

### 2.6.Creación de diccionarios para usar en el modelo y solve:
 
Se cargan los DataFrame con nombres claros y cortos para cada uno y se crean los diccionarios y listas que se utilizan en el modelo de optimización. 

### 2.7. Modelo PYOMO:

Conjuntos -Sets: Se definen los conjuntos horas H, áreas A, flujo entre enlaces LINKS, hidroeléctricas sin embalse ROR, hidroeléctricas con embalse HRES, embalse RES térmicas commitment TC, térmicas no-commitment TNC, plantas de generación renovables eólica y solar REN y baterías BESS. En el caso de las horas se define que el conjunto tiene un orden de 1 a 24 para poder aplicar en el modelo la dinámica del embalse y de las baterías. 
 
Variables: Se definen las variables necesarias para el modelo usando el formato estándar m.nombre_de_la_variable = Var(índice, dominio). 
 
	Flow: flujo de intercambio entre cada enlace LINKS en cada hora. En este caso no se definen flujos específicos porque no se tiene la información pero esta variable sirve para definir entre que áreas existe intercambio. 
	Exch: flujo neto de intercambio del área a en la hora h
	Load_shed: Demanda no atendida del área A en la hora H
	Dump: Energía excedente que se bota en el área A en la hora H. Se utiliza para evaluar el funcionamiento del modelo
	P_ror: Potencia generada por la planta ROR en la hora H
	P_hres: Potencia generada por la planta HRES en la hora H
	P_tc: Potencia generada por la térmica TC en la hora H
	P_tnc: Potencia generada por la térmicno TNC en la hora H
	P_ren: Potencia generada por la eólica/solar en la hora H
	Curt: Curtailment de energía renovable (eólica/solar) por restricciones
	Q_turb: caudal turbinado por central hidroeléctrica HRES en la hora H
	Q_spill: caudal vertido por central hidroeléctrica HRES en la hora H
	V: Volumen almacenado en el embalse asociado a HRES en la hora H
	Slack_vmin / slack_vmax: variables artificales que relajan la restricción de volumen mínimo y máximo permitiendo violarlas en casos extremos. Esta variable se coloc+o siguiendo la recomendación de PSR (2009) para evitar infactibilidades
	 Slack_vend: Variable que penaliza bajar el embalse al final del día por debajo del nivel del embalse inicial. Esta variable permite darle un valor al recurso hídrico
	P_dis: Potencia de descarga de la batería BESS en la hora H
	P_ch: Potencia de carga de la batería BESS en la hora H
	Soc: Estado de carga de la batería BESS en la hora H

Restricciones: Se definen las restricciones necesarias para el modelo

	Límite de intercambio por área: exch[a,h] = flujos que entran - flujos que salen.
    Permite reconocer que las áreas reciben o exportan energía a áreas específicas, no necesariamente a todas las demás áreas, además de definir que si entra es positivo el flujo y si el flujo es de salida, entonces tiene un valor negativo
 
	Límite de intercambios de acuerdo con los límites de importación [UB] y exportación [LB] definidos
 
	Conservación de la energía durante el intercambio: La suma de los intercambios entre las diferentes áreas debe ser igual a cero
 
	Ecuación de balance: la generación + Importación + descarga de las baterías debe ser igual a la demanda + exportación + carga de las baterías+ demanda no atendida
 
	Límites de potencias: 
	La potencia generada por cada planta ROR debe ser menor o igual a su potencia máxima afectada por el perfil de la hora.
 	La potencia generada por HRES debe ser igual a su caudal turbinado multiplicado por el factor medio de producción. Así miso, la potencia generada debe esta entre su potencia mínima (se define cero por default) y su potencia máxima. 
 	La potencia generada por las térmicas no puede estar por debajo de su potencia mínima (cero por default ni ser superior a su potencia máxima.
	La energía generada por la planta eólica y solar + el vertimiento renovable (curtailment) debe ser igual al recurso disponible para generar.
 
	Caudales máximos y mínimos: 
	El caudal turbinado junto con el caudal vertido no puede superar el caudal máximo. 
 
	Dinámica del embalse y límites:
	El volumen del embalse al final de la hora 1 corresponde al volumen inicial más el caudal recibido menos el caudal turbinado y vertido. En el caso de las horas 2 a la 24 el volumen al final del embalse, toma el volumen de la hora anterior añadiendo el caudal            recibido y restando el caudal turbinado y vertido.  
	
	Adicional a la dinámica se define las restricciones asociadas al límite del volumen de los embalses y la penalización para evitar que el nivel de cada embalse disminuya por debajo de su nivel inicial.
 

	Dinámica del almacenamiento de las baterías y límites de carga y descarga
	La potencia cargada o descargada en cada hora no puede ser superior a su límite máximo  
	El almacenamiento de la batería al final de la hora 1 corresponde al almacenamiento inicial sumando la carga y restando la descarga de la hora. Para las horas 2 a la 24, en lugar de tomar el almacenamiento inicial se toma como punto de partida el nivel de     		almacenamiento de la hora anterior. 
 	El estado de carga de la batería debe estar entre su mínimo y máximo permitido. 
    Finalmente se define la restricción que indica que no se puede cargar la batería más  de lo que le permite el estado de carga inicial en la hora y no se puede descarga la batería más de lo que se tiene disponible que se encuentra por encima del nivel mínimo.
 
### 2.8. Función objetivo y solución

Se calcula el costo horario de la térmica en función de su heat rate y el costo del combustible que utiliza. Se define el costo que penaliza los vertimientos de las plantas hidroeléctricas y se define la función objetivo a minimizar: 
Finalmente, habiendo descrito la función objetivo a minimiza se soluciona utilizando la licencia académica de Gurobi  y especificando las variables duales que se quieren obtener. 
 
