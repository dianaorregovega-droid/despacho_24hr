"""
Despacho económico 24h (un solo día) con Pyomo, leyendo un LIBRO EXCEL (múltiples hojas).

Incluye:
- Carga Excel + limpieza columnas + conversión numérica europea
- Modelo Pyomo 24h
- 2 escenarios: No-BESS y BESS
- Manejo de infeasible (no imprime si no hay solución)
- Robustez embalses: filtra embalses fantasma + slack_vmax penalizado + saneos

Requisitos:
  pip install pandas openpyxl pyomo numpy
  + Gurobi instalado/licenciado
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np

from pyomo.environ import (
    ConcreteModel, Set, Var, Constraint, Objective, Suffix,
    NonNegativeReals, Reals, minimize, SolverFactory, value
)
from pyomo.opt import TerminationCondition


"""
from __future__ import annotations: se usa para que Python trate los tipos (diccionarios, listas, tuplas, etc.) solo como etiquetas informativas, sin intentar evaluarlos ni ejecutarlos en el momento de leer el código.
import re: se usará para buscar y limpiar texto, por ejemplo para normalizar nombres de columnas de Excel y evitar problemas por espacios invisibles o formatos raros.
Dict, List, Optional, Tuple: se usan para indicar el tipo de datos que se espera manejar:
  - Dict: diccionarios (clave → valor)
  - List: listas
  - Optional: el dato puede existir o no (None, por ejemplo precio = None o precio = 35)
  - Tuple: lista fija que no cambia y tiene un orden definido
pandas (pd): se usa para leer archivos Excel y manejar datos en forma de tablas (DataFrames).
numpy (np): se usa para operaciones numéricas, manejo de arrays, conversiones y valores nulos (NaN).
ConcreteModel: contenedor del modelo. Es donde se define todo el problema de optimización (conjuntos, variables, restricciones y función objetivo).
Set: conjuntos o índices del modelo. Ejemplos: conjunto de horas, plantas, tecnologías, áreas.
Var: variables de decisión del modelo. Ejemplos: potencia generada, carga/descarga de baterías, encendido de unidades térmicas.
Constraint: restricciones que debe cumplir el modelo. Ejemplo: balance de energía (demanda = generación + importaciones − pérdidas).
Objective: función objetivo del modelo (lo que se desea minimizar o maximizar). Ejemplo: costo total de operación.
Binary: tipo de variable binaria (0 o 1). Ejemplo: unidad térmica apagada (0) o encendida (1).
NonNegativeReals: variables reales no negativas (≥ 0). Ejemplo: potencia generada, energía almacenada.
Reals: variables reales sin restricción de signo (pueden ser positivas o negativas). Ejemplo: flujos con signo, desviaciones o intercambios netos.
minimize: indica que la función objetivo se minimiza.
SolverFactory: crea la interfaz para llamar al solver de optimización (por ejemplo: CBC, GLPK, Gurobi, CPLEX).
value: función para extraer el valor numérico de una variable o expresión de Pyomo.
TerminationCondition: se utiliza para verificar cómo terminó el solver (óptimo, infeasible/sin solución, límite de tiempo, etc.).
"""

# ============================================================
# 0) CONFIG
# ============================================================
EXCEL_PATH = r"D:\beca fabio chaparro\tesis\modelizacion\STP Y SOC\analisis energetico25-30\input_despacho24hv16.xlsx"

SHEETS_REQUIRED = [
    "Master_area",
    "Master_estacion_hidrologica",
    "Caudalesm3s",
    "Parametros_globables",
    "Hydro_ror",
    "Profile_ror_by_area",
    "Hydro_with_reservoir_base",
    "Termica_commitment_base",
    "Termica_no_commitment_base",
    "Costo_combustible",
    "Parametro_renovable",
    "Profile_renovable",
    "Limites_areas",
    "Demanda",
    "Bess_base",
    "Req_regulation",   # se carga pero NO se usa
    "Cap_renovable",    # se carga pero NO se usa
]

# ============================================================
# 1.) Limpieza / conversión europea /skip por hoja / post-proceso de tipos /carga excel
# ============================================================

# 1.1) Limpieza / conversión europea
# ============================================================

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia nombres de columnas y los normaliza a snake_case lower."""
    df = df.copy()
    cols = []
    for c in df.columns:
        c = str(c).strip() #asegura que cada columna sea texto y quita espacios al inicio y al final
        c = re.sub(r"\s+", "_", c) #reemplaza espacio por _ ej: demanda total -> demanda_total
        c = re.sub(r"[^0-9a-zA-Z_]+", "", c) #elimina todo lo que no sea letras, números y guiones bajos ej: potencia(mw) ->potencia_mw
        cols.append(c.lower()) #pasa el todo a minúsculas y lo guarda en la lista cols
    df.columns = cols #Reemplaza los nombres originales de las columnas por los nombres limpios que acaba de crear
    return df


def _to_numeric_eu(series: pd.Series) -> pd.Series:
    """
    Convierte formato EU:
    - "1.234,56" -> 1234.56
    - "112,20" -> 112.2
    """
    if pd.api.types.is_numeric_dtype(series):
        return series                                                          # Si la columa ya es númerica no hace nada y la devuelve tal cual
    s = series.astype(str).str.strip()                                         # Si la columna no es númerica convierte todo a texto, quita espacios 
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan})                 # y reemplaza valores vacios o raros por NanN
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False) # 1.234,56 -> 1234.56
    return pd.to_numeric(s, errors="coerce")                                   # Intenta convertir a número y si no puede pone NaN


def coerce_numeric_columns(df: pd.DataFrame, skip_cols: Optional[List[str]] = None) -> pd.DataFrame:
    """Convierte columnas object a numérico EU, excepto columnas en skip_cols."""
    df = df.copy()
    skip_cols = set([c.lower() for c in (skip_cols or [])])

    # skip_cols or []: si no se pasan columnas a excluir (skip_cols = None),
    # se usa una lista vacía; si se pasa una lista (por ejemplo ["ID", "Area"]),
    # se usa esa lista.
    #
    # [c.lower() for c in ...]: bucle compacto que recorre los nombres de las
    # columnas y los convierte a minúsculas para evitar problemas de mayúsculas.
    #
    # set(...): convierte la lista en un conjunto para que la comprobación
    # if c in skip_cols sea más rápida y porque el orden no importa.
    #
    # Ejemplo final: skip_cols = {"id", "area"}

    for c in df.columns:
        if c.lower() in skip_cols:                #si la columnas está en la lista de no tocar "skip_cols la salta
            continue  
        if df[c].dtype == object:                 #Si la columna es tipo texto (object) puede esconder números mal formateados entonces:
            converted = _to_numeric_eu(df[c])     # si al menos un valor se pudo convertir correctamente  
            if converted.notna().sum() > 0:       # reemplazas la columna original por la convertida
                df[c] = converted

    return df


# 1.2) Skip por hoja
# ============================================================

skip_cols_by_sheet: Dict[str, List[str]] = {
    "Master_area": ["id_area", "name_area", "type_area"],
    "Master_estacion_hidrologica": ["id_estacionhidrologica","name_estacionhidrologica", "id_embalse"],
    "Parametros_globables": ["name_parameter"],
    "caudalesm3s": ["id_estacionhidrologica"],
    "Hydro_ror": ["id_planta", "name_planta","id_area"],
    "Profile_ror_by_area": ["id_area"],
    "Hydro_with_reservoir_base": ["id_planta", "name_planta","id_area","id_estacionhidrologica", "id_embalse"],
    "Termica_commitment_base": ["id_planta", "name_planta", "id_area","fuel_code"],
    "Termica_no_commitment_base": ["id_planta", "name_planta", "id_area", "fuel_code"],
    "Costo_combustible": ["fuel_code", "fuel_name", "fuel_unit"],
    "parametro_renovable": ["id_planta", "name_planta","id_area", "tech"],
    "Profile_renovable": ["id_area","tech"],
    "Limites_areas": ["id_area","scenario"],
    "demanda": ["id_area"],
    "bess_base": ["id_bess", "name_bess","id_area"],
    "cap_renovable": ["id_planta","id_area"],
}


# 1.3) Postproceso de tipos
# ============================================================

def _postprocess_types(df: pd.DataFrame) -> pd.DataFrame: # Recibe un DataFrame y devuelve un DataFrame
    """Ajusta tipos básicos: IDs a str limpio y hora a Int64 si existe."""
    df = df.copy()

    str_cols = [
        "id_area", "name_area", "type_area", "id_estacionhidrologica","name_estacionhidrologica","id_embalse",
        "name_parameter", "id_planta", "name_planta",
        "id_hydro_ror", "id_hres",
        "id_tc", "id_tnc",
        "fuel_code",
        "id_ren", "id_planta",
        "id_bess",
        "tech", "plant_type",
        "name", "name_planta", "name_bess",
        "key"
    ]
    for col in str_cols:                                 # Recorre una por una las columnas listadas en str_cols
        if col in df.columns:                            # Si la columna realmente existe en el DataFrame
            df[col] = df[col].astype(str).str.strip()    # Convierte toda la columna a texto y quita espacios al inicio y al final en cada celda
  
    for col in ["hour", "hour_in_day", "hora", "h"]:     #revisa en todas las hojas posibles columnas que podrías significar hora y si existe esa columna
        if col in df.columns:                    
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64") # pd.to_numeric(..., errors="coerce") intenta convertir a numero y para valores raros losconvierte a NaN y .astype("Int64") convierte a entero 

    return df


# 1.4) Carga Excel
# ============================================================

def _normalize_sheet_name(s: str) -> str: #Lafunción recibe un nombre de hoja y devuelve el nombre convertido a texto, sin espacio al inicio/final en mínusculas
    return str(s).strip().lower()

def load_all_tables_from_excel(excel_path: str, sheets_required: List[str]) -> Dict[str, pd.DataFrame]: #la función recibe la ruta de excel y la lista de hojas que necesita y devuelve un diccionario
    xls = pd.ExcelFile(excel_path, engine="openpyxl") #abre el excel una sola vez no lee aun ninguna hoja solo inspecciona el archivo
    sheet_map = {_normalize_sheet_name(sh): sh for sh in xls.sheet_names} # crea un diccionario donde la clave es el nombre normalizado de la hoja y evlaor el nombre real de la hoja ej: {"demanda":"Demanda"}

    data: Dict[str, pd.DataFrame] = {} #se crea el diccionario vacio que tendrá todas las hojas limpias

    for sh_req in sheets_required:               #recorre cada una de las hojas que el modelo necesita 
        key = _normalize_sheet_name(sh_req)      #normaliza el nombre esperado 

        aliases = [key]                          #si la hoja tiene otro nombre, la acpeta igual  
        if key == "parametros_globables":
            aliases += ["parametros_globales"]
        if key == "parametros_globales":
            aliases += ["parametros_globables"]

        real_sheet = None                       # Se inicializa una variable vacia yrecorre todos los alias posibles, si exite guarda el nombre real y sale del bucle   
        for a in aliases:
            if a in sheet_map:
                real_sheet = sheet_map[a]
                break
        if real_sheet is None:                 # Si no exitse la hoja, el programa se detiene y avisa la hoja que falta   
            raise ValueError(f"Falta la hoja requerida en el Excel: {sh_req}")

        df = pd.read_excel(xls, sheet_name=real_sheet, engine="openpyxl") #se lee la hoja de excel y aplica las funciones que definimos en el punto 1-3
        df = _clean_columns(df)

        skip = skip_cols_by_sheet.get(key, [])           
        df = coerce_numeric_columns(df, skip_cols=skip)
        df = _postprocess_types(df)

        # fillna(0) para numéricas
        num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])] # identifica las columnas númericas y para ellas rellena los Nan con 0 ( si no hay dato ->0)
        if num_cols:
            df[num_cols] = df[num_cols].fillna(0)

        data[key] = df  # Guarda la hoja ya limpia y usa el nombre normalizado como clave

    return data

# ============================================================
# 2) Funciones auxiliares 
# ============================================================

#Recibe un DataFrame ylista de nombres posibles de columna y devuelve el primer nombre que encuentra o None si no encontró ninguna:

def pick_first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        cc = c.lower()           #Para cada nombre posible de columna (candidato) pasa el candidato a minusculas
        if cc in df.columns:     #Si ese nombre está dentro de las columnas reales del DataFrame
            return cc            #devuelve el primer nombre que encuentre y términa la función 
    return None                  #Si recorrió todos y ninguno existe devuelve None  

#Saca penalizaciones del DataFrame de parámetros globales y devuelve un diccionario. Si excel no tiene los datos define los valores por default para que el modelo no dependa del excel

def get_global_penalties(df_params: pd.DataFrame) -> dict: 
    defaults = dict( 
        pen_curt=0.02,
        pen_load_shed=10000.0,
        pen_dump=0,
        pen_slack_vmin=1000.0,
        pen_slack_vmax=1000.0,
    )
    if df_params is None or df_params.empty:   #Si no existe la tabla o está vacía, devuelve defaults y listo
        return defaults

    key_col = pick_first_existing(df_params, ["key", "name_parameter"]) #intenta encontrarla columna donde viene el nombre del parámetro
    val_col = pick_first_existing(df_params, ["value", "value_parameter"]) #intenta encontrar la columna donde viene el valor del parámetro
    if not key_col or not val_col:  #Si no encontró alguna de las dos columnas -> vuelve a defaults
        return defaults

    out = {} #Crea diccionario vacío donde guardará lo que encuentra en el Excel 
    for k, v in zip(df_params[key_col], df_params[val_col]): #Recorre fila a fila, emparejando k (nombre parámetro) y v (valor del parámetro). zip(...) los recorre en paralelo. 
        kk = str(k).strip().lower() #pasa el nombre a texto, quita los espacios y los pasa a minusculas 
        try:
            out[kk] = float(v)      # Intenta convertir el valor a float 
        except Exception:           # y si no puede lo ignora sin romper el programa
            pass

    for k, v in defaults.items():
        out.setdefault(k, v)         #si k no está en out, ponle v
    return out

#Recibe tabla con límites por área y lista de IDs de áreas y devuelve diccionario area ->(lb,lu)

def first_area_exchange_limits(
    df_lim: pd.DataFrame, 
    areas: List[str],
    default_lb: float = -9999.0,
    default_ub: float =  9999.0,
    verbose: bool = True
) -> Dict[str, Tuple[float, float]]:

    
    #Devuelve límites de intercambio por área: area -> (lb, ub).Si falta la hoja / está vacía / faltan columnas / falta la fila del área / lb o ub es NaN:
    #usa (default_lb, default_ub) y (si verbose=True) imprime un aviso.

    out = {a: (default_lb, default_ub) for a in areas} #Default sin límites si no hay datos
    if df_lim is None or df_lim.empty: 
         if verbose:
            print("[WARN] Limites_areas vacío o no existe. Todas las áreas quedan sin límite por default.")
         return out

    lb_col = pick_first_existing(df_lim, ["lb_mw", "lb"]) #busca como se llama la columna lb_mw o lb
    ub_col = pick_first_existing(df_lim, ["ub_mw", "ub"]) #busca como se llama la columna lu_mw o lu
    if not lb_col or not ub_col:                          #Si no las encuentra, devuelve (0,0) para todos
        if verbose:
            print("[WARN] No se encontraron columnas lb/ub en Limites_areas. Todas las áreas quedan sin límite por default.")
        return out
        
    missing_row = []
    missing_vals = []

    for a in areas:
        rows = df_lim[df_lim["id_area"].astype(str) == str(a)]
        if len(rows) == 0:
            missing_row.append(a)
            continue  # se queda con el default

        r0 = rows.iloc[0]

        lb_is_nan = pd.isna(r0[lb_col])
        ub_is_nan = pd.isna(r0[ub_col])

        lb = float(r0[lb_col]) if not lb_is_nan else default_lb
        ub = float(r0[ub_col]) if not ub_is_nan else default_ub

        if lb_is_nan or ub_is_nan:
            missing_vals.append(a)

        if lb > ub:
            lb, ub = ub, lb

        out[a] = (lb, ub)

    if verbose:
        if missing_row:
            print(f"[WARN] Áreas sin fila en Limites_areas (se usa default {default_lb},{default_ub}): {missing_row}")
        if missing_vals:
            print(f"[WARN] Áreas con lb/ub vacío (NaN) (se usa default {default_lb},{default_ub}): {missing_vals}")

    return out

#Recibe df_station (tabla que relaciona estación-embalse) y df_flow (tabla de caudales por estación) y devuelve embalse->caudal total
#IMPORTANTE!!!! En esta primera versión no se tiene en cuenta las centrales en cascada. Por ejemplo los caudales que entran a Sancarlos No depende de lo que sale de Playas

def build_inflow_by_reservoir(df_station: pd.DataFrame, df_flow: pd.DataFrame) -> Dict[str, float]:
    inflow = {}          #Crea diccionario vacío donde acumulará caudales
    if df_station is None or df_station.empty or df_flow is None or df_flow.empty:
        return inflow    #Si la tabla falta o está vacío devuelve vacío

# Busca nombres de columnas (por si varían): 
    
    st_id = pick_first_existing(df_station, ["id_estacion", "id_estacionhidrologica"])
    st_emb = pick_first_existing(df_station, ["id_embalse"])
    fl_id = pick_first_existing(df_flow, ["id_estacion", "id_estacionhidrologica"])
    fl_q  = pick_first_existing(df_flow, ["caudal_m3s", "caudalm3s", "caudal", "flow_m3s"])

    if not st_id or not st_emb or not fl_id or not fl_q:  #Si falta alguna columna necesaria o no se puede calcular, devuelve vacío
        return inflow
    """
    Une (merge) estación<-> embalse con caudal. La tabla df_station dice que estación hidrológica pertenece a que embalse y df_flow dice cual es 
    el caudal medido en cada estación. Une las filas cuando el ID de estación de la izquierda sea igual al ID de estación de la derecha”.
    Como en ambos casos es "id_estacionhidrologica", estás uniendo por esa columna.      
    """
    merged = df_station[[st_id, st_emb]].merge(   #de la tabla de estaciones quedese solo con la columna de id_estacionhidrologica y id_embalse
        df_flow[[fl_id, fl_q]],                   #de la tabla de caudales quedese solo con la columna id_estacionhidrologica y la columna del caudal
        left_on=st_id,                            #Usa esta columna de la tabla izquierda (df_station) como clave
        right_on=fl_id,                           #Usa esta columna de la tabla derecha (df_flow) como clave
        how="left"                                #Quedese con todas las filas de la tabla izquierda (df_station) aunque no exista caudal para esa estación  
    )
    merged[st_emb] = merged[st_emb].astype(str).str.strip()  # merged[st_emb] es la columna del embalse (en tu caso st_emb = "id_embalse"). .astype(str) convierte todo a texto (por si venía como número o tenía NaN).str.strip() quita espacios al inicio y al final.

    merged = merged[merged[st_emb].notna()]   #quedese solo con filas donde la columna id_embalse no es NaN real
    merged = merged[~merged[st_emb].isin(["", "nan", "None"])]  #Se queda con los que NO son ""(vacío), "nan" (lo que quedó cuando NaN se convirtipo a texto "None". Con esto se evita sumar caudales nan o vacío

    for _, r in merged.iterrows():              #recorre el DataFrame fila por fila - r es una fila (tipo Series) _ es el índice de la fila, pero no lo necesitas (por eso lo llamamos_)
        emb = str(r[st_emb]).strip()            #lee el valor del embalse de esa fila, lo convierte a texto (por seguridad) y le quitas espacios
        q = float(r[fl_q]) if pd.notna(r[fl_q]) else 0.0   #fl_q es la columna de caudal. Si el valor existe (notna), lo convierte a float y si en NaN lo pone en 0
        inflow[emb] = inflow.get(emb, 0.0) + q  #Si ya existe emb en el diccionario, dame lo que lleva y si no existe, dame 0, luego suma q y lo guarda. Si hay varias estaciones que terminan en el mismo embalse, aqui se suman

    return inflow                               #devuelve el diccionario final embalse -> caudal total


# ============================================================
# 3) CREACIÓN DE DICCIONARIOS PARA USAR EN EL MODELO + SOLVE
# ============================================================

# 3.1) Del DataFrames a Diccionarios
# ============================================================

def build_and_solve_model(data: Dict[str, pd.DataFrame], scenario_name: str, enable_bess: bool):
    # --- tablas (solo coloca nombres claros y cortos para esas hojas)
    
    df_area = data["master_area"]
    df_dem  = data["demanda"]
    df_lim  = data["limites_areas"]
    df_params = data.get("parametros_globables", pd.DataFrame()) #Intenta obtener la hoja "parametros_glboales" y si no existe devuelve un DataFrame vacio
    if df_params.empty and "parametros_globales" in data:        #Si salio vacio (no estaba/no tenia filas) y existe una hoja con nombre alternativo, usa esa hoja
        df_params = data["parametros_globales"]

    df_ror = data["hydro_ror"]                        #hidro run of river. Toma las que no tienen embalse
    df_ror_prof = data["profile_ror_by_area"]         #perfil de generacion de Run of River por area (porcentaje respecto a potencia máxima Pmax)
    df_hres = data["hydro_with_reservoir_base"]       #paramétros de las hidráulicas con embalse 
    df_station = data["master_estacion_hidrologica"]  #Asocia estación hidrológica con embalse 
    df_flow = data["caudalesm3s"]                     #Caudales promedio m3s para la semana 100 correspondiente al 15 de noviembre de 2027 

    df_tc = data["termica_commitment_base"]           #Parámetros de las térmicas commitment
    df_tnc = data["termica_no_commitment_base"]       #Parámetros de las térmicas que no participan en el despacho (tomadores de precio)
    df_fuel = data["costo_combustible"]               

    df_ren = data["parametro_renovable"]              #Parametros de las plantas solares y eólicas estimadas que estén en operación en la semana 100 (15/11/2027)
    df_ren_prof = data["profile_renovable"]           #Perfil de generación horario por área y tecnología (porcentaje respecto a Pmax)

    df_bess = data["bess_base"]                       #Parámetros Bess, se incluye el SAEB del Atlantico y los 5 SAEB propuestos en el plan maestro de transmisión de 2025 

    penalties = get_global_penalties(df_params)       #llama la función que construimos en el punto 6 y devuelve un diccionario tipo penalties["pen_curt"] penalties["pen_load_shed"] que usamos luego en la función objetivo 

    # --- sets base
    AREAS = [str(x) for x in df_area["id_area"].astype(str).tolist()]    #crea una lista de áreas y convierte todo a texto por seguridad
    H = list(range(1, 25))                            #Define el set de horas del día de 1 a 24 


    # Topología entre áreas- Enlaces NO dirigidos (i < j)
    LINKS = [("1","2"), ("1","4"), ("2","3"), ("2","4"), ("2","5"), ("3","5"), ("4","5")]

    # normaliza a tu formato de AREAS (strings)
    LINKS = [(str(i), str(j)) for (i, j) in LINKS if str(i) in AREAS and str(j) in AREAS]
    
        
    # --- demanda
    
    # Busca cual de los nombres existe para la columna de hora y para la columna de demanda y si no encuentra alguna de las dos se detiene y manda aviso:
    
    dem_hour_col = pick_first_existing(df_dem, ["hour", "hour_in_day", "hora"])
    dem_val_col  = pick_first_existing(df_dem, ["demand_mw", "demanda_mw", "demanda", "demand"])
    if not dem_hour_col or not dem_val_col:
        raise ValueError("En hoja 'Demanda' falta columna hora (hour/hour_in_day) o demanda (Demand_MW).")

    # Crea un diccionario con todas las combinaciones (area, hora) y lo inicializa todo en cero Ejemplo: demand[("norte", 5)] = 0.0. Luego rellena el 
    # diccionario recorriendo cada fila de la tabla de demanda, toma el área de esa fila, toma la hora y la convierte a entero y si esta vacía deja 
    # hh=None. Solo si el área y la hora son válidas guarda la demanda en el diccionario, si la celda de demanda está vácia pone cero. 
    # resultado: demand[(area, hora)] queda listo para usar como parámetro del modelo.
    
    demand = {(a, h): 0.0 for a in AREAS for h in H}
    for _, r in df_dem.iterrows():
        a = str(r["id_area"])
        hh = int(r[dem_hour_col]) if pd.notna(r[dem_hour_col]) else None
        if a in AREAS and hh in H:
            demand[(a, hh)] = float(r[dem_val_col]) if pd.notna(r[dem_val_col]) else 0.0

    # Límites de intercambio y caudales por embalse:
    
    exch_lim = first_area_exchange_limits(df_lim, AREAS)         #Devuelve un diccionario exch_lim["norte"] = (lb, ub)
        
    inflow_m3s = build_inflow_by_reservoir(df_station, df_flow)  #Devuelve embalse -> caudal total usando el merge  
    FLOW_TO_HM3_PER_H = 0.0036                                   #constante de conversiónpara convertir de m3s a hm3/h

    # --- IDs recursos: busca la columna que contiene el ID de cada planta ROR, hydro_with_reservoir y embalse. Si no existe aborta y avisa. El resultado es listas con los IDs de las plantas ROR, hres y embalses
    
    # ID Recursos ROR:
    ror_id_col = pick_first_existing(df_ror, ["id_hydro_ror", "id_planta"])
    if not ror_id_col:
        raise ValueError("En 'Hydro_ror' falta id_hydro_ror o id_planta.")
    ROR = [str(x) for x in df_ror[ror_id_col].astype(str).tolist()]

    # ID Recursos hres:
    hres_id_col = pick_first_existing(df_hres, ["id_hres", "id_planta"])
    if not hres_id_col:
        raise ValueError("En 'Hydro_with_reservoir_base' falta id_hres o id_planta.")
    HRES = [str(x) for x in df_hres[hres_id_col].astype(str).tolist()]

    # Lista embalses:
    emb_col = pick_first_existing(df_hres, ["id_embalse"])
    if not emb_col:
        raise ValueError("En 'Hydro_with_reservoir_base' falta id_embalse.")

    RES = sorted({                                      #ordena y con {} se crea un set con una lita ordenada de embalses únicos
        str(x).strip()                                  #limpia espacios
        for x in df_hres[emb_col].astype(str).tolist()  #saca la lista de embalses de la tabla
        if str(x).strip() not in ("", "nan", "None")    #elimina basura
    })

    #ID Recursos térmicos:
    
    tc_id_col = pick_first_existing(df_tc, ["id_tc", "id_planta"])
    tnc_id_col = pick_first_existing(df_tnc, ["id_tnc", "id_planta"])
    TC = [str(x) for x in df_tc[tc_id_col].astype(str).tolist()] if tc_id_col else []
    TNC = [str(x) for x in df_tnc[tnc_id_col].astype(str).tolist()] if tnc_id_col else []

    
     # ID recursos eólicos y solares:
    
    ren_id_col = pick_first_existing(df_ren, ["id_ren", "id_planta"])
    if not ren_id_col:
        raise ValueError("En 'Parametro_renovable' falta id_ren o id_planta.")
    REN = [str(x) for x in df_ren[ren_id_col].astype(str).tolist()]

   
    # ID recursos baterías (BESS o SAEB):
    
    bess_id_col = pick_first_existing(df_bess, ["id_bess"])
    BESS = [str(x) for x in df_bess[bess_id_col].astype(str).tolist()] if bess_id_col else []

    # --- map a áreas: Función que sirve para crear un mapa tipo area_of_planta["PLANTA_X"] = "NORTE".  Si la tabla esta vacía o no existe la columna del ID o id_area devuelve el diccionario vacío:
    
    def map_area(df: pd.DataFrame, idcol: str) -> Dict[str, str]:
        if df.empty or idcol not in df.columns or "id_area" not in df.columns:
            return {}
        return {str(r[idcol]): str(r["id_area"]) for _, r in df.iterrows()} #recorre filas y crea un diccionario (llave: id_planta, valor: id_area), todo convertido a string.

    ror_area  = map_area(df_ror, ror_id_col)
    hres_area = map_area(df_hres, hres_id_col)
    tc_area   = map_area(df_tc, tc_id_col) if tc_id_col else {}
    tnc_area  = map_area(df_tnc, tnc_id_col) if tnc_id_col else {}
    ren_area  = map_area(df_ren, ren_id_col)
    bess_area = map_area(df_bess, bess_id_col) if bess_id_col else {}

    # Filtra el DataFrame quedándose solo con las filas cuyo idcol sea igual a idx. astype(str) y str(idx) asegura que compare texto con texto e .iloc[0] toma la primera fila del resultado:
    
    def get_row(df, idcol, idx):       
        return df[df[idcol].astype(str) == str(idx)].iloc[0]

    # Es un atajo que llama a la función pick_first_existing para encontrar la primera columna que exista dentro de la lista names. 
    #Ejemplo: si buscas ["om_usd_per_mwh", "oandm_usd_per_mwh"] y solo existe om_usd_per_mwh, devuelve esa:    
    def col(df, names): return pick_first_existing(df, names)

    # --- Diccionarios parámetros ROR:
    # Crea dos diccionarios vacíos, busca que columnas contiene Pmax y O&M y para cada planta ROR busca su fila en df_ror y lee pmax y om. 
    #Si la columna no existe, usa 0.0 row-get(col,deafault) evita erro si falta la celda
    
    ror_pmax = {}; ror_om = {}                                            
    ror_pmax_col = col(df_ror, ["pmax_mw"])  
    ror_om_col   = col(df_ror, ["oandm_usd_per_mwh", "om_usd_per_mwh"])
    for r in ROR:
        row = get_row(df_ror, ror_id_col, r)
        ror_pmax[r] = float(row.get(ror_pmax_col, 0.0)) if ror_pmax_col else 0.0
        ror_om[r]   = float(row.get(ror_om_col, 0.0)) if ror_om_col else 0.0

    # --- Diccionario perfil ROR:
    #Crea un diccionario por defecto ror_shape[(area, hora)] = 1.0. De esta forma, asume "plano" (100% todo el tiempo).
    #Luego busca columnas de hora y shape (fracción) y si existen la columna, recorre fila a fila el perfil, convierte hora a int y si 
    #es pareja (area,hora), reemplazoa 1.0 por el valor real del perfl. Si el valor está vacío deja 1.0

    ror_shape = {(a, h): 1.0 for a in AREAS for h in H}
    rorph_hour = col(df_ror_prof, ["hour", "hour_in_day", "hora"])
    rorph_shape = col(df_ror_prof, ["shape_frac","profile","perfil","shape"])
    if rorph_hour and rorph_shape and "id_area" in df_ror_prof.columns:
        for _, rr in df_ror_prof.iterrows():
            a = str(rr["id_area"])
            hh = int(rr[rorph_hour]) if pd.notna(rr[rorph_hour]) else None
            if (a, hh) in ror_shape:
                ror_shape[(a, hh)] = float(rr[rorph_shape]) if pd.notna(rr[rorph_shape]) else 1.
    
    # --- Diccionario parámetros HRES:
    """
    Busca los nombres de las columnas detectando que colummnas usar para potencias, caudales min/max, MW/m3s), volumenes min/max/inicial y O&M.
    Luego crea los diccionarios que guaradapra el parámetro por planta: hres_emb[g] = embalse, hres_pmax[g] = pmax...
    Para cada planta con embalse, busca la fila, guarda emablse y parámetros. Si falta el valor coloca 0.0 y en el caso de o&m (cuando hay columna de O&M y el valor no es NaN) asume 0.01
    """
    
    h_pmax = col(df_hres, ["pmax_mw"])
    h_pmin = col(df_hres, ["pmin_mw"])
    h_qmax = col(df_hres, ["qmax_m3s"])
    h_qmin = col(df_hres, ["qmin_m3s"])
    h_prod = col(df_hres, ["prod_mw_per_m3s"])
    h_vmin = col(df_hres, ["vmin_hm3"])
    h_vmax = col(df_hres, ["vmax_hm3"])
    h_vini = col(df_hres, ["vinic_hm3", "vinic"])
    h_om   = col(df_hres, ["oandm_usd_per_mwh", "om_usd_per_mwh"])

    hres_emb = {}
    hres_pmax = {}; hres_pmin = {}; hres_qmax = {}; hres_qmin = {}; hres_prod = {}
    hres_vmin = {}; hres_vmax = {}; hres_vinic = {}; hres_om = {}

    for g in HRES:
        row = get_row(df_hres, hres_id_col, g)
        emb = str(row.get(emb_col, "")).strip()
        hres_emb[g] = emb
        hres_pmax[g] = float(row.get(h_pmax, 0.0)) if h_pmax else 0.0
        hres_pmin[g] = float(row.get(h_pmin, 0.0)) if h_pmin else 0.0
        hres_qmax[g] = float(row.get(h_qmax, 0.0)) if h_qmax else 0.0
        hres_qmin[g] = float(row.get(h_qmin, 0.0)) if h_qmin else 0.0
        hres_prod[g] = float(row.get(h_prod, 0.0)) if h_prod else 0.0
        hres_vmin[g] = float(row.get(h_vmin, 0.0)) if h_vmin else 0.0
        hres_vmax[g] = float(row.get(h_vmax, 0.0)) if h_vmax else 0.0
        hres_vinic[g] = float(row.get(h_vini, 0.0)) if h_vini else 0.0
        # suposición: O&M si falta
        if h_om and pd.notna(row.get(h_om, np.nan)):
            hres_om[g] = float(row.get(h_om, 0.01))
        else:
            hres_om[g] = 0.01


    # Diccionario caudales por embalse:

    res_inflow_m3s = {r: float(inflow_m3s.get(r, 0.0)) for r in RES} # Crea un diccionario completo para cada embalse. Si no aparece en inflow_m3s, lo pone en 0.0

    # saneos anti-infeasible

    #1.) qmax >= qmin y pmax >=pmin
    
    for g in HRES:
        if hres_qmax[g] < hres_qmin[g]:
            hres_qmax[g] = hres_qmin[g]
        if hres_pmax[g] < hres_pmin[g]:
            hres_pmax[g] = hres_pmin[g]

    #2.) Para cada embalse, busca plantas que estén asociados a ese embalse, toma la primera (g0) y fuerza que vmax>=vmin y vmax>=vini. Esto evita infeasible por tener vini fuera del rango 

    for r in RES:
        gens = [g for g in HRES if hres_emb.get(g) == r]
        if not gens:
            continue
        g0 = gens[0]
        hres_vmax[g0] = max(hres_vmax.get(g0, 0.0), hres_vmin.get(g0, 0.0), hres_vinic.get(g0, 0.0))


    # --- Diccionario costo de combustible:
    """
    Detecta columnas y crea un diccionario vacío, luego busca cada combustible y le asigna su costo.
    """

    fuel_id_col = col(df_fuel, ["id_fuel", "fuel_code",])
    fuel_cost_col = col(df_fuel, ["fuelcost_usd_per_mmbtu", "fuel_cost_usd_per_mmbtu", "cost_usd_per_mmbtu","fuel_cost_usd_per_fuel_unit"])
    fuel_cost = {}
    if fuel_id_col and fuel_cost_col:
        for _, rr in df_fuel.iterrows():
            fuel_cost[str(rr[fuel_id_col])] = float(rr[fuel_cost_col]) if pd.notna(rr[fuel_cost_col]) else 0.0
    else:
        print(f"[WARN] No pude detectar columnas fuel_id_col/fuel_cost_col en Costo_combustible. "
        f"fuel_id_col={fuel_id_col}, fuel_cost_col={fuel_cost_col}")

    # --- Diccionarios parámetros térmicas

    #Función que recibe un df térmico y devuelve 5 diccionarios: Pmax, Pmin, HeRate, OM, FuelID y . Si df está vacío, devuelve 5 diccionarios vacíos
    #luego llena diccionarios por cada planta y Pmax >= Pmin. 

    def therm_dicts(df, idcol):
        if df.empty or not idcol:
            return {}, {}, {}, {}, {}
        pmax = col(df, ["pmax_mw"])
        pmin = col(df, ["pmin_mw"])
        hr   = col(df, ["heatrate_mmbtu_per_mwh", "heat_rate_mmbtu_per_mwh","heat_rate_fuelMBTU_per_mwh"])
        om   = col(df, ["oandm_usd_per_mwh", "om_usd_per_mwh","om_var_usd_per_mwh"])
        fid  = col(df, ["id_fuel", "fuel_code"])
        Pmax = {}; Pmin = {}; HR = {}; OM = {}; FID = {}
        for g in df[idcol].astype(str).tolist():
            row = get_row(df, idcol, g)
            gg = str(g)
            Pmax[gg] = float(row.get(pmax, 0.0)) if pmax else 0.0
            Pmin[gg] = float(row.get(pmin, 0.0)) if pmin else 0.0
            HR[gg]   = float(row.get(hr, 0.0)) if hr else 0.0
            OM[gg]   = float(row.get(om, 0.0)) if om else 0.0
            FID[gg]  = str(row.get(fid, "")) if fid else ""
            if Pmax[gg] < Pmin[gg]:
                Pmax[gg] = Pmin[gg]
        return Pmax, Pmin, HR, OM, FID

    # Finalmente, esos diccionarios se usa para TC y TNC haciendo desempaquetado de tuplas.Ej: 

    
    tc_pmax, tc_pmin, tc_hr, tc_om, tc_fid = therm_dicts(df_tc, tc_id_col)
    tnc_pmax, tnc_pmin, tnc_hr, tnc_om, tnc_fid = therm_dicts(df_tnc, tnc_id_col)

    
    # --- Diccionarios parámetros renovables

    #Detecta columnas de pmax, om y tech, luego crea diccionarios y los llena para cada planta REN:

    ren_pmax_col = col(df_ren, ["pmax_mw"])
    ren_om_col = col(df_ren, ["oandm_usd_per_mwh", "om_usd_per_mwh"])
    ren_tech_col = col(df_ren, ["tech"])

    ren_pmax = {}; ren_om = {}; ren_tech = {}
    for g in REN:
        row = get_row(df_ren, ren_id_col, g)
        ren_pmax[g] = float(row.get(ren_pmax_col, 0.0)) if ren_pmax_col else 0.0
        ren_om[g]   = float(row.get(ren_om_col, 0.0)) if ren_om_col else 0.0
        ren_tech[g] = str(row.get(ren_tech_col, "")) if ren_tech_col else ""

     # --- Diccionarios perfil de generación renovable respecto a potencia máxima: se tienen vrios casos de forma que este código pueda servir en caso de que se tenga información por planta o se quiera aplicar un perfil por área (como esta en este momento) o únicamente por tecnología.
    
    pr_hour  = col(df_ren_prof, ["hour", "hour_in_day", "hora"])
    pr_shape = col(df_ren_prof, ["shape_frac", "profile"])   
    pr_id    = col(df_ren_prof, ["id_ren", "id_planta"])
    pr_tech  = col(df_ren_prof, ["tech"])
    
        # Caso 1: ren_shape por planta-hora
    ren_shape = {(g, h): 0.0 for g in REN for h in H}
    
    if pr_hour and pr_shape:
    
        # Caso 1: perfil viene por planta (id_ren/id_planta + hour)
        if pr_id and pr_id in df_ren_prof.columns:
            for _, rr in df_ren_prof.iterrows():
                gg = str(rr[pr_id]).strip()
                hh = int(rr[pr_hour]) if pd.notna(rr[pr_hour]) else None
                if (gg, hh) in ren_shape:
                    ren_shape[(gg, hh)] = float(rr[pr_shape]) if pd.notna(rr[pr_shape]) else 0.0
    
        # Caso 2: perfil viene por (id_area, tech, hour)  
        else:
            if "id_area" in df_ren_prof.columns and pr_tech and pr_tech in df_ren_prof.columns:
    
                # 1) Construir diccionario perfil_by_area_tech[(area, tech, hour)] = shape
                profile_by_area_tech = {}
                for _, rr in df_ren_prof.iterrows():
                    a = str(rr["id_area"]).strip()
                    t = str(rr[pr_tech]).strip()
                    hh = int(rr[pr_hour]) if pd.notna(rr[pr_hour]) else None
                    if hh in H:
                        profile_by_area_tech[(a, t, hh)] = float(rr[pr_shape]) if pd.notna(rr[pr_shape]) else 0.0
    
                # 2) Aplicar a cada planta usando SU area y SU tech
                for g in REN:
                    a = str(ren_area.get(g, "")).strip()
                    t = str(ren_tech.get(g, "")).strip()
                    if a == "" or t == "":
                        continue
                    for hh in H:
                        ren_shape[(g, hh)] = profile_by_area_tech.get((a, t, hh), profile_by_area_tech.get(("0", t, hh), 0.0))
                        # Se incluye fallback A (0", tech, hour) para perfil global ya que en la eólica manejamos un unico perfil para todas las áreas

    
            # Caso 3: perfil por tech único (sin id_area) 
            elif pr_tech and pr_tech in df_ren_prof.columns:
                profile_by_tech = {}
                for _, rr in df_ren_prof.iterrows():
                    t = str(rr[pr_tech]).strip()
                    hh = int(rr[pr_hour]) if pd.notna(rr[pr_hour]) else None
                    if hh in H:
                        profile_by_tech[(t, hh)] = float(rr[pr_shape]) if pd.notna(rr[pr_shape]) else 0.0
    
                for g in REN:
                    t = str(ren_tech.get(g, "")).strip()
                    for hh in H:
                        ren_shape[(g, hh)] = profile_by_tech.get((t, hh), 0.0)


    # --- Diccionarios de baterías - bess
    if df_bess.empty or not bess_id_col: #Si no existe la hoja o no esta la columna ID, no hay baterías
        BESS = []

    # Busca las columnas:
    
    b_pdis = col(df_bess, ["p_dis_max_mw"])
    b_pch  = col(df_bess, ["p_ch_max_mw"])
    b_emax = col(df_bess, ["e_max_mwh"])
    b_effch = col(df_bess, ["eff_ch", "eta_ch"])
    b_effdis = col(df_bess, ["eff_dis", "eta_dis"])
    b_socini = col(df_bess, ["soc_ini_frac", "soc0_frac"])
    b_socmin = col(df_bess, ["soc_min_frac"])
    b_socmax = col(df_bess, ["soc_max_frac"])

    # Crea diccionarios
    bess_pdis = {}; bess_pch = {}; bess_emax = {}
    bess_effch = {}; bess_effdis = {}
    bess_socini = {}; bess_socmin = {}; bess_socmax = {}

    #Llena los diccionarios por batería:

    for b in BESS:
        row = get_row(df_bess, bess_id_col, b)
        bess_pdis[b] = float(row.get(b_pdis, 0.0)) if b_pdis else 0.0
        bess_pch[b]  = float(row.get(b_pch, 0.0)) if b_pch else 0.0
        bess_emax[b] = float(row.get(b_emax, 0.0)) if b_emax else 0.0
        bess_effch[b] = float(row.get(b_effch, 1.0)) if b_effch else 1.0
        bess_effdis[b] = float(row.get(b_effdis, 1.0)) if b_effdis else 1.0
        bess_socini[b] = float(row.get(b_socini, 0.5)) if b_socini else 0.2
        bess_socmin[b] = float(row.get(b_socmin, 0.0)) if b_socmin else 0.0
        bess_socmax[b] = float(row.get(b_socmax, 1.0)) if b_socmax else 1.0

    if not enable_bess: #esto apaga las baterías
        for b in BESS:
            bess_pdis[b] = 0.0
            bess_pch[b]  = 0.0 
   
    # =========================
    # 4.) PYOMO MODEL
    # =========================
    m = ConcreteModel(name=f"dispatch_24h_{scenario_name}") #Crea un modelo Pyomo concreto (osea ya con datos) y le pone un nombre. Este es el contenedor
    m.dual = Suffix(direction=Suffix.IMPORT)
        
    # 4.1) CONJUNTOS - SETS
    # ========================= 
    
    m.H = Set(initialize=H, ordered=True) # Conjunto de horas ordered=true hace que tenga un orden de 1 a 24 ya que no es lo mismo 1,2,3 que 3,2,1
    m.A = Set(initialize=AREAS)
    m.LINKS = Set(initialize=LINKS, dimen=2) #flujo por enlace (positivo i+j)
    m.ROR = Set(initialize=ROR)
    m.HRES = Set(initialize=HRES)
    m.RES = Set(initialize=RES)
    m.TC = Set(initialize=TC)
    m.TNC = Set(initialize=TNC)
    m.REN = Set(initialize=REN)
    m.BESS = Set(initialize=BESS)

    # 4.2) VARIABLES
    # ========================= 

    """
    la forma general es: m.nombre = Var(índices, domain=DOMINIO): Defino una variable nombre, definida para cada combinación de los índices, y que 
    vive en cierto dominio matemática. m es el contenedor del modelo que definimos en m = ConcreteModel de forma que todo lo que cuelga de m
    forma parte del problema ya sea como sts, variables, restricciones o función objetivo. Es decir, m es como el cuaderno donde escribes el modelo matemático
    """
    m.flow = Var(m.LINKS, m.H, domain=Reals)    
    m.exch = Var(m.A, m.H, domain=Reals) # Variable para cada par (área, hora=) el flujo neto de intercambio del área a en la hora h
    m.load_shed = Var(m.A, m.H, domain=NonNegativeReals) # energia no servida del área por hora. No puede ser negativa y solo existe si el modelo no logra cubrir demanda
    m.dump = Var(m.A, m.H, domain=NonNegativeReals) #Energia excedente que se bota en un área en cada hora cuando hay sobreoferta para evitar infactbile. Se penaliza en la función objetivo 

    m.p_ror = Var(m.ROR, m.H, domain=NonNegativeReals) # potencia generada por cada planta ROR en cada hora. No puede ser negativa
    m.p_hres = Var(m.HRES, m.H, domain=NonNegativeReals) #potencia generada por cada planta HRES en cada hora
    m.p_tc = Var(m.TC, m.H, domain=NonNegativeReals) # potencia generada por cada planta termica commitment en cada hora
    m.p_tnc = Var(m.TNC, m.H, domain=NonNegativeReals) # potencia generada por cada planta termica no-commitment en cada hora

    m.p_ren = Var(m.REN, m.H, domain=NonNegativeReals) # potencia generada por cada planta renovable (eólica/solar) en cada hora
    m.curt = Var(m.REN, m.H, domain=NonNegativeReals) # curtailment de energia renovable (eolica/solar) por restricciones 

    m.q_turb = Var(m.HRES, m.H, domain=NonNegativeReals) #caudal turbinado por central hidroeléctrica con embalse en cada hora
    m.q_spill = Var(m.HRES, m.H, domain=NonNegativeReals) # caudal vertido por central hidroeléctrica con embalse en cada hora
    m.v = Var(m.RES, m.H, domain=NonNegativeReals) # volumen almacenado. Se conecta entre horas con una ecuación dinámica
    m.slack_vmin = Var(m.RES, m.H, domain=NonNegativeReals) # Variable artifical que relaja la restriccion de volumen mínimo del embalse cuando no puede cumplirse. Esto para evitar infactibles y tiene penalizacion alta.
    m.slack_vmax = Var(m.RES, m.H, domain=NonNegativeReals) # variable artificial que permite violar el volumen máximo del embalse si es estrictamente necesario
    m.slack_vend = Var(m.RES, domain=NonNegativeReals)  # penaliza bajar al final del día
    
    m.p_dis = Var(m.BESS, m.H, domain=NonNegativeReals) #potencia de descarga por cada cada batería en  cada hora 
    m.p_ch = Var(m.BESS, m.H, domain=NonNegativeReals) #potencia de carga por cada batería en  cada hora 
    m.soc = Var(m.BESS, m.H, domain=NonNegativeReals) #estado de carga de cada batería en cada hora


    # 4.3) RESTRICCIONES
    # ========================= 

    #4.3.1. Límites de intercambio por área y por horas:

    # DEFINICIÓN DEL INTERCAMBIO NETO POR ÁREA --> exch[a,h] = flujos que entran - flujos que salen. ESte es importante porque  permite reconcoer que las areas reciben o exportan energia a areas especificas, no necesariamente a todas las demás áreas
    def exch_def_rule(mdl, a, h):
        expr = 0.0
        for (i, j) in mdl.LINKS:
            if a == i:
                expr -= mdl.flow[i, j, h]   # sale de a
            elif a == j:
                expr += mdl.flow[i, j, h]   # entra a a
        return mdl.exch[a, h] == expr
    
    m.ExchDef = Constraint(m.A, m.H, rule=exch_def_rule)

    # Límite superior: exch[a,h] <= ub

    # busca los límites de intercambio de esa área en el diccionario exch_lim y si no existe el área pone intercambio forzado a cero. la restricción se divide en dos partes: 
    
    def exch_ub_rule(mdl, a, h):
        lb, ub = exch_lim.get(a, (0.0, 0.0))
        return mdl.exch[a, h] <= ub
    m.ExchUB = Constraint(m.A, m.H, rule=exch_ub_rule)
    
    # Límite inferior: exch[a,h] >= lb
    def exch_lb_rule(mdl, a, h):
        lb, ub = exch_lim.get(a, (0.0, 0.0))
        return mdl.exch[a, h] >= lb
    m.ExchLB = Constraint(m.A, m.H, rule=exch_lb_rule)

    
    # 4.3.2.) Conservación del intercambio (suma cero): Para cada hora suma todos los intercambios de todas las áreas en esa hora y obliga a que sea cero
    def exch_zero(mdl, h):  
        return sum(mdl.exch[a, h] for a in mdl.A) == 0
    m.ExchZero = Constraint(m.H, rule=exch_zero)

    #4.3.3.) Ecuación de balance:
    # (+)La suma de la potencia de todas las plantas ROR, HRES, TC, TNC Y REN en el área a en la hora h 
    # (+) la suma de descarga de baterías en esa área 
    # (-) la suma de carga de baterías en esa área
    # (+) demanda no servida 
    # (-) lo que sobró de energía (energía botada) -> para que el modelo no sea infactible
    # (+) Importación de energía de otras áreas - Si el área exporta se restaría en la ecuación dicha exportación
    
    def balance_rule(mdl, a, h):
        gen_ror = sum(mdl.p_ror[r, h] for r in mdl.ROR if ror_area.get(r) == a) 
        gen_hres = sum(mdl.p_hres[g, h] for g in mdl.HRES if hres_area.get(g) == a)
        gen_tc = sum(mdl.p_tc[g, h] for g in mdl.TC if tc_area.get(g) == a)
        gen_tnc = sum(mdl.p_tnc[g, h] for g in mdl.TNC if tnc_area.get(g) == a)
        gen_ren = sum(mdl.p_ren[g, h] for g in mdl.REN if ren_area.get(g) == a)

        bess_dis = sum(mdl.p_dis[b, h] for b in mdl.BESS if bess_area.get(b) == a)
        bess_ch  = sum(mdl.p_ch[b, h] for b in mdl.BESS if bess_area.get(b) == a)

        return (
            gen_ror + gen_hres + gen_tc + gen_tnc + gen_ren
            + bess_dis - bess_ch
            + mdl.load_shed[a, h] - mdl.dump[a, h]     #LOS DEJO MIENTRAS REVISO QUE EL MODELO FUNCIONE
            - demand.get((a, h), 0.0)
            + mdl.exch[a, h]
            == 0
        )
    m.Balance = Constraint(m.A, m.H, rule=balance_rule)

    # 4.3.4.) Límite de ROR por Pmax y perfil por área-hora: la potencia generada por cada planta ROR en la hora h debe ser <= a Pmax*perfil horario 

    def ror_limit(mdl, r, h):
        a = ror_area.get(r, None)
        return mdl.p_ror[r, h] <= ror_pmax.get(r, 0.0) * ror_shape.get((a, h), 1.0)
    m.RORLim = Constraint(m.ROR, m.H, rule=ror_limit)

    # 4.3.5.) Relación potencia-caudal en la hidro con embalse: 
    #Para cada hidro con embalse y hora, la potencia generada debe ser ingual al caudal turbinado* su factor de producción MW/m3s. Si no existe MW/m3s, usa 0. 
    
    def hres_power_flow(mdl, g, h):
        return mdl.p_hres[g, h] == hres_prod.get(g, 0.0) * mdl.q_turb[g, h]
    m.HRESPowerFlow = Constraint(m.HRES, m.H, rule=hres_power_flow) # Pyomo recorre todas las plantas HRES y todas las horas h y para cada par (g,h) crea una ecuación usando hres_power_flow

    # 4.3.6.) Pmin/Pmax de las hidroeléctricas con embalse: 
    
    def hres_pmax_rule(mdl, g, h):
        return mdl.p_hres[g, h] <= hres_pmax.get(g, 0.0) #como P_hres es no negativa se cumple que pmin<=p_hres<=pmax
    m.HRESPmax = Constraint(m.HRES, m.H, rule=hres_pmax_rule) # Pyomo crea una restricción por cada (g,h), no asigna valores, solo impone que cualquier solución factible cumpla la desigualdad
   

    # 4.3.7.) Caudales máximos y mínimos: 
    
    def hres_qmax_rule(mdl, g, h):
        return mdl.q_turb[g, h] + mdl.q_spill[g, h] <= hres_qmax.get(g, 0.0) # caudal turbinado + vertimiento <= qmax
    
    m.HRESQmax = Constraint(m.HRES, m.H, rule=hres_qmax_rule)
    

    #4.3.8.) Dinámica del embalse (volumen hora a hora): 
    """
    Para la hora 1: v[r,1] = vinic + (inflow - outflow) * factor de conversión hm3/h
    Para las horas 2 - 24: v [r,h] = v[r, h-1] + (inflow - outflow) * factor de conversión hm3/h
    """
    
    def res_dyn(mdl, r, h):
        inflow = res_inflow_m3s.get(str(r), 0.0)
        outflow = sum(
            (mdl.q_turb[g, h] + mdl.q_spill[g, h]) 
            for g in mdl.HRES 
            if hres_emb.get(str(g), None) == str(r)
        )
        
        gens = [g for g in mdl.HRES if hres_emb.get(str(g), None) == str(r)]
        vinic = hres_vinic.get(str(gens[0]), 0.0) if gens else 0.0

        if h == 1:
            return mdl.v[r, h] == vinic + (inflow - outflow) * FLOW_TO_HM3_PER_H
        return mdl.v[r, h] == mdl.v[r, h-1] + (inflow - outflow) * FLOW_TO_HM3_PER_H
    m.ResDyn = Constraint(m.RES, m.H, rule=res_dyn)

    #4.3.9.) Límites de volumen del embalse (con slacks):
    
    def res_vmax(mdl, r, h):
        gens = [g for g in mdl.HRES if hres_emb.get(str(g), "") == str(r)]
        vmax = hres_vmax.get(str(gens[0]), 0.0) if gens else 0.0
        return mdl.v[r, h] <= vmax + mdl.slack_vmax[r, h]
    
    def res_vmin(mdl, r, h):
        gens = [g for g in mdl.HRES if hres_emb.get(str(g), "") == str(r)]
        vmin = hres_vmin.get(str(gens[0]), 0.0) if gens else 0.0
        return mdl.v[r, h] + mdl.slack_vmin[r, h] >= vmin

    m.ResVmax = Constraint(m.RES, m.H, rule=res_vmax)
    m.ResVmin = Constraint(m.RES, m.H, rule=res_vmin)

    # Slack si volumen final queda por debajo del inicial ---
    def vend_def(mdl, r):
        gens = [g for g in mdl.HRES if hres_emb.get(str(g), "") == str(r)]
        vinic = hres_vinic.get(str(gens[0]), 0.0) if gens else 0.0
        last_h = max(list(mdl.H))
        return mdl.slack_vend[r] >= vinic - mdl.v[r, last_h]
    
    m.VendDef = Constraint(m.RES, rule=vend_def)

    #4.3.10.) Límites de térmicas (Para commitment: pmin<=p<=pmax asociado a ON/OFF para no commitment no se asocia a ON/OFF):
    
    def tc_pmax_rule(mdl, g, h):
        return mdl.p_tc[g, h] <= tc_pmax.get(g, 0.0)  # * mdl.tc_on[g, h] para que pueda estar apagada o encendida

    def tc_pmin_rule(mdl, g, h):
        return mdl.p_tc[g, h] >= tc_pmin.get(g, 0.0) #* mdl.tc_on[g, h] para que pueda estar apagada o encendida

    m.TCPmax = Constraint(m.TC, m.H, rule=tc_pmax_rule)
    m.TCPmin = Constraint(m.TC, m.H, rule=tc_pmin_rule)
    
    def tnc_lim(mdl, g, h):
        return (tnc_pmin.get(g, 0.0), mdl.p_tnc[g, h], tnc_pmax.get(g, 0.0))
    m.TNCLim = Constraint(m.TNC, m.H, rule=tnc_lim)
        
    #4.3.11.) Repartición de lo disponible para generar renovable entre: potencia generada y curtailment:
    
    def ren_split(mdl, g, h):
        avail = ren_pmax.get(g, 0.0) * ren_shape.get((g, h), 0.0)
        return mdl.p_ren[g, h] + mdl.curt[g, h] == avail
    m.RenSplit = Constraint(m.REN, m.H, rule=ren_split)

    # 4.3.12) BESS: límites simples (LP) SIN binaria
    
    def bess_ch_ub(mdl, b, h):
        return mdl.p_ch[b, h] <= bess_pch.get(b, 0.0) # lo que carga debe ser menor o igual a la potencia de carga
    
    def bess_dis_ub(mdl, b, h):
        return mdl.p_dis[b, h] <= bess_pdis.get(b, 0.0)
    
    m.BESSChUB = Constraint(m.BESS, m.H, rule=bess_ch_ub)
    m.BESSDisUB = Constraint(m.BESS, m.H, rule=bess_dis_ub)


    # 4.4.13.) Ecuación dinámica del estado de carga de la batería: 
    
    def bess_soc_dyn(mdl, b, h):
        emax = bess_emax.get(b, 0.0)
        eff_ch = bess_effch.get(b, 1.0)
        eff_dis = max(bess_effdis.get(b, 1.0), 1e-6)
        soc0 = bess_socini.get(b, 0.5) * emax
        if h == 1:
            return mdl.soc[b, h] == soc0 + eff_ch * mdl.p_ch[b, h] - (1.0 / eff_dis) * mdl.p_dis[b, h]
        return mdl.soc[b, h] == mdl.soc[b, h-1] + eff_ch * mdl.p_ch[b, h] - (1.0 / eff_dis) * mdl.p_dis[b, h]
    m.BESSSOCDyn = Constraint(m.BESS, m.H, rule=bess_soc_dyn)

    # 4.4.14.) Límite de SOC mínimo y máximo:
    
    def bess_soc_bounds(mdl, b, h):
        emax = bess_emax.get(b, 0.0)
        return (bess_socmin.get(b, 0.0) * emax, mdl.soc[b, h], bess_socmax.get(b, 1.0) * emax)
    m.BESSSOCBounds = Constraint(m.BESS, m.H, rule=bess_soc_bounds)

    # 4.4.15.) (LP) Límite energético de carga: no puedo meter más de lo que me cabe en SOC
    def bess_ch_energy_limit(mdl, b, h):
        emax = bess_emax.get(b, 0.0)
        eff_ch = bess_effch.get(b, 1.0)
        soc_max = bess_socmax.get(b, 1.0) * emax
    
        # SOC "antes" de decidir en h (h=1 usa soc0, si no usa soc[h-1])
        soc_prev = (bess_socini.get(b, 0.5) * emax) if h == 1 else mdl.soc[b, h-1]
    
        # p_ch <= (soc_max - soc_prev)/eff_ch
        # (si eff_ch=0 por error, protegemos)
        return mdl.p_ch[b, h] <= (soc_max - soc_prev) / max(eff_ch, 1e-6)
    
    m.BESSChEnergy = Constraint(m.BESS, m.H, rule=bess_ch_energy_limit)


# (LP) Límite energético de descarga: no puedo sacar más de lo que tengo por encima del mínimo
    def bess_dis_energy_limit(mdl, b, h):
        emax = bess_emax.get(b, 0.0)
        eff_dis = max(bess_effdis.get(b, 1.0), 1e-6)
        soc_min = bess_socmin.get(b, 0.0) * emax
    
        soc_prev = (bess_socini.get(b, 0.5) * emax) if h == 1 else mdl.soc[b, h-1]
    
        # p_dis <= (soc_prev - soc_min)*eff_dis
        return mdl.p_dis[b, h] <= (soc_prev - soc_min) * eff_dis

    m.BESSDisEnergy = Constraint(m.BESS, m.H, rule=bess_dis_energy_limit)


    # 4.4) FUNCIÓN OBJETIVO
    # ========================= 

    # 4.4.1.) Función auxiliar para calcular el costo horario de la térmica= costo de combustible * Heat rate + O&M por MWh * lapotencia generada en la hora.
    
    def therm_cost(p_mw, hr, fuel_id, om):
        fc = fuel_cost.get(fuel_id, 0.0) # si no encuentra costo de combustible lo asume cero
        return (hr * fc + om) * p_mw

    # 4.4.2.) Penalización por vertimiento (USD/hm3):
    
    spill_cost_usd_per_hm3 = 100.0

    # 4.4.3.) Función objetivo
    
    def obj_rule(mdl):       # define la función objetivo del modelo (mdl es el modelo d m)
        total = 0.0          # Inicializa acumulador del costo total
   
        for g in mdl.TC:     # recorre cada térmica commitment y obtiene su fuel_code desde el diccionario tc_fid. Si no existe ""
            fid = tc_fid.get(g, "")
            for h in mdl.H:  # Recorre cada hora y suma al total el costo de esa planta y esa hora 
                total += therm_cost(mdl.p_tc[g, h], tc_hr.get(g, 0.0), fid, tc_om.get(g, 0.0))
        for g in mdl.TNC:
            fid = tnc_fid.get(g, "")
            for h in mdl.H:
                total += therm_cost(mdl.p_tnc[g, h], tnc_hr.get(g, 0.0), fid, tnc_om.get(g, 0.0))

        for r in mdl.ROR:   #Recorre cada planta ROR y cada hora y suma al total el costo de esa planta ROR y esa hora (O&M[USD/MWh]* generación)
            for h in mdl.H:
                total += ror_om.get(r, 0.0) * mdl.p_ror[r, h]

        for g in mdl.HRES:  # Recorre cada hidro con embalse y hora y suma al total el costo variable de esa planta y esa hora (O&M* generación [MW])
            for h in mdl.H:
                total += hres_om.get(g, 0.01) * mdl.p_hres[g, h]
                total += spill_cost_usd_per_hm3 * (mdl.q_spill[g, h] * FLOW_TO_HM3_PER_H) 

                #sobre el costo de spill:  (mdl.q_spill[g, h] en m3s * FLOW_TO_HM3_PER_H)*spill_cost_usd_per_hm3 = USD/h

        for g in mdl.REN: # suma el total de costo O&M* Gneracion MW y penalización de curtailment* curtailment (MW) 
            for h in mdl.H:
                total += ren_om.get(g, 0.0) * mdl.p_ren[g, h]
                total += penalties["pen_curt"] * mdl.curt[g, h]

        # --- Costo por throughput BESS (para poder reemplazar el hecho de no usar binarias de forma qu ela batería no cargue y descargue al mismo tiempo)
        c_deg = 0  # USD/MWh 
        for b in mdl.BESS:
            for h in mdl.H:
                total += c_deg * (mdl.p_ch[b, h] + mdl.p_dis[b, h])

        
        for a in mdl.A:
            for h in mdl.H:
                total += penalties["pen_load_shed"] * mdl.load_shed[a, h]
                total += penalties["pen_dump"] * mdl.dump[a, h]

        for r in mdl.RES:
            for h in mdl.H:
                total += penalties["pen_slack_vmin"] * mdl.slack_vmin[r, h]
                total += penalties["pen_slack_vmax"] * mdl.slack_vmax[r, h]

        pen_vend = 5000.0  # penalizacion por bajar el embalse al final del día 
        for r in mdl.RES:
            total += pen_vend * mdl.slack_vend[r]


        return total

    m.Obj = Objective(rule=obj_rule, sense=minimize)

    solver = SolverFactory("gurobi")
    if not solver.available(False):
        raise RuntimeError("No encuentro GUROBI disponible.")

    solver.options["TimeLimit"] = 3600
    solver.options["Threads"] = 0

    results = solver.solve(m, tee=False, suffixes=["dual"])

    lmp = {(a, h): float(m.dual.get(m.Balance[a, h], np.nan)) for a in m.A for h in m.H}
    dual_exch_ub = {(a,h): float(m.dual.get(m.ExchUB[a,h], np.nan)) for a in m.A for h in m.H}
    dual_exch_lb = {(a,h): float(m.dual.get(m.ExchLB[a,h], np.nan)) for a in m.A for h in m.H}

    # ============================================================
    # GUARDAR MAPAS Y PARÁMETROS EN EL MODELO (para reportes Excel)
    # ============================================================
    m._report = {}
    m._report["demand"] = demand
    m._report["ror_area"] = ror_area
    m._report["hres_area"] = hres_area
    m._report["tc_area"] = tc_area
    m._report["tnc_area"] = tnc_area
    m._report["ren_area"] = ren_area

    m._report["fuel_cost"] = fuel_cost
    m._report["tc_hr"] = tc_hr
    m._report["tc_om"] = tc_om
    m._report["tc_fid"] = tc_fid
    # m._report["tc_startup_kusd"] = tc_startup_kusd

    m._report["FLOW_TO_HM3_PER_H"] = FLOW_TO_HM3_PER_H
    m._report["spill_cost_usd_per_hm3"] = spill_cost_usd_per_hm3
    m._report["exch_lim"] = exch_lim
    m._report["AREAS"] = AREAS   # opcional, pero útil para checks

    # Para diagnóstico de curtailment:
    
    m._report["ren_pmax"] = ren_pmax
    m._report["ren_shape"] = ren_shape

    m._report["tc_pmin"] = tc_pmin
    m._report["hres_pmin"] = hres_pmin
    m._report["hres_emb"] = hres_emb

    # mapear embalse -> area (para slacks vmin/vmax por area)
    res_area = {}
    for r in RES:
        gens = [g for g in HRES if hres_emb.get(g) == r]
        if gens:
            res_area[r] = hres_area.get(gens[0], "")
        else:
            res_area[r] = ""
    m._report["res_area"] = res_area
    m._report["lmp_balance"] = lmp
    m._report["dual_exch_ub"] = dual_exch_ub
    m._report["dual_exch_lb"] = dual_exch_lb

    return m, results

# ============================================================
# 8) Resultados
# ============================================================

def export_results_to_excel(m: ConcreteModel, scenario_name: str, out_path: str):
    # Verifica que haya solución
    # (Esta función se llama solo si print_results detectó feasible/optimal)
    rep = getattr(m, "_report", {})
    if not rep:
        raise RuntimeError("No encuentro m._report. Asegúrate de haberlo guardado en build_and_solve_model().")

    demand = rep["demand"]
    ror_area = rep["ror_area"]
    hres_area = rep["hres_area"]
    tc_area = rep["tc_area"]
    tnc_area = rep["tnc_area"]
    ren_area = rep["ren_area"]

    fuel_cost = rep["fuel_cost"]
    tc_hr = rep["tc_hr"]
    tc_om = rep["tc_om"]
    tc_fid = rep["tc_fid"]
    #tc_startup_kusd = rep["tc_startup_kusd"]

    FLOW_TO_HM3_PER_H = rep["FLOW_TO_HM3_PER_H"]
    spill_cost_usd_per_hm3 = rep["spill_cost_usd_per_hm3"]

    exch_lim = rep["exch_lim"]


    A = list(m.A)
    H = list(m.H)

    lmp = rep.get("lmp_balance", {})
    dual_exch_ub = rep.get("dual_exch_ub", {})
    dual_exch_lb = rep.get("dual_exch_lb", {})

    rows_lmp = []
    for a in A:
        for h in H:
            rows_lmp.append({
                "scenario": scenario_name,
                "id_area": str(a),
                "hour": int(h),
                "lmp_usd_per_mwh": float(lmp.get((a, h), np.nan)),
                "dual_exch_ub": float(dual_exch_ub.get((a, h), np.nan)),
                "dual_exch_lb": float(dual_exch_lb.get((a, h), np.nan)),
            })
    df_lmp = pd.DataFrame(rows_lmp)

    
    rows = []
    
    # ----------------------------
    # 1) Generación por tecnología, por área y hora (ROR, HRES, TC, NTC(TNC), REN)
    # ----------------------------
    for a in A:
        for h in H:
            gen_ror = sum(value(m.p_ror[r, h]) for r in m.ROR if ror_area.get(r) == a)
            gen_hres = sum(value(m.p_hres[g, h]) for g in m.HRES if hres_area.get(g) == a)
            gen_tc = sum(value(m.p_tc[g, h]) for g in m.TC if tc_area.get(g) == a)
            gen_tnc = sum(value(m.p_tnc[g, h]) for g in m.TNC if tnc_area.get(g) == a)
            gen_ren = sum(value(m.p_ren[g, h]) for g in m.REN if ren_area.get(g) == a)
            exch = value(m.exch[a, h])

            dem = float(demand.get((a, h), 0.0))

            rows.append({
                "scenario": scenario_name,
                "id_area": a,
                "hour": h,
                "demand_mw": dem,
                "exchange_mw": exch,
                "gen_ror_mw": gen_ror,
                "gen_hres_mw": gen_hres,
                "gen_tc_mw": gen_tc,
                "gen_ntc_mw": gen_tnc,   # NTC solicitado (tu modelo lo llama TNC)
                "gen_ren_mw": gen_ren,
            })

    df_gen = pd.DataFrame(rows)

    df_gen["import_mw"] = df_gen["exchange_mw"].clip(lower=0.0)
    df_gen["export_mw"] = (-df_gen["exchange_mw"]).clip(lower=0.0)

    # ----------------------------
    # RENOVABLE POR PLANTA Y HORA
    # ----------------------------
    rows_ren_planta = []
    
    for g in m.REN:
        a = ren_area.get(g, "")
        for h in H:
            p = float(value(m.p_ren[g, h]))
            c = float(value(m.curt[g, h]))
    
            rows_ren_planta.append({
                "scenario": scenario_name,
                "id_planta": g,
                "id_area": a,
                "hour": h,
                "gen_ren_mw": p,
                "curtailment_mw": c,
                "avail_mw": p + c
            })
    
    df_ren_planta_h = pd.DataFrame(rows_ren_planta)

    df_ren_planta_day = (
        df_ren_planta_h
        .groupby(["scenario", "id_planta", "id_area"], as_index=False)
        .agg({
            "gen_ren_mw": "sum",
            "curtailment_mw": "sum",
            "avail_mw": "sum"
        })
    )


    #Diagnóstico de intercambios:
    
    rows_exch = []
    eps = 1e-6
    
    for a in A:
        lb, ub = exch_lim.get(a, (-9999.0, 9999.0))
    
        for h in H:
            exch = float(value(m.exch[a, h]))
    
            # generación por tecnología del área (MISMO criterio que df_gen)
            gen_ror  = sum(value(m.p_ror[r, h])  for r in m.ROR  if ror_area.get(r) == a)
            gen_hres = sum(value(m.p_hres[g, h]) for g in m.HRES if hres_area.get(g) == a)
            gen_tc   = sum(value(m.p_tc[g, h])   for g in m.TC   if tc_area.get(g) == a)
            gen_tnc  = sum(value(m.p_tnc[g, h])  for g in m.TNC  if tnc_area.get(g) == a)
            gen_ren  = sum(value(m.p_ren[g, h])  for g in m.REN  if ren_area.get(g) == a)
    
            gen_total = gen_ror + gen_hres + gen_tc + gen_tnc + gen_ren
    
            curt_a = sum(value(m.curt[g, h]) for g in m.REN if ren_area.get(g) == a)
            dump_a = float(value(m.dump[a, h]))
            shed_a = float(value(m.load_shed[a, h]))
            dem = float(demand.get((a, h), 0.0))
    
            rows_exch.append({
                "scenario": scenario_name,
                "id_area": a,
                "hour": h,
    
                "demand_mw": dem,
    
                # 👇 ahora quedan también por tecnología
                "gen_ror_mw": gen_ror,
                "gen_hres_mw": gen_hres,
                "gen_tc_mw": gen_tc,
                "gen_tnc_mw": gen_tnc,
                "gen_ren_mw": gen_ren,
                "gen_total_mw": gen_total,
    
                "curtailment_mw": curt_a,
                "dump_mw": dump_a,
                "load_shed_mw": shed_a,
    
                "exchange_mw": exch,
                "lb_mw": lb,
                "ub_mw": ub,
    
                # binding robusto (mejor que abs(exch-lb)<1e-3 si hay tolerancias)
                "bind_lb": int(exch <= lb + eps),
                "bind_ub": int(exch >= ub - eps),
    
                "import_mw": max(exch, 0.0),
                "export_mw": max(-exch, 0.0),
            })
    
    df_exch_diag = pd.DataFrame(rows_exch)

    # Totales diarios por área (sum 24h)
    df_gen_area = (
        df_gen.groupby(["scenario", "id_area"], as_index=False)
        .agg({
            "demand_mw": "sum",
            "gen_ror_mw": "sum",
            "gen_hres_mw": "sum",
            "gen_tc_mw": "sum",
            "gen_ntc_mw": "sum",
            "gen_ren_mw": "sum",
            "exchange_mw": "sum",
            "import_mw": "sum",
            "export_mw": "sum",
        })
    )

    df_flow_area = df_gen_area[[
        "scenario", "id_area",
        "exchange_mw", "import_mw", "export_mw"
    ]].copy()

   
    # NUEVO: REPORTE EXCH POR HORA + BINDING A LIMITES
    # ===========================
    
    # 1) Reconstruir exch_lim en esta función (porque aquí no lo estás pasando)
    #    (usa la hoja Limites_areas del Excel que ya está accesible como data en build_and_solve_model,
    #     pero export_results_to_excel no la recibe; por eso: guarda exch_lim en m._report y léelo aquí)
    exch_lim = rep.get("exch_lim", None)
    if exch_lim is None:
        exch_lim = {a: (-9999.0, 9999.0) for a in A}  # fallback
    
    # 2) agregar lb/ub por área
    df_exch_h = df_gen[["scenario","id_area","hour","exchange_mw"]].copy()
    df_exch_h["lb_mw"] = df_exch_h["id_area"].map(lambda a: float(exch_lim.get(str(a), (-9999.0, 9999.0))[0]))
    df_exch_h["ub_mw"] = df_exch_h["id_area"].map(lambda a: float(exch_lim.get(str(a), (-9999.0, 9999.0))[1]))
    
    # 3) flags si está pegando en límites (tolerancia)
    TOL = 1e-3
    df_exch_h["bind_lb"] = (df_exch_h["exchange_mw"] <= df_exch_h["lb_mw"] + TOL).astype(int)
    df_exch_h["bind_ub"] = (df_exch_h["exchange_mw"] >= df_exch_h["ub_mw"] - TOL).astype(int)
    
    # 4) import/export por hora (por claridad)
    df_exch_h["import_mw"] = df_exch_h["exchange_mw"].clip(lower=0.0)
    df_exch_h["export_mw"] = (-df_exch_h["exchange_mw"]).clip(lower=0.0)


    
    # Totales diarios por tecnología (suma todas las áreas)
    df_gen_tech = pd.DataFrame([{
        "scenario": scenario_name,
        "gen_ror_mwh": df_gen["gen_ror_mw"].sum(),
        "gen_hres_mwh": df_gen["gen_hres_mw"].sum(),
        "gen_tc_mwh": df_gen["gen_tc_mw"].sum(),
        "gen_ntc_mwh": df_gen["gen_ntc_mw"].sum(),
        "gen_ren_mwh": df_gen["gen_ren_mw"].sum(),
        "demand_mwh": df_gen["demand_mw"].sum(),
    }])

    # ----------------------------
    # 2) Curtailment por área y hora (sum de REN en el área)
    # ----------------------------
    rows_curt = []
    for a in A:
        for h in H:
            curt_a = sum(value(m.curt[g, h]) for g in m.REN if ren_area.get(g) == a)
            rows_curt.append({
                "scenario": scenario_name,
                "id_area": a,
                "hour": h,
                "curtailment_mw": curt_a
            })
    df_curt = pd.DataFrame(rows_curt)
    df_curt_area = df_curt.groupby(["scenario", "id_area"], as_index=False)["curtailment_mw"].sum()

    # ----------------------------
    # 3) Spill por área y hora (sum q_spill de HRES en el área)
    #    - en m3/s y también convertido a hm3/h (útil para costos)
    # ----------------------------
    rows_spill = []
    for a in A:
        for h in H:
            spill_m3s = sum(value(m.q_spill[g, h]) for g in m.HRES if hres_area.get(g) == a)
            spill_hm3ph = spill_m3s * FLOW_TO_HM3_PER_H
            rows_spill.append({
                "scenario": scenario_name,
                "id_area": a,
                "hour": h,
                "spill_m3s": spill_m3s,
                "spill_hm3_per_h": spill_hm3ph,
            })
    df_spill = pd.DataFrame(rows_spill)
    df_spill_area = df_spill.groupby(["scenario", "id_area"], as_index=False).agg({
        "spill_m3s": "sum",
        "spill_hm3_per_h": "sum"
    })

    # ----------------------------
    # 4) Costos TC por área (combustible+HR + OM) + startup
    # ----------------------------
    rows_tc_cost = []
    for a in A:
        for h in H:
            cost_energy = 0.0
            cost_start = 0.0
            for g in m.TC:
                if tc_area.get(g) != a:
                    continue
                fid = tc_fid.get(g, "")
                fc = float(fuel_cost.get(fid, 0.0))
                hr = float(tc_hr.get(g, 0.0))
                om = float(tc_om.get(g, 0.0))

                p = float(value(m.p_tc[g, h]))
                cost_energy += (hr * fc + om) * p

                # su = float(tc_startup_kusd.get(g, 0.0))
                # cost_start += su * float(value(m.tc_start[g, h]))

            rows_tc_cost.append({
                "scenario": scenario_name,
                "id_area": a,
                "hour": h,
                "tc_cost_energy_usd": cost_energy,
                # "tc_cost_startup_usd": cost_start,
                "tc_cost_total_usd": cost_energy + cost_start
            })

    df_tc_cost = pd.DataFrame(rows_tc_cost)
    df_tc_cost_area = df_tc_cost.groupby(["scenario", "id_area"], as_index=False).agg({
        "tc_cost_energy_usd": "sum",
        #"tc_cost_startup_usd": "sum",
        "tc_cost_total_usd": "sum"
    })

    # (Opcional extra) costo de spill por área (si lo quieres)
    df_spill_cost = df_spill.copy()
    df_spill_cost["spill_cost_usd"] = df_spill_cost["spill_hm3_per_h"] * spill_cost_usd_per_hm3
    df_spill_cost_area = df_spill_cost.groupby(["scenario", "id_area"], as_index=False)["spill_cost_usd"].sum()


        # ============================================================
    # DIAGNÓSTICO CURTAILMENT: causa por área y hora
    # ============================================================
    ren_pmax = rep.get("ren_pmax", {})
    ren_shape = rep.get("ren_shape", {})

    tc_pmin = rep.get("tc_pmin", {})
    hres_pmin = rep.get("hres_pmin", {})
    hres_emb = rep.get("hres_emb", {})
    res_area = rep.get("res_area", {})

    # 1) Renovable disponible por área-hora (sum(pmax*shape))
    ren_avail_area_h = {(a, h): 0.0 for a in A for h in H}
    for g in m.REN:
        a = ren_area.get(str(g), "")
        if a == "":
            continue
        for h in H:
            avail = float(ren_pmax.get(str(g), 0.0)) * float(ren_shape.get((str(g), h), 0.0))
            ren_avail_area_h[(a, h)] += avail

    
    # 4) Spill  por área-hora (solo HRES se mapea fácil a área)
    spill_area_h = {(a, h): 0.0 for a in A for h in H}
    for g in m.HRES:
        a = hres_area.get(str(g), "")
        if a == "":
            continue
        for h in H:
            spill_area_h[(a, h)] += float(value(m.q_spill[g, h]))
           

    # 5) Slack vmin/vmax por área-hora (map RES->area vía res_area)
    slack_vmin_area_h = {(a, h): 0.0 for a in A for h in H}
    slack_vmax_area_h = {(a, h): 0.0 for a in A for h in H}
    for r in m.RES:
        a = res_area.get(str(r), "")
        if a == "":
            continue
        for h in H:
            slack_vmin_area_h[(a, h)] += float(value(m.slack_vmin[r, h]))
            slack_vmax_area_h[(a, h)] += float(value(m.slack_vmax[r, h]))

    # 6) Construir df_curt_cause usando df_exch_diag (ya tiene gen, demand, exch, binds, curt, etc.)
    rows_cause = []
    for _, rr in df_exch_diag.iterrows():
        a = str(rr["id_area"])
        h = int(rr["hour"])

        curt = float(rr["curtailment_mw"])
        avail = float(ren_avail_area_h.get((a, h), 0.0))
        ren_gen = float(rr["gen_ren_mw"])

        bind_ub = int(rr.get("bind_ub", 0))
        bind_lb = int(rr.get("bind_lb", 0))

        spill = float(spill_area_h.get((a, h), 0.0))
        svmin = float(slack_vmin_area_h.get((a, h), 0.0))
        svmax = float(slack_vmax_area_h.get((a, h), 0.0))

        # etiqueta (heurística)
        label = ""
        if curt > 1e-6:
            if bind_ub == 1:
                label = "Congestión: límite exportación (exch pegado a UB)"
            elif bind_lb == 1:
                label = "Congestión: límite importación (exch pegado a LB)"
            else:
                label = "Otro (revisar red/perfiles/unidades)"
        else:
            label = "Sin curtailment"

        rows_cause.append({
            "scenario": scenario_name,
            "id_area": a,
            "hour": h,
            "ren_avail_mw": avail,
            "ren_gen_mw": ren_gen,
            "curtailment_mw": curt,
            "exch_mw": float(rr["exchange_mw"]),
            "lb_mw": float(rr["lb_mw"]),
            "ub_mw": float(rr["ub_mw"]),
            "bind_lb": bind_lb,
            "bind_ub": bind_ub,
            "spill_m3s": spill,
            "slack_vmin": svmin,
            "slack_vmax": svmax,
            "cause_label": label,
        })

    df_curt_cause = pd.DataFrame(rows_cause)

    # ----------------------------
    # ESCRIBIR EXCEL (múltiples hojas)
    # ----------------------------
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_gen.to_excel(writer, sheet_name=f"{scenario_name}_GEN_H", index=False)
        df_gen_area.to_excel(writer, sheet_name=f"{scenario_name}_GEN_AREA", index=False)
        df_gen_tech.to_excel(writer, sheet_name=f"{scenario_name}_GEN_TECH", index=False)

        df_curt.to_excel(writer, sheet_name=f"{scenario_name}_CURT_H", index=False)
        df_curt_area.to_excel(writer, sheet_name=f"{scenario_name}_CURT_AREA", index=False)

        df_spill.to_excel(writer, sheet_name=f"{scenario_name}_SPILL_H", index=False)
        df_spill_area.to_excel(writer, sheet_name=f"{scenario_name}_SPILL_AREA", index=False)

        df_tc_cost.to_excel(writer, sheet_name=f"{scenario_name}_TC_COST_H", index=False)
        df_tc_cost_area.to_excel(writer, sheet_name=f"{scenario_name}_TC_COST_AREA", index=False)

        df_spill_cost_area.to_excel(writer, sheet_name=f"{scenario_name}_SPILL_COST_AREA", index=False)

        df_flow_area.to_excel(writer, sheet_name=f"{scenario_name}_FLOW_AREA", index=False)
        df_exch_h.to_excel(writer, sheet_name=f"{scenario_name}_EXCH_H", index=False)
        df_exch_diag.to_excel(writer, sheet_name=f"{scenario_name}_EXCH_DIAG", index=False)

        df_ren_planta_h.to_excel(
            writer,
            sheet_name=f"{scenario_name}_REN_PLANTA_H",
            index=False
        )
        
        df_ren_planta_day.to_excel(
            writer,
            sheet_name=f"{scenario_name}_REN_PLANTA_DAY",
            index=False
        )

        df_curt_cause.to_excel(writer, sheet_name=f"{scenario_name}_CURT_CAUSE", index=False)
        df_lmp.to_excel(writer, sheet_name=f"{scenario_name}_LMP", index=False)



def print_results(m: ConcreteModel, results, scenario_name: str, enable_bess: bool):
    tc = results.solver.termination_condition
    
    print("\n" + "=" * 72)
    print(f"ESCENARIO: {scenario_name}")
    print("=" * 72)
    print(f"Termination condition: {tc}")
    
    if tc not in (TerminationCondition.optimal, TerminationCondition.feasible):
        print("\nNo hay solución factible/óptima => no se imprimen valores (variables no inicializadas).")
        return

    # ============================================================
    # EXPORTAR EXCEL (solo si hay solución)
    # ============================================================
    out_excel = rf"D:\beca fabio chaparro\tesis\modelizacion\STP Y SOC\analisis energetico25-30\RESULTADOS_{scenario_name}.xlsx"
    export_results_to_excel(m, scenario_name=scenario_name, out_path=out_excel)
    print(f"\n✅ Excel de resultados guardado en:\n{out_excel}\n")


    obj = value(m.Obj)
    
    def sum_var_2d(var, I, H):
        return sum(value(var[i, h]) for i in I for h in H)

    gen_ror  = sum_var_2d(m.p_ror,  m.ROR,  m.H)
    gen_hres = sum_var_2d(m.p_hres, m.HRES, m.H)
    gen_tc   = sum_var_2d(m.p_tc,   m.TC,   m.H)
    gen_tnc  = sum_var_2d(m.p_tnc,  m.TNC,  m.H)
    gen_ren  = sum_var_2d(m.p_ren,  m.REN,  m.H)

    curt = sum_var_2d(m.curt, m.REN, m.H)
    shed = sum_var_2d(m.load_shed, m.A, m.H)
    dump = sum_var_2d(m.dump, m.A, m.H)

    bess_dis = sum_var_2d(m.p_dis, m.BESS, m.H) if enable_bess and len(list(m.BESS)) else 0.0
    bess_ch  = sum_var_2d(m.p_ch,  m.BESS, m.H) if enable_bess and len(list(m.BESS)) else 0.0

    print(f"\nCosto total (USD): {obj:,.2f}\n")
    print("Generación total diaria (aprox. MWh):")
    print(f"  Hidro ROR           : {gen_ror:,.2f}")
    print(f"  Hidro Embalse       : {gen_hres:,.2f}")
    print(f"  Térmica (TC)        : {gen_tc:,.2f}")
    print(f"  Térmica (TNC)       : {gen_tnc:,.2f}")
    print(f"  Renovable           : {gen_ren:,.2f}")
    if enable_bess and len(list(m.BESS)) > 0:
        print(f"  BESS descarga       : {bess_dis:,.2f}")
        print(f"  BESS carga          : {bess_ch:,.2f}")

    print("\nSlacks / pérdidas:")
    print(f"  Load shedding total : {shed:,.2f}")
    print(f"  Curtailment total   : {curt:,.2f}")
    print(f"  Dump total          : {dump:,.2f}")

    if len(list(m.RES)) > 0:
        last_h = max(list(m.H))
        print("\nVolumen final embalses (hm3):")
        for r in m.RES:
            print(f"  {r}: {value(m.v[r, last_h]):,.3f}")

    if enable_bess and len(list(m.BESS)) > 0:
        last_h = max(list(m.H))
        print("\nSOC final BESS (MWh):")
        for b in m.BESS:
            print(f"  {b}: {value(m.soc[b, last_h]):,.3f}")


    # dentro de print_results, cuando ya hay solucion
    last_h = max(list(m.H))
    
    bind_count = 0
    total_points = 0
    max_gap = 0.0
    
    for r in m.ROR:
        # necesitas acceder a ror_pmax/ror_shape/ror_area -> si no están accesibles aquí,
        # pásalos como argumento o imprime este check dentro de build_and_solve_model antes del return.
        pass

    
# ============================================================
# 9) MAIN: llamar la función (si no, no ejecuta nada)
# ============================================================

# AQUI SE INCLUYO: LLAMADA REAL A build_and_solve_model PARA QUE "EJECUTE ALGO"
def main():
    data = load_all_tables_from_excel(EXCEL_PATH, SHEETS_REQUIRED)

    m1, r1 = build_and_solve_model(data, scenario_name="No-BESS", enable_bess=False)
    print_results(m1, r1, "No-BESS", enable_bess=False)

    m2, r2 = build_and_solve_model(data, scenario_name="BESS", enable_bess=True)
    print_results(m2, r2, "BESS", enable_bess=True)

    

if __name__ == "__main__":
    main()

