import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, numbers
from openpyxl.utils import get_column_letter

# ============================================================================
# CONFIGURACION NADF-009-AIRE-2017 (ICA)
# ============================================================================
VENTANAS_NADF = {
    "O3_ppb": 1, "NO2_ppb": 1, "SO2_ppb": 24,
    "CO_ppm": 8, "PM10_ug/m3": 24, "PM2.5_ug/m3": 24
}

BANDAS_NADF = {
    "O3_ppm": [(0.000,0.070,0,50),(0.071,0.095,51,100),(0.096,0.154,101,150),
               (0.155,0.204,151,200),(0.205,0.404,201,300),(0.405,0.504,301,400),
               (0.505,0.604,401,500)],
    "NO2_ppm": [(0.000,0.105,0,50),(0.106,0.210,51,100),(0.211,0.430,101,150),
                (0.431,0.649,151,200),(0.650,1.249,201,300),(1.250,1.649,301,400),
                (1.650,2.049,401,500)],
    "SO2_ppm": [(0.000,0.025,0,50),(0.026,0.110,51,100),(0.111,0.207,101,150),
                (0.208,0.304,151,200),(0.305,0.604,201,300),(0.605,0.804,301,400),
                (0.805,1.004,401,500)],
    "CO_ppm": [(0.0,5.5,0,50),(5.6,11.0,51,100),(11.1,13.0,101,150),
               (13.1,15.4,151,200),(15.5,30.4,201,300),(30.5,40.4,301,400),
               (40.5,50.4,401,500)],
    "PM10_ug/m3": [(0,40,0,50),(41,75,51,100),(76,214,101,150),(215,354,151,200),
                   (355,424,201,300),(425,504,301,400),(505,604,401,500)],
    "PM2.5_ug/m3": [(0.0,12.0,0,50),(12.1,45.0,51,100),(45.1,97.4,101,150),
                    (97.5,150.4,151,200),(150.5,250.4,201,300),(250.5,350.4,301,400),
                    (350.5,500.4,401,500)]
}

COLORES_NADF = {
    (0,50): "9ACA3C", (51,100): "F7EC0F", (101,150): "F8991D",
    (151,200): "ED2124", (201,300): "7D287D", (301,500): "7E0023"
}

# ============================================================================
# CONFIGURACION NOM-172-SEMARNAT-2023 (AIRE Y SALUD)
# ============================================================================
VENTANAS_NOM = {
    "O3_ppb": 1, "NO2_ppb": 1, "SO2_ppb": 1,
    "CO_ppm": 8, "PM10_ug/m3": 12, "PM2.5_ug/m3": 12
}

# Bandas para clasificacion (segun Tablas 4-9, valores a partir de enero 2026)
BANDAS_NOM = {
    "O3_ppm": [(0, 0.058, "Buena"), (0.058, 0.090, "Aceptable"), (0.090, 0.135, "Mala"),
               (0.135, 0.175, "Muy Mala"), (0.175, float("inf"), "Extremadamente Mala")],
    "NO2_ppm": [(0, 0.053, "Buena"), (0.053, 0.106, "Aceptable"), (0.106, 0.160, "Mala"),
                (0.160, 0.213, "Muy Mala"), (0.213, float("inf"), "Extremadamente Mala")],
    "SO2_ppm": [(0, 0.035, "Buena"), (0.035, 0.075, "Aceptable"), (0.075, 0.185, "Mala"),
                (0.185, 0.304, "Muy Mala"), (0.304, float("inf"), "Extremadamente Mala")],
    "CO_ppm": [(0, 5.00, "Buena"), (5.00, 9.00, "Aceptable"), (9.00, 12.00, "Mala"),
               (12.00, 16.00, "Muy Mala"), (16.00, float("inf"), "Extremadamente Mala")],
    "PM10_ug/m3": [(0, 45, "Buena"), (45, 50, "Aceptable"), (50, 132, "Mala"),
                   (132, 213, "Muy Mala"), (213, float("inf"), "Extremadamente Mala")],
    "PM2.5_ug/m3": [(0, 15, "Buena"), (15, 25, "Aceptable"), (25, 79, "Mala"),
                    (79, 130, "Muy Mala"), (130, float("inf"), "Extremadamente Mala")]
}

# Colores oficiales (Tabla 11)
COLORES_NOM = {
    "Buena": "00E400",          # Verde
    "Aceptable": "FFFF00",       # Amarillo
    "Mala": "FF7E00",            # Naranja
    "Muy Mala": "FF0000",        # Rojo
    "Extremadamente Mala": "8F3F97"  # Morado
}

# Orden de las categorias para determinar la peor
ORDEN_CATEGORIAS = {
    "Buena": 0,
    "Aceptable": 1,
    "Mala": 2,
    "Muy Mala": 3,
    "Extremadamente Mala": 4
}

SUFICIENCIA = 0.75   # 75% de datos requeridos

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def calcular_ica(conc, bandas):
    """Calcula ICA segun NADF-009 mediante interpolacion lineal"""
    for pcinf, pcsup, iinf, isup in bandas:
        if pcinf <= conc <= pcsup:
            k = (isup - iinf) / (pcsup - pcinf)
            return round((k * (conc - pcinf)) + iinf)
    return np.nan

def clasificar_nom(conc, bandas):
    """Asigna categoria segun NOM-172 a partir de la concentracion"""
    if pd.isna(conc):
        return None
    for lim_inf, lim_sup, cat in bandas:
        if lim_inf < conc <= lim_sup:
            return cat
        elif conc == lim_inf and lim_inf == 0:   # incluye el cero en primera banda
            return cat
    return None

def promedio_movil_simple(serie, ventana):
    """Promedio movil con suficiencia del 75%"""
    min_datos = int(np.ceil(ventana * SUFICIENCIA))
    return serie.rolling(window=ventana, min_periods=min_datos).mean()

def nowcast(serie, pollutant):
    """
    Implementacion del NowCast segun NOM-172 Anexo A.
    Solo calcula desde la hora 11 (12 horas de datos) y requiere al menos
    2 de las 3 horas mas recientes.
    """
    fa = 0.714 if pollutant == "PM10" else 0.694
    valores = serie.values
    n = len(valores)
    resultado = np.full(n, np.nan)
    for i in range(n):
        if i < 11:            # no hay suficientes horas
            continue
        ultimas3 = valores[i-2:i+1]
        if np.sum(~np.isnan(ultimas3)) < 2:   # condicion de las 3 ultimas
            continue
        inicio = i - 11
        ventana = valores[inicio:i+1]
        validos = ventana[~np.isnan(ventana)]
        if len(validos) == 0:
            continue
        cmax = np.max(validos)
        cmin = np.min(validos)
        if cmax == 0:
            w = 1.0
        else:
            w = 1 - (cmax - cmin) / cmax
        W = round(max(w, 0.5), 2)          # factor de ponderacion redondeado
        suma_num = 0.0
        suma_den = 0.0
        for j, idx in enumerate(range(i, inicio-1, -1)):
            if j >= 12:
                break
            if not np.isnan(valores[idx]):
                peso = W ** j
                suma_num += valores[idx] * peso
                suma_den += peso
        if suma_den > 0:
            resultado[i] = (suma_num / suma_den) * fa
    return pd.Series(resultado, index=serie.index)

def redondear_nom(valor, contaminante, unidad):
    """Redondeo segun Tabla 2 y punto 5.2.4 de la NOM-172"""
    if pd.isna(valor):
        return np.nan
    if contaminante in ["PM10", "PM2.5"]:
        return int(round(valor))
    elif contaminante in ["O3", "NO2", "SO2"]:
        return round(valor, 3)
    elif contaminante == "CO":
        return round(valor, 2)
    else:
        return valor

def obtener_color_ica(valor):
    """Devuelve color hexadecimal para un valor ICA segun NADF-009"""
    for (lo, hi), color in COLORES_NADF.items():
        if lo <= valor <= hi:
            return color
    return None

def preparar_datos_hoja(df):
    """
    Convierte las primeras filas en metadatos y reindexa a frecuencia horaria.
    Retorna estaciones, contaminantes, unidades, data_df (con indice datetime) y numero de columnas.
    """
    estaciones = df.iloc[0].values
    contaminantes = df.iloc[1].values
    unidades = df.iloc[2].values
    datos_raw = df.iloc[3:].reset_index(drop=True)

    # Procesar fechas (formato dd/mm/yyyy)
    dates_raw = datos_raw.iloc[:, 0]
    dates = pd.to_datetime(dates_raw, errors='coerce', dayfirst=True)
    if dates.isna().any():
        print(f"ADVERTENCIA: {dates.isna().sum()} filas con fecha no valida seran descartadas.")
        datos_raw = datos_raw.loc[~dates.isna()]
        dates = dates[~dates.isna()]
    datos_raw.index = dates
    datos_raw.index.name = 'Fecha'
    datos_raw = datos_raw.drop(columns=0)
    if datos_raw.index.duplicated().any():
        print(f"ADVERTENCIA: {datos_raw.index.duplicated().sum()} indices duplicados; se conserva la primera ocurrencia.")
        datos_raw = datos_raw[~datos_raw.index.duplicated(keep='first')]
    full_range = pd.date_range(start=datos_raw.index.min(), end=datos_raw.index.max(), freq='h')
    data_df = datos_raw.reindex(full_range)
    gaps = full_range.difference(data_df.index)
    if len(gaps) > 0:
        print(f"INFO: Se agregaron {len(gaps)} horas faltantes (NaN).")
    return estaciones, contaminantes, unidades, data_df, len(df.columns)

def peor_categoria(series_categorias):
    """Dadas varias series de categorias, devuelve la peor (mayor riesgo) por fila"""
    if not series_categorias:
        return pd.Series(index=pd.Index([]), dtype='object')
    df_cat = pd.concat(series_categorias, axis=1)
    df_num = df_cat.apply(lambda col: col.map(ORDEN_CATEGORIAS).fillna(-1))
    max_num = df_num.max(axis=1)
    inverso = {v: k for k, v in ORDEN_CATEGORIAS.items()}
    return max_num.map(inverso).where(max_num >= 0, None)

# ============================================================================
# PROCESAMIENTO PRINCIPAL
# ============================================================================
archivo_entrada = "datos/datos_calidad_aire.xlsx"
salida_ica = "datos/datos_calidad_aire_ICA.xlsx"
salida_aire = "datos/datos_calidad_aire_AIRE_Y_SALUD.xlsx"
salida_diario = "datos/datos_calidad_aire_DIARIO.xlsx"

xls = pd.ExcelFile(archivo_entrada)

# ----------------------------------------------------------------------------
# 1. Archivo ICA (NADF-009)
# ----------------------------------------------------------------------------
with pd.ExcelWriter(salida_ica, engine="openpyxl") as writer:
    for hoja in xls.sheet_names:
        print(f"Procesando hoja {hoja} para ICA...")
        df = pd.read_excel(xls, sheet_name=hoja, header=None)
        estaciones, contaminantes, unidades, data_df, num_orig_cols = preparar_datos_hoja(df)

        columnas_salida = {"Fecha & Hora": data_df.index}

        for i in range(1, num_orig_cols):
            col_in_data = i - 1
            estacion = estaciones[i]
            contaminante = contaminantes[i]
            unidad = unidades[i]

            if not isinstance(contaminante, str) or contaminante == "Status":
                continue

            clave_orig = f"{contaminante}_{unidad}"
            if clave_orig not in VENTANAS_NADF:
                continue

            ventana = VENTANAS_NADF[clave_orig]
            valores = pd.to_numeric(data_df.iloc[:, col_in_data], errors="coerce")

            # Aplicar filtro de status: solo se conserva si la celda de status es exactamente "ok"
            if i + 1 < num_orig_cols:
                status_series = data_df.iloc[:, i]
                status_str = status_series.astype(str).str.strip().str.lower()
                valores = valores.where(status_str == "ok", np.nan)

            # Descartar valores negativos
            valores = valores.where(valores >= 0, np.nan)

            if (valores == 0).all():
                print(f"ADVERTENCIA: {contaminante} en {estacion} tiene todos los valores en 0.")
                valores[:] = np.nan

            valores_prom = promedio_movil_simple(valores, ventana)

            # Convertir gases de ppb a ppm
            if contaminante in ["O3", "NO2", "SO2"]:
                valores_prom = valores_prom / 1000.0
                clave_bandas = f"{contaminante}_ppm"
            else:
                clave_bandas = clave_orig

            ica_lista = [calcular_ica(x, BANDAS_NADF[clave_bandas]) if not np.isnan(x) else np.nan for x in valores_prom]
            columnas_salida[f"ICA_{contaminante}_{estacion}"] = ica_lista

        df_salida = pd.DataFrame(columnas_salida)
        cols_datos = [c for c in df_salida.columns if c != "Fecha & Hora"]
        df_salida = df_salida.dropna(how='all', subset=cols_datos)
        df_salida.to_excel(writer, sheet_name=hoja, index=False)

# Formato y colores para ICA
wb_ica = load_workbook(salida_ica)
for hoja in wb_ica.sheetnames:
    ws = wb_ica[hoja]
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column > 1 and isinstance(cell.value, (int, float)) and not pd.isna(cell.value):
                color = obtener_color_ica(int(cell.value))
                if color:
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 25
wb_ica.save(salida_ica)
print("Archivo NADF-009 (ICA) generado.")

# ----------------------------------------------------------------------------
# 2. Archivo AIRE Y SALUD horario (NOM-172)
# ----------------------------------------------------------------------------
with pd.ExcelWriter(salida_aire, engine="openpyxl") as writer:
    for hoja in xls.sheet_names:
        print(f"Procesando hoja {hoja} para AIRE Y SALUD...")
        df = pd.read_excel(xls, sheet_name=hoja, header=None)
        estaciones, contaminantes, unidades, data_df, num_orig_cols = preparar_datos_hoja(df)

        columnas_salida = {"Fecha & Hora": data_df.index}
        lista_categorias = []

        for i in range(1, num_orig_cols):
            col_in_data = i - 1
            estacion = estaciones[i]
            contaminante = contaminantes[i]
            unidad = unidades[i]

            if not isinstance(contaminante, str) or contaminante == "Status":
                continue

            clave_orig = f"{contaminante}_{unidad}"
            if clave_orig not in VENTANAS_NOM:
                continue

            valores = pd.to_numeric(data_df.iloc[:, col_in_data], errors="coerce")

            # Filtro de status
            if i + 1 < num_orig_cols:
                status_series = data_df.iloc[:, i]
                status_str = status_series.astype(str).str.strip().str.lower()
                valores = valores.where(status_str == "ok", np.nan)

            valores = valores.where(valores >= 0, np.nan)

            if (valores == 0).all():
                print(f"ADVERTENCIA: {contaminante} en {estacion} tiene todos los valores en 0.")
                valores[:] = np.nan

            # Calcular concentracion base segun tipo de contaminante
            if clave_orig in ["PM10_ug/m3", "PM2.5_ug/m3"]:
                tipo = "PM10" if clave_orig.startswith("PM10") else "PM2.5"
                conc_base = nowcast(valores, tipo)
                clave_bandas = clave_orig
            elif clave_orig == "CO_ppm":
                conc_base = promedio_movil_simple(valores, 8)
                clave_bandas = "CO_ppm"
            else:  # gases O3, NO2, SO2 (en ppb)
                conc_base = valores / 1000.0
                clave_bandas = f"{contaminante}_ppm"

            conc_redondeada = [redondear_nom(x, contaminante, unidad) for x in conc_base]
            categorias = [clasificar_nom(x, BANDAS_NOM[clave_bandas]) for x in conc_redondeada]

            col_cat = f"AIRE_{contaminante}_{estacion}"
            col_conc = f"CANTIDAD_{contaminante}_{estacion}"
            columnas_salida[col_cat] = categorias
            columnas_salida[col_conc] = conc_redondeada
            lista_categorias.append(pd.Series(categorias, index=data_df.index))

        # Calcular la peor categoria por hora
        if lista_categorias:
            columnas_salida["Calidad del aire"] = peor_categoria(lista_categorias)
        else:
            columnas_salida["Calidad del aire"] = None

        df_salida = pd.DataFrame(columnas_salida)
        cols_datos = [c for c in df_salida.columns if c != "Fecha & Hora"]
        df_salida = df_salida.dropna(how='all', subset=cols_datos)
        df_salida.to_excel(writer, sheet_name=hoja, index=False)

# Formato para archivo de Aire y Salud
wb_aire = load_workbook(salida_aire)
for hoja in wb_aire.sheetnames:
    ws = wb_aire[hoja]
    columnas_categoria = []
    for col in ws.iter_cols(min_row=1, max_row=1):
        if col[0].value and isinstance(col[0].value, str) and col[0].value.startswith("AIRE_"):
            columnas_categoria.append(col[0].column)
    for col in ws.iter_cols(min_row=1, max_row=1):
        if col[0].value == "Calidad del aire":
            columnas_categoria.append(col[0].column)
            break

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Colorear celdas de categoria
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in columnas_categoria and cell.value in COLORES_NOM:
                cell.fill = PatternFill(start_color=COLORES_NOM[cell.value],
                                        end_color=COLORES_NOM[cell.value],
                                        fill_type="solid")
                # Contraste de texto
                if cell.value in ["Buena", "Aceptable"]:
                    cell.font = Font(bold=True, color="000000")
                else:
                    cell.font = Font(bold=True, color="FFFFFF")

    # Formato de numero para columnas de cantidad
    for col in ws.columns:
        if col[0].value and isinstance(col[0].value, str) and col[0].value.startswith("CANTIDAD_"):
            nombre_col = col[0].value
            if "O3" in nombre_col or "NO2" in nombre_col or "SO2" in nombre_col:
                fmt = '0.000'
            elif "CO" in nombre_col:
                fmt = '0.00'
            else:  # PM
                fmt = '0'
            for cell in col[1:]:
                if cell.value is not None:
                    cell.number_format = fmt

    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 25

wb_aire.save(salida_aire)
print("Archivo NOM-172 (AIRE Y SALUD) generado.")

# ----------------------------------------------------------------------------
# 3. Archivo DIARIO (NOM-172)
# ----------------------------------------------------------------------------
with pd.ExcelWriter(salida_diario, engine="openpyxl") as writer:
    for hoja in xls.sheet_names:
        print(f"Procesando hoja {hoja} para reporte DIARIO...")
        df = pd.read_excel(xls, sheet_name=hoja, header=None)
        estaciones, contaminantes, unidades, data_df, num_orig_cols = preparar_datos_hoja(df)

        data_df['Fecha_dia'] = data_df.index.date
        dias = data_df['Fecha_dia'].unique()
        dias_ordenados = sorted(dias)

        columnas_diario = {"Fecha": dias_ordenados}
        dict_categorias_diarias = {}

        for i in range(1, num_orig_cols):
            col_in_data = i - 1
            estacion = estaciones[i]
            contaminante = contaminantes[i]
            unidad = unidades[i]

            if not isinstance(contaminante, str) or contaminante == "Status":
                continue

            clave_orig = f"{contaminante}_{unidad}"
            if clave_orig not in VENTANAS_NOM:
                continue

            valores = pd.to_numeric(data_df.iloc[:, col_in_data], errors="coerce")

            # Filtro de status
            if i + 1 < num_orig_cols:
                status_series = data_df.iloc[:, i]
                status_str = status_series.astype(str).str.strip().str.lower()
                valores = valores.where(status_str == "ok", np.nan)

            valores = valores.where(valores >= 0, np.nan)

            serie_valores = pd.Series(valores.values, index=data_df.index)

            # Valor diario segun contaminante
            if clave_orig in ["PM10_ug/m3", "PM2.5_ug/m3"]:
                min_datos = int(np.ceil(24 * SUFICIENCIA))  # 18
                valor_diario = serie_valores.resample('D').apply(
                    lambda x: x.mean() if x.count() >= min_datos else np.nan
                )
                valor_redondeado = valor_diario.apply(lambda v: int(round(v)) if not pd.isna(v) else np.nan)
                clave_bandas = clave_orig
            elif clave_orig == "CO_ppm":
                prom_8h = promedio_movil_simple(serie_valores, 8)
                valor_diario = prom_8h.resample('D').max()
                valor_redondeado = valor_diario.apply(lambda v: round(v, 2) if not pd.isna(v) else np.nan)
                clave_bandas = "CO_ppm"
            else:  # gases
                serie_ppm = serie_valores / 1000.0
                valor_diario = serie_ppm.resample('D').max()
                valor_redondeado = valor_diario.apply(lambda v: round(v, 3) if not pd.isna(v) else np.nan)
                clave_bandas = f"{contaminante}_ppm"

            categorias_diarias = [clasificar_nom(v, BANDAS_NOM[clave_bandas]) for v in valor_redondeado]

            col_cat = f"AIRE_{contaminante}_{estacion}"
            col_conc = f"CANTIDAD_{contaminante}_{estacion}"
            columnas_diario[col_cat] = categorias_diarias
            columnas_diario[col_conc] = valor_redondeado.values
            dict_categorias_diarias[col_cat] = pd.Series(categorias_diarias, index=valor_redondeado.index)

        # Peor categoria diaria
        if dict_categorias_diarias:
            df_cat_diario = pd.DataFrame(dict_categorias_diarias)
            df_num = df_cat_diario.apply(lambda col: col.map(ORDEN_CATEGORIAS).fillna(-1))
            max_num = df_num.max(axis=1)
            inverso = {v: k for k, v in ORDEN_CATEGORIAS.items()}
            peor_diario = max_num.map(inverso).where(max_num >= 0, None)
            columnas_diario["Calidad del aire"] = peor_diario.values

        df_diario = pd.DataFrame(columnas_diario)
        cols_datos = [c for c in df_diario.columns if c not in ["Fecha", "Calidad del aire"]]
        df_diario = df_diario.dropna(how='all', subset=cols_datos)
        df_diario.to_excel(writer, sheet_name=hoja, index=False)

# Formato para archivo diario
wb_diario = load_workbook(salida_diario)
for hoja in wb_diario.sheetnames:
    ws = wb_diario[hoja]
    columnas_categoria = []
    for col in ws.iter_cols(min_row=1, max_row=1):
        if col[0].value and isinstance(col[0].value, str) and col[0].value.startswith("AIRE_"):
            columnas_categoria.append(col[0].column)
    for col in ws.iter_cols(min_row=1, max_row=1):
        if col[0].value == "Calidad del aire":
            columnas_categoria.append(col[0].column)
            break

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in columnas_categoria and cell.value in COLORES_NOM:
                cell.fill = PatternFill(start_color=COLORES_NOM[cell.value],
                                        end_color=COLORES_NOM[cell.value],
                                        fill_type="solid")
                if cell.value in ["Buena", "Aceptable"]:
                    cell.font = Font(bold=True, color="000000")
                else:
                    cell.font = Font(bold=True, color="FFFFFF")

    # Formato de numero para cantidades diarias
    for col in ws.columns:
        if col[0].value and isinstance(col[0].value, str) and col[0].value.startswith("CANTIDAD_"):
            nombre_col = col[0].value
            if "O3" in nombre_col or "NO2" in nombre_col or "SO2" in nombre_col:
                fmt = '0.000'
            elif "CO" in nombre_col:
                fmt = '0.00'
            else:
                fmt = '0'
            for cell in col[1:]:
                if cell.value is not None:
                    cell.number_format = fmt

    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 25

wb_diario.save(salida_diario)
print("Archivo NOM-172 (DIARIO) generado.")