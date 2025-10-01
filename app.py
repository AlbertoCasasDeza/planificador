# app.py
import pandas as pd
import streamlit as st
from datetime import timedelta
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Planificador Lotes Jamcal", layout="wide")
st.title("🧠 Planificador de Lotes Salazón Jamcal")

# -------------------------------
# Panel de configuración (globales)
# -------------------------------
st.sidebar.header("Parámetros de planificación")

# Capacidad global ENTRADA SAL
st.sidebar.subheader("Capacidad global · ENTRADA SAL")
cap_ent_1 = st.sidebar.number_input("Entrada · 1º intento", value=3800, step=100, min_value=0, key="cap_ent_1")
cap_ent_2 = st.sidebar.number_input("Entrada · 2º intento", value=4200, step=100, min_value=0, key="cap_ent_2")

# Capacidad global SALIDA SAL
st.sidebar.subheader("Capacidad global · SALIDA SAL")
cap_sal_1 = st.sidebar.number_input("Salida · 1º intento", value=3800, step=100, min_value=0, key="cap_sal_1")
cap_sal_2 = st.sidebar.number_input("Salida · 2º intento", value=4200, step=100, min_value=0, key="cap_sal_2")

# Límite GLOBAL en días naturales entre DIA (recepción) y ENTRADA_SAL
st.sidebar.subheader("Días máx. almacenamiento (GLOBAL)")
dias_max_almacen_global = st.sidebar.number_input("Días máx. almacenamiento (GLOBAL)", value=2, step=1)

# Capacidad de estabilización (valor base)
st.sidebar.subheader("Capacidad cámara de estabilización (GLOBAL)")
estab_cap = st.sidebar.number_input(
    "Capacidad cámara de estabilización (unds)",
    value=4700, step=100, min_value=0
)

# --- Capacidad de PRENSAS (global 1º/2º intento) ---
st.sidebar.subheader("Capacidad global· ENTRADA PRENSAS")
cap_prensas_ent_1 = st.sidebar.number_input("Entrada · 1º intento", value=3800, step=100, min_value=0, key="cap_pr_ent_1")
cap_prensas_ent_2 = st.sidebar.number_input("Entrada · 2º intento", value=4200, step=100, min_value=0, key="cap_pr_ent_2")

st.sidebar.subheader("Capacidad global· SALIDA PRENSAS")
cap_prensas_sal_1 = st.sidebar.number_input("Salida · 1º intento", value=3800, step=100, min_value=0, key="cap_pr_sal_1")
cap_prensas_sal_2 = st.sidebar.number_input("Salida · 2º intento", value=4200, step=100, min_value=0, key="cap_pr_sal_2")

dias_festivos_default = [
    "2025-01-01", "2025-04-18", "2025-05-01", "2025-08-15",
    "2025-10-12", "2025-11-01", "2025-12-25"
]
dias_festivos_list = st.sidebar.multiselect(
    "Selecciona los días festivos",
    options=dias_festivos_default,
    default=dias_festivos_default
)
dias_festivos = pd.to_datetime(dias_festivos_list)

ajuste_finde = st.sidebar.checkbox("Ajustar fines de semana (SALIDA)", value=True)
ajuste_festivos = st.sidebar.checkbox("Ajustar festivos (SALIDA)", value=True)

# Botón opcional para limpiar estado
if st.sidebar.button("🔄 Reiniciar sesión"):
    st.session_state.clear()
    st.rerun()

# -------------------------------
# Subir archivo Excel
# -------------------------------
uploaded_file = st.file_uploader("📂 Sube tu Excel con los lotes", type=["xlsx"])

# -------------------------------
# Funciones auxiliares
# -------------------------------
def es_habil(fecha):
    return fecha.weekday() < 5 and fecha.normalize() not in dias_festivos

def siguiente_habil(fecha):
    f = fecha + timedelta(days=1)
    while not es_habil(f):
        f += timedelta(days=1)
    return f

def anterior_habil(fecha):
    f = fecha - timedelta(days=1)
    while not es_habil(f):
        f -= timedelta(days=1)
    return f

def _sumar_en_rango(dic, fecha_ini, fecha_fin_inclusive, unds):
    if pd.isna(fecha_ini) or pd.isna(fecha_fin_inclusive):
        return
    for d in pd.date_range(fecha_ini, fecha_fin_inclusive, freq="D"):
        d0 = d.normalize()
        dic[d0] = dic.get(d0, 0) + unds

def calcular_estabilizacion_diaria(df_plan: pd.DataFrame, cap: int, estab_cap_overrides: dict | None = None) -> pd.DataFrame:
    carga_total  = {}
    carga_paleta = {}
    carga_jamon  = {}

    for _, r in df_plan.iterrows():
        dia     = r.get("DIA")
        entrada = r.get("ENTRADA_SAL")
        unds    = int(r.get("UNDS", 0) or 0)
        prod    = str(r.get("PRODUCTO", ""))

        if pd.isna(dia) or pd.isna(entrada) or unds <= 0:
            continue

        fin = entrada - pd.Timedelta(days=1)
        if fin.date() < dia.date():
            continue

        for d in pd.date_range(dia.normalize(), fin.normalize(), freq="D"):
            d0 = d.normalize()
            carga_total[d0] = carga_total.get(d0, 0) + unds
            if prod.startswith("P"):
                carga_paleta[d0] = carga_paleta.get(d0, 0) + unds
            elif prod.startswith("J"):
                carga_jamon[d0] = carga_jamon.get(d0, 0) + unds

    if not carga_total:
        return pd.DataFrame(columns=[
            "FECHA", "ESTAB_UNDS", "ESTAB_PALETA", "ESTAB_JAMON",
            "CAPACIDAD", "UTIL_%", "EXCESO"
        ])

    df_estab = (
        pd.Series(carga_total, name="ESTAB_UNDS")
        .sort_index()
        .to_frame()
        .reset_index()
        .rename(columns={"index": "FECHA"})
    )
    df_estab["ESTAB_PALETA"] = df_estab["FECHA"].map(lambda d: int(carga_paleta.get(d.normalize(), 0)))
    df_estab["ESTAB_JAMON"]  = df_estab["FECHA"].map(lambda d: int(carga_jamon.get(d.normalize(), 0)))

    if estab_cap_overrides is None:
        estab_cap_overrides = {}

    def _cap_for_date(d):
        if pd.isna(d):
            return int(cap)
        key = pd.to_datetime(d).normalize()
        if key in estab_cap_overrides:
            return int(estab_cap_overrides[key])
        return int(cap)

    df_estab["CAPACIDAD"] = df_estab["FECHA"].apply(_cap_for_date)
    df_estab["UTIL_%"] = (df_estab["ESTAB_UNDS"] / df_estab["CAPACIDAD"] * 100).round(1)
    df_estab["EXCESO"] = (df_estab["ESTAB_UNDS"] - df_estab["CAPACIDAD"]).clip(lower=0).astype(int)

    df_estab = df_estab[
        ["FECHA", "ESTAB_UNDS", "ESTAB_PALETA", "ESTAB_JAMON",
         "CAPACIDAD", "UTIL_%", "EXCESO"]
    ]
    return df_estab

def generar_excel(df_out, filename="archivo.xlsx"):
    output = BytesIO()
    df_out.to_excel(output, index=False)
    output.seek(0)
    return output

# --- Helper: normaliza floats desde texto con \xa0 y comas ---
def _to_float_or_nan(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).replace("\xa0", "").replace(",", ".").strip()
    try:
        return float(s) if s != "" else pd.NA
    except Exception:
        return pd.NA

# --- Helper: detecta rango EXACTO por MIN/MAX (12.0 y 13.0) ---
def es_rango_12_13(row) -> bool:
    try:
        mn = row.get("MIN_PESO", pd.NA)
        mx = row.get("MAX_PESO", pd.NA)
        if pd.isna(mn) or pd.isna(mx):
            return False
        # Tolerancia pequeña por si vienen como 12.0000001
        return abs(float(mn) - 12.0) < 1e-6 and abs(float(mx) - 13.0) < 1e-6
    except Exception:
        return False

# -------------------------------
# Planificador (incluye prensas)
# -------------------------------
def planificar_filas_na(
    df_plan,
    dias_max_almacen_global,
    dias_max_por_producto,
    estab_cap,
    cap_overrides_ent,
    cap_overrides_sal,
    estab_cap_overrides,
    cap_prensas_ent_1, cap_prensas_ent_2,
    cap_prensas_sal_1, cap_prensas_sal_2,
    cap_overrides_prensas_ent,
    cap_overrides_prensas_sal
):
    df_corr = df_plan.copy()

    # Asegurar columnas auxiliares (incluye prensas)
    for col in ["LOTE_NO_ENCAJA", "ENTRADA_PRENSAS", "SALIDA_PRENSAS"]:
        if col not in df_corr.columns:
            df_corr[col] = pd.NA

    # Cargas ya planificadas (se respetan)
    carga_entrada = df_corr.dropna(subset=["ENTRADA_SAL"]).groupby("ENTRADA_SAL")["UNDS"].sum().to_dict()
    carga_salida  = df_corr.dropna(subset=["SALIDA_SAL"]).groupby("SALIDA_SAL")["UNDS"].sum().to_dict()

    # Ocupación diaria ya existente en estabilización
    estab_stock = {}
    for _, r in df_corr.dropna(subset=["ENTRADA_SAL"]).iterrows():
        dia_rec = r["DIA"]
        ent     = r["ENTRADA_SAL"]
        unds    = r["UNDS"]
        if pd.notna(dia_rec) and pd.notna(ent) and ent.date() > dia_rec.date():
            _sumar_en_rango(estab_stock, dia_rec, ent - pd.Timedelta(days=1), unds)

    # Capacidad ya consumida en PRENSAS
    carga_prensas_entrada = df_corr.dropna(subset=["ENTRADA_PRENSAS"]).groupby("ENTRADA_PRENSAS")["UNDS"].sum().to_dict()
    carga_prensas_salida  = df_corr.dropna(subset=["SALIDA_PRENSAS"]).groupby("SALIDA_PRENSAS")["UNDS"].sum().to_dict()

    # Helpers capacidades SAL
    def get_cap_ent(date_dt, attempt):
        dkey = pd.to_datetime(date_dt).normalize()
        ov = cap_overrides_ent.get(dkey)
        if ov is not None:
            if attempt == 1 and pd.notna(ov.get("CAP1")):
                return int(ov["CAP1"])
            if attempt == 2 and pd.notna(ov.get("CAP2")):
                return int(ov["CAP2"])
        return cap_ent_1 if attempt == 1 else cap_ent_2

    def get_cap_sal(date_dt, attempt):
        dkey = pd.to_datetime(date_dt).normalize()
        ov = cap_overrides_sal.get(dkey)
        if ov is not None:
            if attempt == 1 and pd.notna(ov.get("CAP1")):
                return int(ov["CAP1"])
            if attempt == 2 and pd.notna(ov.get("CAP2")):
                return int(ov["CAP2"])
        return cap_sal_1 if attempt == 1 else cap_sal_2

    # Capacidad de estabilización por día (override si existe)
    def get_estab_cap(date_dt):
        dkey = pd.to_datetime(date_dt).normalize()
        ov = estab_cap_overrides.get(dkey)
        return ov if (ov is not None and pd.notna(ov)) else estab_cap

    # Helpers capacidades PRENSAS (por intento)
    def get_cap_prensas_ent(date_dt, attempt):
        dkey = pd.to_datetime(date_dt).normalize()
        ov = cap_overrides_prensas_ent.get(dkey)
        if ov is not None:
            if attempt == 1 and pd.notna(ov.get("CAP1")):
                return int(ov["CAP1"])
            if attempt == 2 and pd.notna(ov.get("CAP2")):
                return int(ov["CAP2"])
        return int(cap_prensas_ent_1 if attempt == 1 else cap_prensas_ent_2)

    def get_cap_prensas_sal(date_dt, attempt):
        dkey = pd.to_datetime(date_dt).normalize()
        ov = cap_overrides_prensas_sal.get(dkey)
        if ov is not None:
            if attempt == 1 and pd.notna(ov.get("CAP1")):
                return int(ov["CAP1"])
            if attempt == 2 and pd.notna(ov.get("CAP2")):
                return int(ov["CAP2"])
        return int(cap_prensas_sal_1 if attempt == 1 else cap_prensas_sal_2)

    # Chequeo estabilización
    def cabe_en_estab_rango(fecha_ini, fecha_fin_inclusive, unds):
        if pd.isna(fecha_ini) or pd.isna(fecha_fin_inclusive):
            return True
        if fecha_fin_inclusive < fecha_ini:
            return True
        for d in pd.date_range(fecha_ini, fecha_fin_inclusive, freq="D"):
            d0 = d.normalize()
            if estab_stock.get(d0, 0) + unds > get_estab_cap(d0):
                return False
        return True

    def deficits_estab(fecha_ini, fecha_fin_inclusive, unds):
        deficits = {}
        if pd.isna(fecha_ini) or pd.isna(fecha_fin_inclusive):
            return deficits
        if fecha_fin_inclusive < fecha_ini:
            return deficits
        for d in pd.date_range(fecha_ini, fecha_fin_inclusive, freq="D"):
            d0 = d.normalize()
            falta = (estab_stock.get(d0, 0) + unds) - get_estab_cap(d0)
            if falta > 0:
                deficits[d0] = int(falta)
        return deficits

    # ===============================
    # Asignación de pendientes (solo filas con ENTRADA_SAL NaN)
    # ===============================
    sugerencias_rows = []

    pendientes = df_corr[df_corr["ENTRADA_SAL"].isna()].copy()

    # --- PRIORIDAD: primero USA/MEX == "SI" ---
    def _prioritario(row):
        usa = str(row.get("USA", "NO")).strip().upper()
        mex = str(row.get("MEX", "NO")).strip().upper()
        return 1 if (usa == "SI" or mex == "SI") else 0

    pendientes["__PRIO__"] = pendientes.apply(_prioritario, axis=1)

    # Orden de trabajo: PRIO desc → DIA asc → PRODUCTO asc (si existe)
    if {"DIA", "PRODUCTO"}.issubset(pendientes.columns):
        pendientes = pendientes.sort_values(["__PRIO__", "DIA", "PRODUCTO"], ascending=[False, True, True], kind="stable")
    elif "DIA" in pendientes.columns:
        pendientes = pendientes.sort_values(["__PRIO__", "DIA"], ascending=[False, True], kind="stable")
    else:
        pendientes = pendientes.sort_values(["__PRIO__"], ascending=[False], kind="stable")

    for idx, row in pendientes.iterrows():
        dia_recepcion    = row["DIA"]
        unds             = int(row["UNDS"])
        dias_sal_optimos = int(row["DIAS_SAL_OPTIMOS"])
        prod             = str(row.get("PRODUCTO", "") or "")
        lote_id          = row.get("LOTE", idx)

        # --- +1 día si el rango es EXACTO MIN=12.0, MAX=13.0
        dias_sal_optimos_eff = dias_sal_optimos + 1 if es_rango_12_13(row) else dias_sal_optimos

        dias_max_almacen = dias_max_por_producto.get(prod, dias_max_almacen_global)
        entrada_ini = dia_recepcion if es_habil(dia_recepcion) else siguiente_habil(dia_recepcion)
        asignado = False

        for attempt in [1, 2]:
            candidatos = []
            entrada = entrada_ini
            while (entrada - dia_recepcion).days <= dias_max_almacen:
                cap_ent_dia = get_cap_ent(entrada, attempt)
                if carga_entrada.get(entrada, 0) + unds <= cap_ent_dia:
                    # Estabilización entre DIA y ENTRADA_SAL
                    if cabe_en_estab_rango(dia_recepcion, entrada - pd.Timedelta(days=1), unds):
                        # Proponer salida de sal (usar óptimos efectivos)
                        salida = entrada + timedelta(days=dias_sal_optimos_eff)
                        if ajuste_finde:
                            if salida.weekday() == 5:
                                salida = anterior_habil(salida)
                            elif salida.weekday() == 6:
                                salida = siguiente_habil(salida)
                        if ajuste_festivos and (salida.normalize() in dias_festivos):
                            dia_semana = salida.weekday()
                            if dia_semana == 0:
                                salida = siguiente_habil(salida)
                            elif dia_semana in [1, 2, 3]:
                                anterior = anterior_habil(salida)
                                siguiente = siguiente_habil(salida)
                                carga_ant  = carga_salida.get(anterior, 0)
                                carga_sig  = carga_salida.get(siguiente, 0)
                                salida = anterior if carga_ant <= carga_sig else siguiente
                            elif dia_semana == 4:
                                salida = anterior_habil(salida)

                        cap_sal_dia = get_cap_sal(salida, attempt)
                        if carga_salida.get(salida, 0) + unds <= cap_sal_dia:
                            # ---- PRS: JDOT no pasa por prensas
                            if prod.strip().upper() != "JDOT":
                                entrada_prensas = salida.normalize()  # MISMO DÍA que SALIDA_SAL
                                # Comprobar capacidad ENTRADA_PRENSAS (mismo día)
                                cap_ent_pr = get_cap_prensas_ent(entrada_prensas, attempt)
                                used_ent_pr = int(carga_prensas_entrada.get(entrada_prensas, 0))
                                if used_ent_pr + unds > cap_ent_pr:
                                    pass  # no cabe en ENTRADA_PRENSAS
                                else:
                                    # SALIDA_PRENSAS: día hábil siguiente (o +1 si no cabe)
                                    salida1 = siguiente_habil(entrada_prensas)
                                    cap1 = get_cap_prensas_sal(salida1, attempt)
                                    used1 = int(carga_prensas_salida.get(salida1, 0))
                                    if used1 + unds <= cap1:
                                        salida_prensas_final = salida1
                                    else:
                                        salida2 = siguiente_habil(salida1)
                                        cap2 = get_cap_prensas_sal(salida2, attempt)
                                        used2 = int(carga_prensas_salida.get(salida2, 0))
                                        if used2 + unds <= cap2:
                                            salida_prensas_final = salida2
                                        else:
                                            salida_prensas_final = None
                                    if salida_prensas_final is not None:
                                        dias_sal_cand = (salida - entrada).days
                                        # Para el score seguimos midiendo contra los óptimos originales
                                        diff = abs(dias_sal_cand - dias_sal_optimos)
                                        score = (diff, entrada, attempt)
                                        candidatos.append((score, entrada, salida, entrada_prensas, salida_prensas_final))
                            else:
                                # Producto JDOT: no pasa por prensas → candidato válido sin prensas
                                dias_sal_cand = (salida - entrada).days
                                diff = abs(dias_sal_cand - dias_sal_optimos)
                                score = (diff, entrada, attempt)
                                candidatos.append((score, entrada, salida, None, None))

                entrada = siguiente_habil(entrada)

            if candidatos:
                # Elegir mejor candidato (ajuste a DIAS_SAL_OPTIMOS, luego entrada temprana, luego intento)
                candidatos.sort(key=lambda t: t[0])
                _, entrada_sel, salida_sel, entrada_pr_sel, salida_pr_sel = candidatos[0]

                # Asignar SAL
                df_corr.at[idx, "ENTRADA_SAL"]      = entrada_sel
                df_corr.at[idx, "SALIDA_SAL"]       = salida_sel
                df_corr.at[idx, "DIAS_SAL"]         = (salida_sel - entrada_sel).days
                df_corr.at[idx, "DIAS_ALMACENADOS"] = (entrada_sel - dia_recepcion).days
                df_corr.at[idx, "DIFERENCIA_DIAS_SAL"] = (
                    (salida_sel - entrada_sel).days - int(row["DIAS_SAL_OPTIMOS"])
                )
                df_corr.at[idx, "LOTE_NO_ENCAJA"]   = "No"

                # Actualizar cargas SAL y estabilización
                carga_entrada[entrada_sel] = carga_entrada.get(entrada_sel, 0) + unds
                carga_salida[salida_sel]   = carga_salida.get(salida_sel, 0) + unds
                if entrada_sel.date() > dia_recepcion.date():
                    _sumar_en_rango(estab_stock, dia_recepcion, entrada_sel - pd.Timedelta(days=1), unds)

                # Asignar PRENSAS si aplica y reservar capacidad
                if prod.strip().upper() != "JDOT":
                    df_corr.at[idx, "ENTRADA_PRENSAS"] = entrada_pr_sel
                    df_corr.at[idx, "SALIDA_PRENSAS"]  = salida_pr_sel
                    if entrada_pr_sel is not None:
                        carga_prensas_entrada[entrada_pr_sel] = carga_prensas_entrada.get(entrada_pr_sel, 0) + unds
                    if salida_pr_sel is not None:
                        carga_prensas_salida[salida_pr_sel] = carga_prensas_salida.get(salida_pr_sel, 0) + unds
                else:
                    df_corr.at[idx, "ENTRADA_PRENSAS"] = pd.NaT
                    df_corr.at[idx, "SALIDA_PRENSAS"]  = pd.NaT

                asignado = True
                break

        # Si no se pudo asignar → generar sugerencias (incluye prensas)
        if not asignado:
            df_corr.at[idx, "LOTE_NO_ENCAJA"] = "Sí"

            sugerencias_rows_lote = []
            entrada = entrada_ini

            while (entrada - dia_recepcion).days <= dias_max_almacen:
                if not es_habil(entrada):
                    entrada = siguiente_habil(entrada)
                    continue

                for attempt in [1, 2]:
                    cap_ent_dia = get_cap_ent(entrada, attempt)
                    deficit_ent = max(0, (carga_entrada.get(entrada, 0) + unds) - cap_ent_dia)

                    def_est = deficits_estab(dia_recepcion, entrada - pd.Timedelta(days=1), unds)
                    deficit_estab_max = max(def_est.values()) if def_est else 0

                    # Usar óptimos efectivos
                    dias_sal_optimos_eff = dias_sal_optimos + 1 if es_rango_12_13(row) else dias_sal_optimos

                    salida = entrada + timedelta(days=dias_sal_optimos_eff)
                    if ajuste_finde:
                        if salida.weekday() == 5:
                            salida = anterior_habil(salida)
                        elif salida.weekday() == 6:
                            salida = siguiente_habil(salida)
                    if ajuste_festivos and (salida.normalize() in dias_festivos):
                        dia_semana = salida.weekday()
                        if dia_semana == 0:
                            salida = siguiente_habil(salida)
                        elif dia_semana in [1, 2, 3]:
                            anterior = anterior_habil(salida)
                            siguiente = siguiente_habil(salida)
                            carga_ant = carga_salida.get(anterior, 0)
                            carga_sig = carga_salida.get(siguiente, 0)
                            salida = anterior if carga_ant <= carga_sig else siguiente
                        elif dia_semana == 4:
                            salida = anterior_habil(salida)

                    cap_sal_dia = get_cap_sal(salida, attempt)
                    deficit_sal = max(0, (carga_salida.get(salida, 0) + unds) - cap_sal_dia)

                    # ---- Déficits y propuesta en PRENSAS (solo si no es JDOT)
                    deficit_ent_pr = 0
                    deficit_sal_pr = 0
                    entrada_pr_prop = pd.NaT
                    salida_pr_prop = pd.NaT

                    if prod.strip().upper() != "JDOT":
                        entrada_pr = pd.to_datetime(salida).normalize()  # propuesta: mismo día que SALIDA_SAL
                        entrada_pr_prop = entrada_pr

                        # ENTRADA_PRENSAS (con intentos)
                        cap_ent_pr = get_cap_prensas_ent(entrada_pr, attempt)
                        used_ent_pr = int(carga_prensas_entrada.get(entrada_pr, 0))
                        deficit_ent_pr = max(0, (used_ent_pr + unds) - cap_ent_pr)

                        # SALIDA_PRENSAS: día hábil siguiente (o +1 si no cabe) → escoger la que tenga MENOR déficit
                        salida1 = siguiente_habil(entrada_pr)
                        cap1 = get_cap_prensas_sal(salida1, attempt); used1 = int(carga_prensas_salida.get(salida1, 0))
                        deficit1 = max(0, (used1 + unds) - cap1)

                        salida2 = siguiente_habil(salida1)
                        cap2 = get_cap_prensas_sal(salida2, attempt); used2 = int(carga_prensas_salida.get(salida2, 0))
                        deficit2 = max(0, (used2 + unds) - cap2)

                        if deficit1 <= deficit2:
                            salida_pr_prop = salida1
                            deficit_sal_pr = deficit1
                        else:
                            salida_pr_prop = salida2
                            deficit_sal_pr = deficit2

                    # Generar texto de recomendación
                    recomendaciones = []
                    if deficit_ent > 0:
                        recomendaciones.append(
                            f"Subir ENTRADA_SAL el {entrada.normalize().date()} en +{int(deficit_ent)} unds (INTENTO {attempt})."
                        )
                    if deficit_sal > 0:
                        recomendaciones.append(
                            f"Subir SALIDA_SAL el {salida.normalize().date()} en +{int(deficit_sal)} unds (INTENTO {attempt})."
                        )
                    if deficit_estab_max > 0:
                        dias_estab = [f"{k.date()}(+{v})" for k, v in list(def_est.items())[:3] if v > 0]
                        if dias_estab:
                            recomendaciones.append("Subir ESTABILIZACIÓN en: " + ", ".join(dias_estab))
                    if prod.strip().upper() != "JDOT":
                        if deficit_ent_pr > 0:
                            recomendaciones.append(
                                f"Subir ENTRADA_PRENSAS el {entrada_pr_prop.date()} en +{int(deficit_ent_pr)} unds."
                            )
                        if deficit_sal_pr > 0:
                            recomendaciones.append(
                                f"Subir SALIDA_PRENSAS ({salida_pr_prop.date()}) en +{int(deficit_sal_pr)} unds."
                            )

                    sugerencias_rows_lote.append({
                        "LOTE": lote_id,
                        "PRODUCTO": prod,
                        "UNDS": unds,
                        "DIA_RECEPCION": pd.to_datetime(dia_recepcion).normalize(),
                        "ENTRADA_PROPUESTA": pd.to_datetime(entrada).normalize(),
                        "SALIDA_PROPUESTA": pd.to_datetime(salida).normalize(),
                        "ENTRADA_PROPUESTA_PRENSAS": pd.to_datetime(entrada_pr_prop) if pd.notna(entrada_pr_prop) else pd.NaT,
                        "SALIDA_PROPUESTA_PRENSAS": pd.to_datetime(salida_pr_prop) if pd.notna(salida_pr_prop) else pd.NaT,
                        "INTENTO": attempt,
                        "DEFICIT_ENTRADA": int(deficit_ent),
                        "DEFICIT_ESTAB_MAX": int(deficit_estab_max),
                        "DEFICIT_SALIDA": int(deficit_sal),
                        "DEFICIT_ENTRADA_PRENSAS": int(deficit_ent_pr),
                        "DEFICIT_SALIDA_PRENSAS": int(deficit_sal_pr),
                        "MAX_DEFICIT": int(max(deficit_ent, deficit_estab_max, deficit_sal, deficit_ent_pr, deficit_sal_pr)),
                        "TOTAL_DEFICIT": int(deficit_ent + deficit_estab_max + deficit_sal + deficit_ent_pr + deficit_sal_pr),
                        "RECOMENDACION": " | ".join(recomendaciones) if recomendaciones else "Sin ajustes necesarios"
                    })

                entrada = siguiente_habil(entrada)

            if sugerencias_rows_lote:
                sugerencias_rows_lote.sort(
                    key=lambda r: (r["MAX_DEFICIT"], r["TOTAL_DEFICIT"], r["ENTRADA_PROPUESTA"])
                )
                sugerencias_rows.extend(sugerencias_rows_lote[:20])

    # Sugerencias DF
    cols_sug = [
        "LOTE", "PRODUCTO", "UNDS", "DIA_RECEPCION",
        "ENTRADA_PROPUESTA", "SALIDA_PROPUESTA",
        "ENTRADA_PROPUESTA_PRENSAS", "SALIDA_PROPUESTA_PRENSAS",
        "INTENTO",
        "DEFICIT_ENTRADA", "DEFICIT_ESTAB_MAX", "DEFICIT_SALIDA",
        "DEFICIT_ENTRADA_PRENSAS", "DEFICIT_SALIDA_PRENSAS",
        "MAX_DEFICIT", "TOTAL_DEFICIT", "RECOMENDACION"
    ]
    df_sugerencias = pd.DataFrame(sugerencias_rows, columns=cols_sug) if sugerencias_rows else pd.DataFrame(columns=cols_sug)
    if not df_sugerencias.empty:
        df_sugerencias = df_sugerencias.sort_values(
            by=["MAX_DEFICIT", "TOTAL_DEFICIT", "ENTRADA_PROPUESTA", "SALIDA_PROPUESTA", "LOTE"],
            ascending=[True, True, True, True, True]
        ).reset_index(drop=True)

    return df_corr, df_sugerencias
    
# -------------------------------
# Ejecución de la app
# -------------------------------
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    alias_map = {
        "DIAS SAL OPTIMOS": "DIAS_SAL_OPTIMOS",
        "DIAS_SAL_OPTIMOS": "DIAS_SAL_OPTIMOS",
        "ENTRADA SAL": "ENTRADA_SAL",
        "SALIDA SAL": "SALIDA_SAL"
    }
    for a, target in alias_map.items():
        if a in df.columns and target not in df.columns:
            df.rename(columns={a: target}, inplace=True)

    for col in ["DIA", "ENTRADA_SAL", "SALIDA_SAL", "ENTRADA_PRENSAS", "SALIDA_PRENSAS"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "UNDS" in df.columns:
        df["UNDS"] = pd.to_numeric(df["UNDS"], errors="coerce").fillna(0).astype(int)

    # --- Normaliza columnas USA/MEX (SI/NO) ---
    for col_flag in ["USA", "MEX"]:
        if col_flag not in df.columns:
            df[col_flag] = "NO"  # por defecto si no existe
        df[col_flag] = (
            df[col_flag]
            .astype(str)
            .str.strip()
            .str.upper()
            .replace({"SÍ": "SI"})
        )
        df.loc[~df[col_flag].isin(["SI", "NO"]), col_flag] = "NO"

    # --- Normaliza MIN_PESO / MAX_PESO a float (solo por columnas numéricas se detecta el rango 12-13) ---
    for colw in ["MIN_PESO", "MAX_PESO"]:
        if colw in df.columns:
            df[colw] = df[colw].apply(_to_float_or_nan)

    # Overrides por PRODUCTO
    dias_max_por_producto = {}
    if "PRODUCTO" in df.columns:
        productos = sorted(df["PRODUCTO"].dropna().astype(str).unique().tolist())
        st.sidebar.markdown("### ⏱️ Días máx. almacenamiento por PRODUCTO")

        if "overrides_df" not in st.session_state or set(st.session_state.get("productos_cache", [])) != set(productos):
            st.session_state.overrides_df = pd.DataFrame({
                "PRODUCTO": productos,
                "DIAS_MAX_ALMACEN": [dias_max_almacen_global] * len(productos)
            })
            st.session_state.productos_cache = productos

        overrides_df = st.sidebar.data_editor(
            st.session_state.overrides_df,
            use_container_width=True,
            num_rows="dynamic",
            disabled=["PRODUCTO"],
            column_config={
                "PRODUCTO": st.column_config.TextColumn("PRODUCTO"),
                "DIAS_MAX_ALMACEN": st.column_config.NumberColumn("Días máx. naturales", step=1, min_value=0)
            },
            key="overrides_editor"
        )
        if not overrides_df.empty:
            dias_max_por_producto = dict(zip(overrides_df["PRODUCTO"], overrides_df["DIAS_MAX_ALMACEN"]))
    else:
        st.sidebar.info("No se encontró columna PRODUCTO. Se aplicará solo el límite GLOBAL.")

    # Overrides capacidad por fecha: SAL (CAP1/CAP2)
    st.sidebar.markdown("### 📅 Overrides capacidad ENTRADA SAL (opcional)")
    if "cap_overrides_ent_df" not in st.session_state:
        st.session_state.cap_overrides_ent_df = pd.DataFrame({
            "FECHA": pd.to_datetime(pd.Series([], dtype="datetime64[ns]")),
            "CAP1":  pd.Series([], dtype="Int64"),
            "CAP2":  pd.Series([], dtype="Int64"),
        })
    st.session_state.cap_overrides_ent_df["FECHA"] = pd.to_datetime(st.session_state.cap_overrides_ent_df["FECHA"], errors="coerce")
    for c in ("CAP1", "CAP2"):
        st.session_state.cap_overrides_ent_df[c] = pd.to_numeric(st.session_state.cap_overrides_ent_df[c], errors="coerce").astype("Int64")
    cap_overrides_ent_df = st.sidebar.data_editor(
        st.session_state.cap_overrides_ent_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha (entrada)", format="YYYY-MM-DD"),
            "CAP1": st.column_config.NumberColumn("Capacidad 1º intento", step=50, min_value=0),
            "CAP2": st.column_config.NumberColumn("Capacidad 2º intento", step=50, min_value=0),
        },
        key="cap_overrides_ent_editor"
    )

    st.sidebar.markdown("### 📅 Overrides capacidad SALIDA SAL (opcional)")
    if "cap_overrides_sal_df" not in st.session_state:
        st.session_state.cap_overrides_sal_df = pd.DataFrame({
            "FECHA": pd.to_datetime(pd.Series([], dtype="datetime64[ns]")),
            "CAP1":  pd.Series([], dtype="Int64"),
            "CAP2":  pd.Series([], dtype="Int64"),
        })
    st.session_state.cap_overrides_sal_df["FECHA"] = pd.to_datetime(st.session_state.cap_overrides_sal_df["FECHA"], errors="coerce")
    for c in ("CAP1", "CAP2"):
        st.session_state.cap_overrides_sal_df[c] = pd.to_numeric(st.session_state.cap_overrides_sal_df[c], errors="coerce").astype("Int64")
    cap_overrides_sal_df = st.sidebar.data_editor(
        st.session_state.cap_overrides_sal_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha (salida)", format="YYYY-MM-DD"),
            "CAP1": st.column_config.NumberColumn("Capacidad 1º intento", step=50, min_value=0),
            "CAP2": st.column_config.NumberColumn("Capacidad 2º intento", step=50, min_value=0),
        },
        key="cap_overrides_sal_editor"
    )

    # Overrides capacidad por fecha: ESTABILIZACIÓN
    st.sidebar.markdown("### 📅 Overrides capacidad ESTABILIZACIÓN (opcional)")
    if "cap_overrides_estab_df" not in st.session_state:
        st.session_state.cap_overrides_estab_df = pd.DataFrame({
            "FECHA": pd.to_datetime(pd.Series([], dtype="datetime64[ns]")),
            "CAP":   pd.Series([], dtype="Int64"),
        })
    st.session_state.cap_overrides_estab_df["FECHA"] = pd.to_datetime(st.session_state.cap_overrides_estab_df["FECHA"], errors="coerce")
    st.session_state.cap_overrides_estab_df["CAP"] = pd.to_numeric(st.session_state.cap_overrides_estab_df["CAP"], errors="coerce").astype("Int64")
    cap_overrides_estab_df = st.sidebar.data_editor(
        st.session_state.cap_overrides_estab_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha (estabilización)", format="YYYY-MM-DD"),
            "CAP":   st.column_config.NumberColumn("Capacidad estabilización (unds)", step=50, min_value=0),
        },
        key="cap_overrides_estab_editor"
    )

    # Overrides capacidad por fecha: PRENSAS (CAP1/CAP2)
    st.sidebar.markdown("### 📅 Overrides capacidad ENTRADA PRENSAS (opcional)")
    if "cap_overrides_prensas_ent_df" not in st.session_state:
        st.session_state.cap_overrides_prensas_ent_df = pd.DataFrame({
            "FECHA": pd.to_datetime(pd.Series([], dtype="datetime64[ns]")),
            "CAP1":  pd.Series([], dtype="Int64"),
            "CAP2":  pd.Series([], dtype="Int64"),
        })
    st.session_state.cap_overrides_prensas_ent_df["FECHA"] = pd.to_datetime(st.session_state.cap_overrides_prensas_ent_df["FECHA"], errors="coerce")
    for c in ("CAP1", "CAP2"):
        st.session_state.cap_overrides_prensas_ent_df[c] = pd.to_numeric(st.session_state.cap_overrides_prensas_ent_df[c], errors="coerce").astype("Int64")
    cap_overrides_prensas_ent_df = st.sidebar.data_editor(
        st.session_state.cap_overrides_prensas_ent_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha (ENTRADA_PRENSAS)", format="YYYY-MM-DD"),
            "CAP1":  st.column_config.NumberColumn("Capacidad 1º intento", step=50, min_value=0),
            "CAP2":  st.column_config.NumberColumn("Capacidad 2º intento", step=50, min_value=0),
        },
        key="cap_overrides_prensas_ent_editor"
    )

    st.sidebar.markdown("### 📅 Overrides capacidad SALIDA PRENSAS (opcional)")
    if "cap_overrides_prensas_sal_df" not in st.session_state:
        st.session_state.cap_overrides_prensas_sal_df = pd.DataFrame({
            "FECHA": pd.to_datetime(pd.Series([], dtype="datetime64[ns]")),
            "CAP1":  pd.Series([], dtype="Int64"),
            "CAP2":  pd.Series([], dtype="Int64"),
        })
    st.session_state.cap_overrides_prensas_sal_df["FECHA"] = pd.to_datetime(st.session_state.cap_overrides_prensas_sal_df["FECHA"], errors="coerce")
    for c in ("CAP1", "CAP2"):
        st.session_state.cap_overrides_prensas_sal_df[c] = pd.to_numeric(st.session_state.cap_overrides_prensas_sal_df[c], errors="coerce").astype("Int64")
    cap_overrides_prensas_sal_df = st.sidebar.data_editor(
        st.session_state.cap_overrides_prensas_sal_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha (SALIDA_PRENSAS)", format="YYYY-MM-DD"),
            "CAP1":  st.column_config.NumberColumn("Capacidad 1º intento", step=50, min_value=0),
            "CAP2":  st.column_config.NumberColumn("Capacidad 2º intento", step=50, min_value=0),
        },
        key="cap_overrides_prensas_sal_editor"
    )

    # Normaliza overrides → dicts por fecha
    cap_overrides_ent = {}
    if not cap_overrides_ent_df.empty:
        tmp = cap_overrides_ent_df.dropna(subset=["FECHA"]).copy()
        tmp["FECHA"] = pd.to_datetime(tmp["FECHA"]).dt.normalize()
        for _, r in tmp.iterrows():
            cap_overrides_ent[r["FECHA"]] = {
                "CAP1": (int(r["CAP1"]) if pd.notna(r["CAP1"]) else None),
                "CAP2": (int(r["CAP2"]) if pd.notna(r["CAP2"]) else None),
            }
    st.session_state.cap_overrides_ent_df = cap_overrides_ent_df

    cap_overrides_sal = {}
    if not cap_overrides_sal_df.empty:
        tmp2 = cap_overrides_sal_df.dropna(subset=["FECHA"]).copy()
        tmp2["FECHA"] = pd.to_datetime(tmp2["FECHA"]).dt.normalize()
        for _, r in tmp2.iterrows():
            cap_overrides_sal[r["FECHA"]] = {
                "CAP1": (int(r["CAP1"]) if pd.notna(r["CAP1"]) else None),
                "CAP2": (int(r["CAP2"]) if pd.notna(r["CAP2"]) else None),
            }
    st.session_state.cap_overrides_sal_df = cap_overrides_sal_df

    estab_cap_overrides = {}
    if not cap_overrides_estab_df.empty:
        tmp3 = cap_overrides_estab_df.dropna(subset=["FECHA"]).copy()
        tmp3["FECHA"] = pd.to_datetime(tmp3["FECHA"]).dt.normalize()
        for _, r in tmp3.iterrows():
            if pd.notna(r["CAP"]):
                estab_cap_overrides[r["FECHA"]] = int(r["CAP"])
    st.session_state.cap_overrides_estab_df = cap_overrides_estab_df

    cap_overrides_prensas_ent = {}
    if not cap_overrides_prensas_ent_df.empty:
        tmp4 = cap_overrides_prensas_ent_df.dropna(subset=["FECHA"]).copy()
        tmp4["FECHA"] = pd.to_datetime(tmp4["FECHA"]).dt.normalize()
        for _, r in tmp4.iterrows():
            cap_overrides_prensas_ent[r["FECHA"]] = {
                "CAP1": (int(r["CAP1"]) if pd.notna(r["CAP1"]) else None),
                "CAP2": (int(r["CAP2"]) if pd.notna(r["CAP2"]) else None),
            }
    st.session_state.cap_overrides_prensas_ent_df = cap_overrides_prensas_ent_df

    cap_overrides_prensas_sal = {}
    if not cap_overrides_prensas_sal_df.empty:
        tmp5 = cap_overrides_prensas_sal_df.dropna(subset=["FECHA"]).copy()
        tmp5["FECHA"] = pd.to_datetime(tmp5["FECHA"]).dt.normalize()
        for _, r in tmp5.iterrows():
            cap_overrides_prensas_sal[r["FECHA"]] = {
                "CAP1": (int(r["CAP1"]) if pd.notna(r["CAP1"]) else None),
                "CAP2": (int(r["CAP2"]) if pd.notna(r["CAP2"]) else None),
            }
    st.session_state.cap_overrides_prensas_sal_df = cap_overrides_prensas_sal_df

    # ===============================
    # 🔧 Modo de planificación
    # ===============================
    st.markdown("### ⚙️ Modo de planificación")
    usar_plan_actual = st.toggle(
        "Usar planificación actual como base (no tocar lo ya planificado)",
        value=True,
        help="Si está activo, se parte de la planificación guardada en la sesión. Solo se intentan los lotes seleccionados (por defecto, los que no encajan o están sin ENTRADA)."
    )

    if usar_plan_actual and ("df_planificado" in st.session_state):
        df_base = st.session_state["df_planificado"].copy()
    else:
        df_base = df.copy()

    candidatos_mask = df_base["ENTRADA_SAL"].isna()
    if "LOTE_NO_ENCAJA" in df_base.columns:
        candidatos_mask = candidatos_mask | (df_base["LOTE_NO_ENCAJA"].astype(str).str.upper() == "SÍ")

    candidatos_df = df_base[candidatos_mask].copy()

    lotes_candidatos = candidatos_df["LOTE"].astype(str).tolist() if "LOTE" in candidatos_df.columns else candidatos_df.index.astype(str).tolist()
    lotes_select = st.multiselect(
        "Elige qué lotes quieres replanificar (solo estos se modificarán):",
        options=lotes_candidatos,
        default=lotes_candidatos,
        help="Por defecto se incluyen los lotes sin ENTRADA o con LOTE_NO_ENCAJA='Sí'."
    )

    if "LOTE" in df_base.columns:
        idx_a_replan = df_base[df_base["LOTE"].astype(str).isin(lotes_select)].index
    else:
        idx_a_replan = df_base.index[df_base.index.astype(str).isin(lotes_select)]

    df_trabajo = df_base.copy()

    # Liberar SOLO las filas seleccionadas preservando tipos
    datetime_cols = [c for c in ["ENTRADA_SAL", "SALIDA_SAL", "ENTRADA_PRENSAS", "SALIDA_PRENSAS"] if c in df_trabajo.columns]
    numeric_cols  = [c for c in ["DIAS_SAL", "DIAS_ALMACENADOS", "DIFERENCIA_DIAS_SAL"] if c in df_trabajo.columns]
    text_cols     = [c for c in ["LOTE_NO_ENCAJA"] if c in df_trabajo.columns]

    if datetime_cols:
        df_trabajo.loc[idx_a_replan, datetime_cols] = pd.NaT
    for c in numeric_cols:
        df_trabajo.loc[idx_a_replan, c] = pd.NA
    for c in text_cols:
        df_trabajo.loc[idx_a_replan, c] = pd.NA

    for c in datetime_cols:
        df_trabajo[c] = pd.to_datetime(df_trabajo[c], errors="coerce")
    for c in numeric_cols:
        df_trabajo[c] = pd.to_numeric(df_trabajo[c], errors="coerce").astype("Int64")

    # Botón de planificación incremental
    if st.button("🚀 Aplicar planificación (solo lotes seleccionados)"):
        df_planificado, df_sugerencias = planificar_filas_na(
            df_trabajo, dias_max_almacen_global, dias_max_por_producto,
            estab_cap, cap_overrides_ent, cap_overrides_sal, estab_cap_overrides,
            cap_prensas_ent_1, cap_prensas_ent_2,
            cap_prensas_sal_1, cap_prensas_sal_2,
            cap_overrides_prensas_ent, cap_overrides_prensas_sal
        )
        st.session_state["df_planificado"] = df_planificado
        st.session_state["df_sugerencias"] = df_sugerencias
        st.success(f"✅ Replanificación aplicada a {len(idx_a_replan)} lote(s). El resto no se ha modificado.")

    # ===============================
    # Mostrar tabla editable, gráficos y estabilización (fuera del botón)
    # ===============================
    if "df_planificado" in st.session_state:
        df_show = st.session_state["df_planificado"]

        with st.expander("🧪 Diagnóstico dtypes", expanded=False):
            st.write(df_show.dtypes.astype(str))

        column_config = {}
        for col in df_show.columns:
            s = df_show[col]
            try:
                if pd.api.types.is_datetime64_any_dtype(s):
                    column_config[col] = st.column_config.DateColumn(col, format="YYYY-MM-DD", disabled=False)
                elif pd.api.types.is_integer_dtype(s) or pd.api.types.is_float_dtype(s):
                    column_config[col] = st.column_config.NumberColumn(col, disabled=False)
                else:
                    column_config[col] = st.column_config.TextColumn(col)
            except Exception:
                column_config[col] = st.column_config.TextColumn(col)

        df_for_editor = df_show.copy()
        column_config2 = dict(column_config)

        if "LOTE_NO_ENCAJA" in df_for_editor.columns:
            valnorm = (
                df_for_editor["LOTE_NO_ENCAJA"]
                .astype(str).str.strip().str.upper().str.replace("Í", "I", regex=False)
            )
            df_for_editor["🚨"] = valnorm.isin(["SI"]).map({True: "❌", False: ""})
            cols = ["🚨"] + [c for c in df_for_editor.columns if c != "🚨"]
            df_for_editor = df_for_editor[cols]
            column_config2["🚨"] = st.column_config.TextColumn("🚨", width="small", help="No encaja")

        df_editable = st.data_editor(
            df_for_editor,
            column_config=column_config2,
            num_rows="dynamic",
            use_container_width=True,
            key="plan_editor"
        )

        # -------------------------------
        # Gráfico: Entradas/Salidas SAL por lote/fecha
        # -------------------------------
        st.subheader("📊 Entradas y salidas de SAL por fecha con detalle por lote")

        fig = go.Figure()

        df_e = df_editable.dropna(subset=["ENTRADA_SAL", "UNDS"]) if "ENTRADA_SAL" in df_editable.columns else pd.DataFrame()
        df_s = df_editable.dropna(subset=["SALIDA_SAL", "UNDS"]) if "SALIDA_SAL" in df_editable.columns else pd.DataFrame()

        pivot_e = (
            df_e.groupby(["ENTRADA_SAL", "LOTE"])["UNDS"].sum().unstack(fill_value=0).sort_index()
            if not df_e.empty and {"ENTRADA_SAL", "LOTE", "UNDS"}.issubset(df_e.columns)
            else pd.DataFrame()
        )
        pivot_s = (
            df_s.groupby(["SALIDA_SAL", "LOTE"])["UNDS"].sum().unstack(fill_value=0).sort_index()
            if not df_s.empty and {"SALIDA_SAL", "LOTE", "UNDS"}.issubset(df_s.columns)
            else pd.DataFrame()
        )

        if not pivot_e.empty:
            for lote in pivot_e.columns:
                y_vals = pivot_e[lote]
                if (y_vals > 0).any():
                    fig.add_trace(go.Bar(
                        x=pivot_e.index, y=y_vals, name=f"Lote {lote}",
                        offsetgroup="entrada", legendgroup=f"lote-{lote}",
                        marker_color="blue", marker_line_color="white", marker_line_width=1.2,
                        hovertemplate="Fecha: %{x|%Y-%m-%d}<br>Lote: " + str(lote) + "<br>UNDS: %{y}<extra></extra>",
                        showlegend=True
                    ))

        if not pivot_s.empty:
            for lote in pivot_s.columns:
                y_vals = pivot_s[lote]
                if (y_vals > 0).any():
                    fig.add_trace(go.Bar(
                        x=pivot_s.index, y=y_vals, name=f"Lote {lote} (Salida)",
                        offsetgroup="salida", legendgroup=f"lote-{lote}",
                        marker_color="orange", marker_line_color="white", marker_line_width=1.2,
                        hovertemplate="Fecha: %{x|%Y-%m-%d}<br>Lote: " + str(lote) + "<br>UNDS: %{y}<extra></extra>",
                        showlegend=False
                    ))

        label_shift = pd.Timedelta(hours=8)
        annotations = []

        tot_e = pd.DataFrame()
        tot_s = pd.DataFrame()
        if not df_e.empty:
            if "LOTE" in df_e.columns:
                tot_e = df_e.groupby("ENTRADA_SAL").agg(UNDS=("UNDS","sum"), LOTES=("LOTE","nunique")).reset_index()
            else:
                tot_e = df_e.groupby("ENTRADA_SAL").agg(UNDS=("UNDS","sum"), LOTES=("UNDS","size")).reset_index()
        if not df_s.empty:
            if "LOTE" in df_s.columns:
                tot_s = df_s.groupby("SALIDA_SAL").agg(UNDS=("UNDS","sum"), LOTES=("LOTE","nunique")).reset_index()
            else:
                tot_s = df_s.groupby("SALIDA_SAL").agg(UNDS=("UNDS","sum"), LOTES=("UNDS","size")).reset_index()

        max_e = int(tot_e["UNDS"].max()) if not tot_e.empty else 0
        max_s = int(tot_s["UNDS"].max()) if not tot_s.empty else 0
        max_y = max(max_e, max_s) or 1

        def add_two_labels(x_dt, y_val, lots_count, is_entry=True):
            x_pos = x_dt - label_shift if is_entry else x_dt + label_shift
            y_base = max(y_val, max_y * 0.02)
            annotations.append(dict(x=x_pos, y=y_base, xref="x", yref="y",
                                    text=f"<b>{int(y_val)}</b>", showarrow=False, yshift=28,
                                    align="center", font=dict(size=13, color="black")))
            annotations.append(dict(x=x_pos, y=y_base, xref="x", yref="y",
                                    text=f"{int(lots_count)} lotes", showarrow=False, yshift=12,
                                    align="center", font=dict(size=11, color="gray")))

        if not tot_e.empty:
            for _, r in tot_e.iterrows():
                add_two_labels(r["ENTRADA_SAL"], r["UNDS"], r["LOTES"], is_entry=True)
        if not tot_s.empty:
            for _, r in tot_s.iterrows():
                add_two_labels(r["SALIDA_SAL"], r["UNDS"], r["LOTES"], is_entry=False)

        ticks = pd.Index(sorted(set(
            (pivot_e.index.tolist() if not pivot_e.empty else []) +
            (pivot_s.index.tolist() if not pivot_s.empty else [])
        )))
        fig.update_layout(
            barmode="relative",
            xaxis_title="Fecha",
            yaxis_title="Unidades",
            xaxis=dict(tickmode="array", tickvals=ticks, tickformat="%d %b (%a)"),
            bargap=0.25, bargroupgap=0.12, annotations=annotations,
            legend=dict(itemclick="toggleothers", itemdoubleclick="toggle", groupclick="togglegroup")
        )
        fig.update_yaxes(range=[0, max_y * 1.25])
        st.plotly_chart(fig, use_container_width=True)

        # -------------------------------
        # Gráfico: Entradas/Salidas PRENSAS por lote/fecha
        # -------------------------------
        st.subheader("📊 Entradas y salidas de PRENSAS por fecha con detalle por lote")

        figp = go.Figure()

        df_pe = df_editable.dropna(subset=["ENTRADA_PRENSAS", "UNDS"]) if "ENTRADA_PRENSAS" in df_editable.columns else pd.DataFrame()
        df_ps = df_editable.dropna(subset=["SALIDA_PRENSAS", "UNDS"]) if "SALIDA_PRENSAS" in df_editable.columns else pd.DataFrame()

        pivot_pe = (
            df_pe.groupby(["ENTRADA_PRENSAS", "LOTE"])["UNDS"].sum().unstack(fill_value=0).sort_index()
            if not df_pe.empty and {"ENTRADA_PRENSAS", "LOTE", "UNDS"}.issubset(df_pe.columns)
            else pd.DataFrame()
        )
        pivot_ps = (
            df_ps.groupby(["SALIDA_PRENSAS", "LOTE"])["UNDS"].sum().unstack(fill_value=0).sort_index()
            if not df_ps.empty and {"SALIDA_PRENSAS", "LOTE", "UNDS"}.issubset(df_ps.columns)
            else pd.DataFrame()
        )

        if not pivot_pe.empty:
            for lote in pivot_pe.columns:
                y_vals = pivot_pe[lote]
                if (y_vals > 0).any():
                    figp.add_trace(go.Bar(
                        x=pivot_pe.index, y=y_vals, name=f"Lote {lote}",
                        offsetgroup="entrada_pr", legendgroup=f"plote-{lote}",
                        marker_color="blue", marker_line_color="white", marker_line_width=1.2,
                        hovertemplate="Fecha: %{x|%Y-%m-%d}<br>Lote: " + str(lote) + "<br>UNDS: %{y}<extra></extra>",
                        showlegend=True
                    ))

        if not pivot_ps.empty:
            for lote in pivot_ps.columns:
                y_vals = pivot_ps[lote]
                if (y_vals > 0).any():
                    figp.add_trace(go.Bar(
                        x=pivot_ps.index, y=y_vals, name=f"Lote {lote} (Salida prensas)",
                        offsetgroup="salida_pr", legendgroup=f"plote-{lote}",
                        marker_color="orange", marker_line_color="white", marker_line_width=1.2,
                        hovertemplate="Fecha: %{x|%Y-%m-%d}<br>Lote: " + str(lote) + "<br>UNDS: %{y}<extra></extra>",
                        showlegend=False
                    ))

        annotations_p = []
        tot_pe = pd.DataFrame()
        tot_ps = pd.DataFrame()
        if not df_pe.empty:
            if "LOTE" in df_pe.columns:
                tot_pe = df_pe.groupby("ENTRADA_PRENSAS").agg(UNDS=("UNDS","sum"), LOTES=("LOTE","nunique")).reset_index()
            else:
                tot_pe = df_pe.groupby("ENTRADA_PRENSAS").agg(UNDS=("UNDS","sum"), LOTES=("UNDS","size")).reset_index()
        if not df_ps.empty:
            if "LOTE" in df_ps.columns:
                tot_ps = df_ps.groupby("SALIDA_PRENSAS").agg(UNDS=("UNDS","sum"), LOTES=("LOTE","nunique")).reset_index()
            else:
                tot_ps = df_ps.groupby("SALIDA_PRENSAS").agg(UNDS=("UNDS","sum"), LOTES=("UNDS","size")).reset_index()

        max_pe = int(tot_pe["UNDS"].max()) if not tot_pe.empty else 0
        max_ps = int(tot_ps["UNDS"].max()) if not tot_ps.empty else 0
        max_y_p = max(max_pe, max_ps) or 1

        def add_two_labels_p(x_dt, y_val, lots_count, is_entry=True):
            x_pos = x_dt - label_shift if is_entry else x_dt + label_shift
            y_base = max(y_val, max_y_p * 0.02)
            annotations_p.append(dict(x=x_pos, y=y_base, xref="x", yref="y",
                                      text=f"<b>{int(y_val)}</b>", showarrow=False, yshift=28,
                                      align="center", font=dict(size=13, color="black")))
            annotations_p.append(dict(x=x_pos, y=y_base, xref="x", yref="y",
                                      text=f"{int(lots_count)} lotes", showarrow=False, yshift=12,
                                      align="center", font=dict(size=11, color="gray")))

        if not tot_pe.empty:
            for _, r in tot_pe.iterrows():
                add_two_labels_p(r["ENTRADA_PRENSAS"], r["UNDS"], r["LOTES"], is_entry=True)
        if not tot_ps.empty:
            for _, r in tot_ps.iterrows():
                add_two_labels_p(r["SALIDA_PRENSAS"], r["UNDS"], r["LOTES"], is_entry=False)

        ticks_p = pd.Index(sorted(set(
            (pivot_pe.index.tolist() if not pivot_pe.empty else []) +
            (pivot_ps.index.tolist() if not pivot_ps.empty else [])
        )))
        figp.update_layout(
            barmode="relative",
            xaxis_title="Fecha",
            yaxis_title="Unidades",
            xaxis=dict(tickmode="array", tickvals=ticks_p, tickformat="%d %b (%a)"),
            bargap=0.25, bargroupgap=0.12, annotations=annotations_p,
            legend=dict(itemclick="toggleothers", itemdoubleclick="toggle", groupclick="togglegroup")
        )
        figp.update_yaxes(range=[0, max_y_p * 1.25])
        st.plotly_chart(figp, use_container_width=True)

        # ===============================
        # 📦 Estabilización
        # ===============================
        df_estab = calcular_estabilizacion_diaria(df_editable, estab_cap, estab_cap_overrides)
        with st.expander("📦 Ocupación diaria de cámara de estabilización", expanded=True):
            if df_estab.empty:
                st.info("No hay días con stock en estabilización.")
            else:
                st.dataframe(df_estab, use_container_width=True, hide_index=True)
                colores = df_estab.apply(
                    lambda r: "crimson" if r["ESTAB_UNDS"] > r["CAPACIDAD"] else "teal",
                    axis=1
                )
                fig_est = go.Figure()
                fig_est.add_trace(go.Bar(x=df_estab["FECHA"], y=df_estab["ESTAB_UNDS"],
                                         marker_color=colores,
                                         hovertemplate="Fecha: %{x|%Y-%m-%d}<br>Unds: %{y}<extra></extra>",
                                         showlegend=False))
                fig_est.add_trace(go.Scatter(x=df_estab["FECHA"], y=df_estab["ESTAB_UNDS"],
                                             mode="text",
                                             text=[str(int(v)) for v in df_estab["ESTAB_UNDS"]],
                                             textposition="top center", showlegend=False))
                fig_est.add_hline(y=estab_cap, line_dash="dash", line_color="orange",
                                  annotation_text=f"Capacidad: {estab_cap}",
                                  annotation_position="top left")
                fig_est.update_layout(xaxis_title="Fecha", yaxis_title="Unidades en estabilización",
                                      bargap=0.25, showlegend=False,
                                      xaxis=dict(tickmode="array", tickvals=df_estab["FECHA"],
                                                 tickformat="%d %b (%a)"))
                st.plotly_chart(fig_est, use_container_width=True)

                estab_xlsx = generar_excel(df_estab, "estabilizacion_diaria.xlsx")
                st.download_button(
                    "💾 Descargar estabilización (Excel)",
                    data=estab_xlsx,
                    file_name="estabilizacion_diaria.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ===============================
        # 📌 Sugerencias
        # ===============================
        if "df_sugerencias" in st.session_state:
            df_sug = st.session_state["df_sugerencias"]
        else:
            _, df_sug = planificar_filas_na(
                df_show, dias_max_almacen_global, dias_max_por_producto,
                estab_cap, cap_overrides_ent, cap_overrides_sal, estab_cap_overrides,
                cap_prensas_ent_1, cap_prensas_ent_2,
                cap_prensas_sal_1, cap_prensas_sal_2,
                cap_overrides_prensas_ent, cap_overrides_prensas_sal
            )
            st.session_state["df_sugerencias"] = df_sug

        with st.expander("🧩 Lotes que no encajan: sugerencias", expanded=not df_sug.empty):
            if df_sug.empty:
                st.success("Todos los lotes encajan con las restricciones actuales. 🎉")
            else:
                st.dataframe(df_sug, use_container_width=True, hide_index=True)
                sug_xlsx = generar_excel(df_sug, "sugerencias_lotes_no_encajan.xlsx")
                st.download_button(
                    "💾 Descargar sugerencias (Excel)",
                    data=sug_xlsx,
                    file_name="sugerencias_lotes_no_encajan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # -------------------------------
        # Descargar Excel
        # -------------------------------
        excel_bytes = generar_excel(df_editable, "planificacion_lotes.xlsx")
        st.download_button(
            label="💾 Descargar Excel con planificación",
            data=excel_bytes,
            file_name="planificacion_lotes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



