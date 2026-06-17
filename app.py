import io
import re
import unicodedata
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ============================================================
# APP: Validador de bases de Prima y Cesantías - Nómina JMC
# Enfoque correcto:
# - NO es proyección.
# - Genera bases prestacionales usando acumulados históricos.
# - Calcula el componente salarial mediante histórico de salarios.
# - Permite validar diferencias entre salario histórico y acumulados SAP.
# ============================================================

st.set_page_config(
    page_title="Bases prima y cesantías | Nómina JMC",
    page_icon="🦜",
    layout="wide",
)

# -----------------------------
# Parámetros base
# -----------------------------
SALARY_CONCEPTS_DEFAULT = {"Y010", "Y011", "Y020", "Y050", "Y051", "Y090"}

DEFAULT_BASE_PRESTACIONES = {
    # Salariales / básicos, normalmente se validan contra histórico de salarios
    "Y010", "Y011", "Y020", "Y050", "Y051", "Y090",
    # Auxilio transporte, si aplica en la parametrización de prestaciones
    "Y200",
    # Horas, recargos, compensatorios y variables habituales
    "Y220", "Y221", "Y300", "Y305", "Y310", "Y315", "Y350", "YM01",
    # Bonos salariales habituales del modelo financiero
    "Y506", "Y610", "Y617", "Y618",
}

EXCEL_EXTS = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}
CSV_EXTS = {".csv", ".txt"}
MAX_SAFE_DATE = date(2099, 12, 31)

SAP_CANDIDATES = [
    "sap", "nº pers", "n° pers", "no pers", "nro pers", "numero personal",
    "número personal", "pernr", "cod sap", "codigo sap", "código sap", "usuario",
    "employee id", "id empleado", "colaborador", "n pers", "n. pers", "personal",
]
CONCEPT_CANDIDATES = [
    "cc-nomina", "cc nomina", "cc-nómina", "cc nómina", "concepto", "codigo concepto",
    "código concepto", "cod concepto", "lgart", "clase de nomina", "clase de nómina",
    "cc nom", "cl.nomina", "cl nómina",
]
CONCEPT_TEXT_CANDIDATES = [
    "texto", "descripcion", "descripción", "desc concepto", "descripcion concepto",
    "descripción concepto", "texto concepto", "nombre concepto", "concepto texto",
]
VALUE_CANDIDATES = [
    "valor", "importe", "monto", "devengo", "valor pago", "valor pagado", "total",
    "amount", "valor concepto", "importe moneda", "valor ml", "valor acumulado",
    "pago", "pagado", "vlr", "valor nómina", "valor nomina",
]
QTY_CANDIDATES = ["cantidad", "cant", "dias", "días", "horas", "numero", "número"]
PERIOD_CANDIDATES = [
    "periodo", "período", "periodo nomina", "período nómina", "periodo para nomina",
    "periodo para nómina", "fecha de pago", "fecha pago", "mes", "mes pago", "for-period",
    "periodo pago", "período pago", "fecha contabilizacion", "fecha contabilización",
    "fecha", "pay period", "payroll period",
]

MD_NAME_CANDIDATES = ["nombre", "nombres", "nombre completo", "empleado", "trabajador"]
MD_DOC_CANDIDATES = ["cedula", "cédula", "numero id", "número id", "documento", "id", "identificacion", "identificación"]
MD_AREA_CANDIDATES = ["area de nomina", "área de nómina", "area nomina", "área nómina"]
MD_CECO_CANDIDATES = ["ce.coste", "ce coste", "ceco", "centro de coste", "centro de costo", "centro coste"]
MD_CARGO_CANDIDATES = ["cargo", "funcion", "función", "posicion", "posición"]
MD_SALARY_CANDIDATES = ["salario total", "salario", "sueldo", "sueldo basico", "sueldo básico"]
MD_INGRESO_CANDIDATES = ["fecha ingreso", "fecha de ingreso", "ingreso", "alta", "fecha alta", "fecha contratación"]
MD_RETIRO_CANDIDATES = ["fecha retiro", "fecha de retiro", "retiro", "baja", "fecha baja"]

SAL_FROM_CANDIDATES = [
    "desde", "fecha desde", "vigencia desde", "fecha inicio", "inicio", "begda",
    "fecha salario", "fecha cambio", "fecha modificacion", "fecha modificación", "fecha",
]
SAL_TO_CANDIDATES = [
    "hasta", "fecha hasta", "vigencia hasta", "fecha fin", "fin", "endda",
    "fecha final", "fecha termino", "fecha término",
]
SAL_PERIOD_CANDIDATES = PERIOD_CANDIDATES + ["mes salario", "periodo salario"]

# -----------------------------
# Utilidades generales
# -----------------------------

def normalize_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = text.replace("_", " ").replace("-", " ").replace(".", " ")
    return text.strip()


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() if str(c).strip() else f"Columna_{i+1}" for i, c in enumerate(out.columns)]
    out = out.dropna(how="all")
    out = out.dropna(axis=1, how="all")
    return out


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    norm_cols = {normalize_text(c): c for c in df.columns}
    norm_candidates = [normalize_text(c) for c in candidates]

    for cand in norm_candidates:
        if cand in norm_cols:
            return norm_cols[cand]

    for cand in norm_candidates:
        for norm_col, original in norm_cols.items():
            if cand and cand in norm_col:
                return original

    for cand in norm_candidates:
        for norm_col, original in norm_cols.items():
            if norm_col and norm_col in cand:
                return original
    return None


def get_extension(name: str) -> str:
    m = re.search(r"(\.[A-Za-z0-9]+)$", name or "")
    return m.group(1).lower() if m else ""


def excel_engine_for_name(name: str) -> Optional[str]:
    lower = name.lower()
    if lower.endswith(".xlsb"):
        return "pyxlsb"
    if lower.endswith(".ods"):
        return "odf"
    if lower.endswith(".xls"):
        return "xlrd"
    return "openpyxl"


def get_sheet_names(uploaded_file) -> List[str]:
    ext = get_extension(uploaded_file.name)
    if ext not in EXCEL_EXTS:
        return ["Archivo plano"]
    data = uploaded_file.getvalue()
    engine = excel_engine_for_name(uploaded_file.name)
    xls = pd.ExcelFile(io.BytesIO(data), engine=engine)
    return xls.sheet_names


def detect_header_row(raw_preview: pd.DataFrame) -> int:
    keywords = (
        SAP_CANDIDATES + CONCEPT_CANDIDATES + VALUE_CANDIDATES + PERIOD_CANDIDATES +
        MD_NAME_CANDIDATES + MD_CECO_CANDIDATES + SAL_FROM_CANDIDATES + SAL_TO_CANDIDATES
    )
    norm_keywords = [normalize_text(k) for k in keywords]
    best_row = 0
    best_score = -1
    for idx in range(min(len(raw_preview), 35)):
        values = [normalize_text(v) for v in raw_preview.iloc[idx].tolist()]
        score = 0
        for v in values:
            if not v:
                continue
            for k in norm_keywords:
                if k and (v == k or k in v or v in k):
                    score += 1
        non_empty = sum(1 for v in values if v)
        if non_empty >= 2 and score > best_score:
            best_score = score
            best_row = idx
    return best_row if best_score > 0 else 0


def read_uploaded_table(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    ext = get_extension(uploaded_file.name)
    data = uploaded_file.getvalue()

    if ext in CSV_EXTS:
        best = None
        best_cols = -1
        for enc in ["utf-8-sig", "latin1", "cp1252"]:
            for sep in [";", ",", "\t", "|"]:
                try:
                    df = pd.read_csv(io.BytesIO(data), sep=sep, encoding=enc, dtype=str, engine="python")
                    df = clean_columns(df)
                    if len(df.columns) > best_cols:
                        best = df
                        best_cols = len(df.columns)
                except Exception:
                    pass
        if best is None:
            raise ValueError("No fue posible leer el archivo plano. Revisa separador o codificación.")
        return best

    if ext in EXCEL_EXTS:
        engine = excel_engine_for_name(uploaded_file.name)
        sheet = sheet_name if sheet_name and sheet_name != "Archivo plano" else 0
        preview = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, nrows=35, engine=engine)
        header_row = detect_header_row(preview)
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=header_row, engine=engine, dtype=object)
        return clean_columns(df)

    raise ValueError(f"Extensión no soportada: {ext}. Usa xlsx, xlsm, xls, xlsb, ods, csv o txt.")


def normalize_sap(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    digits = re.sub(r"\D", "", text)
    if digits:
        return digits.lstrip("0") or "0"
    return text.strip()


def extract_concept(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).upper().strip().replace(" ", "").replace("-", "")
    m = re.search(r"\b(YM\d{2})\b", text)
    if m:
        return m.group(1)
    m = re.search(r"\b([YZ][A-Z0-9]{2,4}|\d{3,4})\b", text)
    if m:
        return m.group(1)
    return text


def parse_number(value) -> float:
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    s = s.replace("$", "").replace("COP", "").replace(" ", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if "-" in s:
        negative = True
        s = s.replace("-", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
    try:
        n = float(s)
        return -n if negative else n
    except Exception:
        return 0.0


def to_number_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_number).astype(float)


def is_last_day_of_month(d: date) -> bool:
    return d == month_last_day(d.year, d.month)


def parse_date_value(value) -> pd.Timestamp:
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, date):
        return pd.Timestamp(value)
    if isinstance(value, (int, float)) and 20000 <= float(value) <= 80000:
        try:
            return pd.to_datetime(value, unit="D", origin="1899-12-30")
        except Exception:
            pass
    s = str(value).strip()
    if not s:
        return pd.NaT
    if "9999" in s:
        return pd.Timestamp(MAX_SAFE_DATE)
    s = re.sub(r"\s+00:00:00$", "", s)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def parse_date_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_date_value)


def parse_period_value(value) -> pd.Timestamp:
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    if isinstance(value, datetime):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    if isinstance(value, date):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    if isinstance(value, (int, float)) and 20000 <= float(value) <= 80000:
        dt = parse_date_value(value)
        if pd.notna(dt):
            return pd.Timestamp(year=dt.year, month=dt.month, day=1)
    s = str(value).strip()
    if not s:
        return pd.NaT
    s2 = s.replace("_", "/").replace("-", "/").replace(".", "/")
    s2 = re.sub(r"\s+", "", s2)
    m = re.fullmatch(r"(\d{6})", s2)
    if m:
        val = m.group(1)
        y1, m1 = int(val[:4]), int(val[4:])
        if 1900 <= y1 <= 2100 and 1 <= m1 <= 12:
            return pd.Timestamp(year=y1, month=m1, day=1)
        m2, y2 = int(val[:2]), int(val[2:])
        if 1 <= m2 <= 12 and 1900 <= y2 <= 2100:
            return pd.Timestamp(year=y2, month=m2, day=1)
    m = re.fullmatch(r"(\d{1,4})/(\d{1,4})", s2)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if 1900 <= a <= 2100 and 1 <= b <= 12:
            return pd.Timestamp(year=a, month=b, day=1)
        if 1 <= a <= 12 and 1900 <= b <= 2100:
            return pd.Timestamp(year=b, month=a, day=1)
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return pd.Timestamp(year=dt.year, month=dt.month, day=1)
    return pd.NaT


def parse_period_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_period_value)


def month_first_day(year: int, month: int) -> date:
    return date(year, month, 1)


def month_last_day(year: int, month: int) -> date:
    if month == 12:
        return date(year, 12, 31)
    return date(year, month + 1, 1) - timedelta(days=1)


def add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    return date(y, m, 1)


def day_30_value(d: date) -> int:
    if is_last_day_of_month(d):
        return 30
    return min(d.day, 30)


def payroll_days_360_inclusive(start_dt: date, end_dt: date) -> int:
    """Días tipo nómina colombiana: cada mes pesa 30; último día de mes cuenta como día 30."""
    if pd.isna(start_dt) or pd.isna(end_dt) or end_dt < start_dt:
        return 0
    return (end_dt.year - start_dt.year) * 360 + (end_dt.month - start_dt.month) * 30 + (day_30_value(end_dt) - day_30_value(start_dt)) + 1


def actual_days_inclusive(start_dt: date, end_dt: date) -> int:
    if pd.isna(start_dt) or pd.isna(end_dt) or end_dt < start_dt:
        return 0
    return (end_dt - start_dt).days + 1


def count_days(start_dt: date, end_dt: date, day_mode: str) -> int:
    if day_mode == "Días calendario reales":
        return actual_days_inclusive(start_dt, end_dt)
    return payroll_days_360_inclusive(start_dt, end_dt)


def as_date(value) -> Optional[date]:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    dt = parse_date_value(value)
    return dt.date() if pd.notna(dt) else None


def effective_window(period_start: date, period_end: date, ingreso, retiro) -> Tuple[Optional[date], Optional[date]]:
    ini = period_start
    fin = period_end
    ing = as_date(ingreso)
    ret = as_date(retiro)
    if ing:
        ini = max(ini, ing)
    if ret and ret < MAX_SAFE_DATE:
        fin = min(fin, ret)
    if fin < ini:
        return None, None
    return ini, fin


def format_month(dt: pd.Timestamp) -> str:
    return dt.strftime("%Y-%m") if pd.notna(dt) else ""


def parse_bool(value) -> bool:
    if value is None or pd.isna(value):
        return False
    s = normalize_text(value)
    return s in {"si", "s", "x", "true", "verdadero", "1", "aplica", "yes", "y", "ok"}

# -----------------------------
# Estandarización de insumos
# -----------------------------

def standardize_accumulated(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    sap_col = find_column(df, SAP_CANDIDATES)
    concept_col = find_column(df, CONCEPT_CANDIDATES)
    text_col = find_column(df, CONCEPT_TEXT_CANDIDATES)
    value_col = find_column(df, VALUE_CANDIDATES)
    qty_col = find_column(df, QTY_CANDIDATES)
    period_col = find_column(df, PERIOD_CANDIDATES)

    missing = []
    if not sap_col:
        missing.append("SAP / Nº pers.")
    if not concept_col:
        missing.append("Concepto / CC-nómina")
    if not value_col:
        missing.append("Valor")
    if not period_col:
        missing.append("Periodo / Fecha de pago")
    if missing:
        raise ValueError("No encontré estas columnas en acumulados: " + ", ".join(missing))

    out = pd.DataFrame()
    out["SAP"] = df[sap_col].apply(normalize_sap)
    out["Concepto"] = df[concept_col].apply(extract_concept)
    out["Texto Concepto"] = df[text_col].astype(str).str.strip() if text_col else out["Concepto"]
    out["Valor"] = to_number_series(df[value_col])
    out["Cantidad"] = to_number_series(df[qty_col]) if qty_col else 0.0
    out["Periodo_Mes"] = parse_period_series(df[period_col])
    out["Periodo original"] = df[period_col]
    out = out[(out["SAP"] != "") & (out["Concepto"] != "")]
    out = out[pd.notna(out["Periodo_Mes"])]

    detected = {
        "SAP": sap_col,
        "Concepto": concept_col,
        "Texto Concepto": text_col or "No detectada",
        "Valor": value_col,
        "Cantidad": qty_col or "No detectada",
        "Periodo": period_col,
    }
    return out, detected


def standardize_md(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    sap_col = find_column(df, SAP_CANDIDATES)
    if not sap_col:
        raise ValueError("No encontré columna SAP / Nº pers. en el Master Data.")

    name_col = find_column(df, MD_NAME_CANDIDATES)
    doc_col = find_column(df, MD_DOC_CANDIDATES)
    area_col = find_column(df, MD_AREA_CANDIDATES)
    ceco_col = find_column(df, MD_CECO_CANDIDATES)
    cargo_col = find_column(df, MD_CARGO_CANDIDATES)
    salary_col = find_column(df, MD_SALARY_CANDIDATES)
    ingreso_col = find_column(df, MD_INGRESO_CANDIDATES)
    retiro_col = find_column(df, MD_RETIRO_CANDIDATES)

    out = pd.DataFrame()
    out["SAP"] = df[sap_col].apply(normalize_sap)
    out["Cédula"] = df[doc_col].astype(str).str.strip() if doc_col else ""
    out["Nombre"] = df[name_col].astype(str).str.strip() if name_col else ""
    out["Área de nómina"] = df[area_col].astype(str).str.strip() if area_col else ""
    out["CECO"] = df[ceco_col].astype(str).str.strip() if ceco_col else ""
    out["Cargo"] = df[cargo_col].astype(str).str.strip() if cargo_col else ""
    out["Salario actual MD"] = to_number_series(df[salary_col]) if salary_col else 0.0
    out["Fecha ingreso"] = parse_date_series(df[ingreso_col]) if ingreso_col else pd.NaT
    out["Fecha retiro"] = parse_date_series(df[retiro_col]) if retiro_col else pd.NaT
    out = out[out["SAP"] != ""].drop_duplicates(subset=["SAP"], keep="last")

    detected = {
        "SAP": sap_col,
        "Cédula": doc_col or "No detectada",
        "Nombre": name_col or "No detectada",
        "Área de nómina": area_col or "No detectada",
        "CECO": ceco_col or "No detectada",
        "Cargo": cargo_col or "No detectada",
        "Salario actual MD": salary_col or "No detectada",
        "Fecha ingreso": ingreso_col or "No detectada",
        "Fecha retiro": retiro_col or "No detectada",
    }
    return out, detected


def standardize_salary_history(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    sap_col = find_column(df, SAP_CANDIDATES)
    salary_col = find_column(df, MD_SALARY_CANDIDATES + ["salario mensual", "valor salario", "sueldo mensual"])
    from_col = find_column(df, SAL_FROM_CANDIDATES)
    to_col = find_column(df, SAL_TO_CANDIDATES)
    period_col = find_column(df, SAL_PERIOD_CANDIDATES)

    missing = []
    if not sap_col:
        missing.append("SAP / Nº pers.")
    if not salary_col:
        missing.append("Salario")
    if not from_col and not period_col:
        missing.append("Desde / Fecha de inicio / Periodo")
    if missing:
        raise ValueError("No encontré estas columnas en histórico de salarios: " + ", ".join(missing))

    out = pd.DataFrame()
    out["SAP"] = df[sap_col].apply(normalize_sap)
    out["Salario"] = to_number_series(df[salary_col])

    if from_col:
        out["Desde"] = parse_date_series(df[from_col])
    else:
        out["Desde"] = parse_period_series(df[period_col])

    if to_col:
        out["Hasta"] = parse_date_series(df[to_col])
    elif period_col and not from_col:
        per = parse_period_series(df[period_col])
        out["Hasta"] = per.apply(lambda x: pd.Timestamp(month_last_day(x.year, x.month)) if pd.notna(x) else pd.NaT)
    else:
        out["Hasta"] = pd.NaT

    out = out[(out["SAP"] != "") & (out["Salario"] > 0) & pd.notna(out["Desde"])]
    out = out.sort_values(["SAP", "Desde"]).copy()

    # Si no viene fecha hasta, se infiere con el siguiente cambio de salario del mismo SAP.
    out["Siguiente_Desde"] = out.groupby("SAP")["Desde"].shift(-1)
    inferred_until = out["Siguiente_Desde"] - pd.Timedelta(days=1)
    out["Hasta"] = out["Hasta"].where(pd.notna(out["Hasta"]), inferred_until)
    out["Hasta"] = out["Hasta"].where(pd.notna(out["Hasta"]), pd.Timestamp(MAX_SAFE_DATE))
    out.loc[out["Hasta"] < out["Desde"], "Hasta"] = out.loc[out["Hasta"] < out["Desde"], "Desde"]
    out = out.drop(columns=["Siguiente_Desde"])

    detected = {
        "SAP": sap_col,
        "Salario": salary_col,
        "Desde": from_col or "No detectada; se usó Periodo",
        "Hasta": to_col or "No detectada; se infirió con siguiente cambio o 31/12/2099",
        "Periodo": period_col or "No detectada",
    }
    return out, detected


def standardize_param(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    concept_col = find_column(df, CONCEPT_CANDIDATES + ["codigo", "código"])
    desc_col = find_column(df, CONCEPT_TEXT_CANDIDATES)
    prima_col = find_column(df, ["base prima", "prima", "aplica prima"])
    ces_col = find_column(df, ["base cesantias", "base cesantías", "cesantias", "cesantías", "aplica cesantias", "aplica cesantías"])
    tipo_col = find_column(df, ["tipo", "tipo componente", "componente", "clasificacion", "clasificación"])

    if not concept_col:
        raise ValueError("La parametrización debe tener una columna de Concepto / CC-nómina.")

    out = pd.DataFrame()
    out["Concepto"] = df[concept_col].apply(extract_concept)
    out["Descripción"] = df[desc_col].astype(str).str.strip() if desc_col else out["Concepto"]
    out["Base_Prima"] = df[prima_col].apply(parse_bool) if prima_col else True
    out["Base_Cesantias"] = df[ces_col].apply(parse_bool) if ces_col else True
    out["Tipo_Componente"] = df[tipo_col].astype(str).str.strip() if tipo_col else "Variable acumulado"
    out = out[out["Concepto"] != ""].drop_duplicates("Concepto", keep="last")

    detected = {
        "Concepto": concept_col,
        "Descripción": desc_col or "No detectada",
        "Base Prima": prima_col or "No detectada; se asumió Sí",
        "Base Cesantías": ces_col or "No detectada; se asumió Sí",
        "Tipo componente": tipo_col or "No detectada; se asumió Variable acumulado",
    }
    return out, detected

# -----------------------------
# Motor de cálculo
# -----------------------------

def build_population(md: Optional[pd.DataFrame], accum: pd.DataFrame, salary_hist: pd.DataFrame) -> pd.DataFrame:
    if md is not None and not md.empty:
        pop = md.copy()
    else:
        saps = sorted(set(accum["SAP"].dropna().astype(str)) | set(salary_hist["SAP"].dropna().astype(str)))
        pop = pd.DataFrame({"SAP": saps})
        for col in ["Cédula", "Nombre", "Área de nómina", "CECO", "Cargo"]:
            pop[col] = ""
        pop["Salario actual MD"] = 0.0
        pop["Fecha ingreso"] = pd.NaT
        pop["Fecha retiro"] = pd.NaT

    extra_saps = sorted((set(accum["SAP"].astype(str)) | set(salary_hist["SAP"].astype(str))) - set(pop["SAP"].astype(str)))
    if extra_saps:
        extra = pd.DataFrame({"SAP": extra_saps})
        for col in ["Cédula", "Nombre", "Área de nómina", "CECO", "Cargo"]:
            extra[col] = ""
        extra["Salario actual MD"] = 0.0
        extra["Fecha ingreso"] = pd.NaT
        extra["Fecha retiro"] = pd.NaT
        pop = pd.concat([pop, extra], ignore_index=True)
    return pop.drop_duplicates("SAP", keep="last")


def employee_divisor_days(population: pd.DataFrame, period_start: date, period_end: date, day_mode: str) -> pd.DataFrame:
    rows = []
    for _, r in population.iterrows():
        sap = r["SAP"]
        eff_start, eff_end = effective_window(period_start, period_end, r.get("Fecha ingreso", pd.NaT), r.get("Fecha retiro", pd.NaT))
        days = count_days(eff_start, eff_end, day_mode) if eff_start and eff_end else 0
        rows.append({
            "SAP": sap,
            "Inicio efectivo": eff_start,
            "Fin efectivo": eff_end,
            "Días divisor": days,
        })
    return pd.DataFrame(rows)


def salary_average_for_period(
    population: pd.DataFrame,
    salary_hist: pd.DataFrame,
    period_start: date,
    period_end: date,
    day_mode: str,
    label: str,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Calcula salario mensual promedio usando histórico salarial y genera detalle por segmento."""
    divisor = employee_divisor_days(population, period_start, period_end, day_mode)
    salary_by_sap = {sap: g.sort_values(["Desde", "Hasta"]).copy() for sap, g in salary_hist.groupby("SAP")}

    summary_rows = []
    segment_rows = []

    for _, pop in population.iterrows():
        sap = pop["SAP"]
        eff_start, eff_end = effective_window(period_start, period_end, pop.get("Fecha ingreso", pd.NaT), pop.get("Fecha retiro", pd.NaT))
        div_days = count_days(eff_start, eff_end, day_mode) if eff_start and eff_end else 0

        total_equiv = 0.0
        covered_days = 0
        missing_days = div_days
        min_salary = None
        max_salary = None
        changes = 0

        if not eff_start or not eff_end or div_days == 0:
            status = "Sin días en periodo según ingreso/retiro"
        else:
            recs = salary_by_sap.get(sap)
            if recs is None or recs.empty:
                status = "Sin histórico salarial"
            else:
                points = {eff_start, eff_end + timedelta(days=1)}
                usable_records = []
                for _, rec in recs.iterrows():
                    rec_start = as_date(rec["Desde"])
                    rec_end = as_date(rec["Hasta"]) or MAX_SAFE_DATE
                    if not rec_start:
                        continue
                    if rec_end < eff_start or rec_start > eff_end:
                        continue
                    usable_records.append(rec)
                    points.add(max(eff_start, rec_start))
                    if rec_end < eff_end:
                        points.add(rec_end + timedelta(days=1))
                points = sorted(points)

                missing_days = 0
                previous_salary = None
                for i in range(len(points) - 1):
                    seg_start = points[i]
                    seg_end = points[i + 1] - timedelta(days=1)
                    if seg_end < seg_start:
                        continue
                    days = count_days(seg_start, seg_end, day_mode)
                    if days <= 0:
                        continue
                    matching = []
                    for rec in usable_records:
                        rec_start = as_date(rec["Desde"])
                        rec_end = as_date(rec["Hasta"]) or MAX_SAFE_DATE
                        if rec_start and rec_start <= seg_start <= rec_end:
                            matching.append(rec)
                    if not matching:
                        missing_days += days
                        segment_rows.append({
                            "SAP": sap,
                            "Prestación": label,
                            "Segmento desde": seg_start,
                            "Segmento hasta": seg_end,
                            "Salario": 0.0,
                            "Días segmento": days,
                            "Valor salario equivalente": 0.0,
                            "Estado segmento": "Sin salario para este tramo",
                        })
                        continue

                    # Si hay traslape, se toma el registro con fecha Desde más reciente.
                    chosen = sorted(matching, key=lambda x: as_date(x["Desde"]) or date(1900, 1, 1))[-1]
                    sal = float(chosen["Salario"])
                    val = sal / 30.0 * days
                    total_equiv += val
                    covered_days += days
                    min_salary = sal if min_salary is None else min(min_salary, sal)
                    max_salary = sal if max_salary is None else max(max_salary, sal)
                    if previous_salary is not None and sal != previous_salary:
                        changes += 1
                    previous_salary = sal

                    segment_rows.append({
                        "SAP": sap,
                        "Prestación": label,
                        "Segmento desde": seg_start,
                        "Segmento hasta": seg_end,
                        "Salario": sal,
                        "Días segmento": days,
                        "Valor salario equivalente": val,
                        "Estado segmento": "OK",
                    })

                if missing_days > 0:
                    status = "Revisar: histórico salarial incompleto"
                else:
                    status = "OK"

        avg_salary = total_equiv / div_days * 30 if div_days else 0.0
        summary_rows.append({
            "SAP": sap,
            f"Días divisor {label}": div_days,
            f"Días con salario histórico {label}": covered_days,
            f"Días sin salario histórico {label}": missing_days,
            f"Salario histórico acumulado equivalente {label}": total_equiv,
            f"Salario histórico promedio {label}": avg_salary,
            f"Salario mínimo histórico {label}": min_salary or 0.0,
            f"Salario máximo histórico {label}": max_salary or 0.0,
            f"Cambios salario detectados {label}": changes,
            f"Estado salario histórico {label}": status,
        })

    return pd.DataFrame(summary_rows), pd.DataFrame(segment_rows)


def accum_base_for_period(
    population: pd.DataFrame,
    accum: pd.DataFrame,
    concepts: List[str],
    period_start: date,
    period_end: date,
    day_mode: str,
    label: str,
    output_prefix: str,
) -> pd.DataFrame:
    concepts_set = set(concepts)
    start_month = pd.Timestamp(period_start.year, period_start.month, 1)
    end_month = pd.Timestamp(period_end.year, period_end.month, 1)
    divisor = employee_divisor_days(population, period_start, period_end, day_mode)

    used = accum[
        (accum["Periodo_Mes"] >= start_month)
        & (accum["Periodo_Mes"] <= end_month)
        & (accum["Concepto"].isin(concepts_set))
    ].copy()

    if used.empty:
        grouped = pd.DataFrame(columns=["SAP", f"Valor acumulado {output_prefix} {label}", f"Registros acumulados {output_prefix} {label}", f"Meses con acumulado {output_prefix} {label}"])
    else:
        grouped = used.groupby("SAP", as_index=False).agg(
            **{
                f"Valor acumulado {output_prefix} {label}": ("Valor", "sum"),
                f"Registros acumulados {output_prefix} {label}": ("Valor", "size"),
                f"Meses con acumulado {output_prefix} {label}": ("Periodo_Mes", lambda s: s.dt.strftime("%Y-%m").nunique()),
            }
        )

    out = divisor.merge(grouped, on="SAP", how="left")
    val_col = f"Valor acumulado {output_prefix} {label}"
    reg_col = f"Registros acumulados {output_prefix} {label}"
    mes_col = f"Meses con acumulado {output_prefix} {label}"
    base_col = f"Base promedio {output_prefix} {label}"
    out[val_col] = out[val_col].fillna(0.0)
    out[reg_col] = out[reg_col].fillna(0).astype(int)
    out[mes_col] = out[mes_col].fillna(0).astype(int)
    div = out["Días divisor"].replace(0, pd.NA)
    out[base_col] = (out[val_col] / div * 30).fillna(0.0)
    return out[["SAP", val_col, reg_col, mes_col, base_col]]


def compose_base_sheet(
    population: pd.DataFrame,
    salary_summary: pd.DataFrame,
    accum_var: pd.DataFrame,
    accum_sal: pd.DataFrame,
    accum_total: pd.DataFrame,
    period_start: date,
    period_end: date,
    label: str,
    tolerance: float,
) -> pd.DataFrame:
    base = population.copy()
    base = base.merge(salary_summary, on="SAP", how="left")
    base = base.merge(accum_var, on="SAP", how="left")
    base = base.merge(accum_sal, on="SAP", how="left")
    base = base.merge(accum_total, on="SAP", how="left")

    for c in base.columns:
        if any(x in c for x in ["Valor acumulado", "Registros acumulados", "Meses con acumulado", "Base promedio", "Salario histórico", "Días con", "Días sin", "Días divisor", "Cambios salario", "Salario mínimo", "Salario máximo"]):
            if c.startswith("Estado"):
                continue
            base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0.0)

    base[f"Periodo inicial {label}"] = period_start.strftime("%d/%m/%Y")
    base[f"Periodo final {label}"] = period_end.strftime("%d/%m/%Y")

    sal_prom = f"Salario histórico promedio {label}"
    var_base = f"Base promedio variables acumuladas {label}"
    sal_acc_base = f"Base promedio conceptos salariales acumulados {label}"
    total_acc_base = f"Base promedio total acumulados {label}"

    for col in [sal_prom, var_base, sal_acc_base, total_acc_base]:
        if col not in base.columns:
            base[col] = 0.0

    base[f"Base final {label} histórico + acumulados"] = base[sal_prom] + base[var_base]
    base[f"Validación base acumulada SAP {label}"] = base[total_acc_base]
    base[f"Diferencia final vs acumulado SAP {label}"] = base[f"Base final {label} histórico + acumulados"] - base[f"Validación base acumulada SAP {label}"]
    base[f"Diferencia salario histórico vs salario acumulado {label}"] = base[sal_prom] - base[sal_acc_base]

    estado_sal = f"Estado salario histórico {label}"
    base[f"Estado {label}"] = "OK"
    if estado_sal in base.columns:
        base.loc[base[estado_sal].astype(str).str.contains("Sin histórico", case=False, na=False), f"Estado {label}"] = "Revisar: sin histórico salarial"
        base.loc[base[estado_sal].astype(str).str.contains("incompleto", case=False, na=False), f"Estado {label}"] = "Revisar: histórico salarial incompleto"
    div_col = f"Días divisor {label}"
    if div_col in base.columns:
        base.loc[base[div_col] == 0, f"Estado {label}"] = "Sin días en periodo"
    diff_col = f"Diferencia salario histórico vs salario acumulado {label}"
    if diff_col in base.columns and sal_acc_base in base.columns:
        mask_diff = (base[sal_acc_base].abs() > 0) & (base[diff_col].abs() > tolerance)
        base.loc[mask_diff & (base[f"Estado {label}"] == "OK"), f"Estado {label}"] = "Revisar: diferencia salario histórico vs acumulado"

    ordered = [
        "SAP", "Cédula", "Nombre", "Área de nómina", "CECO", "Cargo", "Salario actual MD",
        "Fecha ingreso", "Fecha retiro", f"Periodo inicial {label}", f"Periodo final {label}",
        f"Días divisor {label}", f"Días con salario histórico {label}", f"Días sin salario histórico {label}",
        f"Salario histórico acumulado equivalente {label}", f"Salario histórico promedio {label}",
        f"Salario mínimo histórico {label}", f"Salario máximo histórico {label}", f"Cambios salario detectados {label}",
        f"Valor acumulado variables acumuladas {label}", f"Base promedio variables acumuladas {label}",
        f"Valor acumulado conceptos salariales acumulados {label}", f"Base promedio conceptos salariales acumulados {label}",
        f"Valor acumulado total acumulados {label}", f"Base promedio total acumulados {label}",
        f"Base final {label} histórico + acumulados", f"Validación base acumulada SAP {label}",
        f"Diferencia final vs acumulado SAP {label}", f"Diferencia salario histórico vs salario acumulado {label}",
        f"Estado salario histórico {label}", f"Estado {label}",
    ]
    existing = [c for c in ordered if c in base.columns]
    rest = [c for c in base.columns if c not in existing]
    return base[existing + rest]


def build_concepts_table(accum: pd.DataFrame, prima_concepts: List[str], ces_concepts: List[str], salary_concepts: List[str]) -> pd.DataFrame:
    concepts = accum.groupby(["Concepto", "Texto Concepto"], as_index=False).agg(
        Valor_Total_Archivo=("Valor", "sum"),
        Registros=("Valor", "size"),
        Primer_Periodo=("Periodo_Mes", "min"),
        Ultimo_Periodo=("Periodo_Mes", "max"),
    ).sort_values(["Concepto", "Texto Concepto"])
    concepts["Primer_Periodo"] = concepts["Primer_Periodo"].dt.strftime("%Y-%m")
    concepts["Ultimo_Periodo"] = concepts["Ultimo_Periodo"].dt.strftime("%Y-%m")
    concepts["Usado_Base_Prima"] = concepts["Concepto"].isin(prima_concepts)
    concepts["Usado_Base_Cesantias"] = concepts["Concepto"].isin(ces_concepts)
    concepts["Concepto_Salarial_Validado_por_Historico"] = concepts["Concepto"].isin(salary_concepts)
    concepts["Tratamiento"] = concepts["Concepto_Salarial_Validado_por_Historico"].map(lambda x: "Salario histórico" if x else "Acumulado variable")
    return concepts


def build_detail_used(
    accum: pd.DataFrame,
    prima_concepts: List[str],
    ces_concepts: List[str],
    salary_concepts: List[str],
    prima_start: date,
    prima_end: date,
    ces_start: date,
    ces_end: date,
) -> pd.DataFrame:
    detail = accum.copy()
    detail["Periodo"] = detail["Periodo_Mes"].dt.strftime("%Y-%m")
    detail["Concepto salarial histórico"] = detail["Concepto"].isin(salary_concepts)
    detail["En periodo prima"] = (detail["Periodo_Mes"] >= pd.Timestamp(prima_start.year, prima_start.month, 1)) & (detail["Periodo_Mes"] <= pd.Timestamp(prima_end.year, prima_end.month, 1))
    detail["En periodo cesantías"] = (detail["Periodo_Mes"] >= pd.Timestamp(ces_start.year, ces_start.month, 1)) & (detail["Periodo_Mes"] <= pd.Timestamp(ces_end.year, ces_end.month, 1))
    detail["Usado prima total"] = detail["En periodo prima"] & detail["Concepto"].isin(prima_concepts)
    detail["Usado prima variable"] = detail["Usado prima total"] & (~detail["Concepto salarial histórico"])
    detail["Usado prima salarial validación"] = detail["Usado prima total"] & detail["Concepto salarial histórico"]
    detail["Usado cesantías total"] = detail["En periodo cesantías"] & detail["Concepto"].isin(ces_concepts)
    detail["Usado cesantías variable"] = detail["Usado cesantías total"] & (~detail["Concepto salarial histórico"])
    detail["Usado cesantías salarial validación"] = detail["Usado cesantías total"] & detail["Concepto salarial histórico"]
    cols = [
        "SAP", "Concepto", "Texto Concepto", "Valor", "Cantidad", "Periodo", "Periodo original",
        "Concepto salarial histórico", "En periodo prima", "Usado prima total", "Usado prima variable", "Usado prima salarial validación",
        "En periodo cesantías", "Usado cesantías total", "Usado cesantías variable", "Usado cesantías salarial validación",
    ]
    return detail[cols]


def build_alerts(base_prima: pd.DataFrame, base_ces: pd.DataFrame, tolerance: float) -> pd.DataFrame:
    rows = []
    for label, df in [("Prima", base_prima), ("Cesantias", base_ces)]:
        estado_col = f"Estado {label}"
        diff_col = f"Diferencia salario histórico vs salario acumulado {label}"
        if estado_col not in df.columns:
            continue
        subset = df[df[estado_col] != "OK"].copy()
        for _, r in subset.iterrows():
            rows.append({
                "SAP": r.get("SAP", ""),
                "Nombre": r.get("Nombre", ""),
                "Prestación": label,
                "Alerta": r.get(estado_col, ""),
                "Diferencia salario histórico vs acumulado": r.get(diff_col, 0),
                "Tolerancia usada": tolerance,
            })
    return pd.DataFrame(rows)


def make_excel_report(
    base_prima: pd.DataFrame,
    base_ces: pd.DataFrame,
    llevar_modelo: pd.DataFrame,
    detail: pd.DataFrame,
    salary_segments: pd.DataFrame,
    concepts: pd.DataFrame,
    alerts: pd.DataFrame,
    log_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        sheets = {
            "Llevar_al_Modelo": llevar_modelo,
            "Base_Prima": base_prima,
            "Base_Cesantias": base_ces,
            "Detalle_Acumulados": detail,
            "Historico_Salarios_Calculo": salary_segments,
            "Conceptos_Usados": concepts,
            "Alertas": alerts,
            "Log": log_df,
        }
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        int_fmt = workbook.add_format({"num_format": "#,##0"})
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})
        pct_fmt = workbook.add_format({"num_format": "0.00%"})

        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df_to_write = df.copy()
            for col in df_to_write.columns:
                if pd.api.types.is_datetime64_any_dtype(df_to_write[col]):
                    df_to_write[col] = df_to_write[col].dt.date
            df_to_write.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            ws.freeze_panes(1, 0)
            if len(df_to_write.columns) > 0:
                ws.autofilter(0, 0, max(len(df_to_write), 1), len(df_to_write.columns) - 1)
            for col_idx, col_name in enumerate(df_to_write.columns):
                ws.write(0, col_idx, col_name, header_fmt)
                width = min(max(len(str(col_name)) + 2, 12), 48)
                if not df_to_write.empty:
                    try:
                        sample = df_to_write[col_name].astype(str).head(300).map(len).max()
                        width = min(max(width, int(sample) + 2), 48)
                    except Exception:
                        pass
                lower = str(col_name).lower()
                fmt = None
                if any(k in lower for k in ["valor", "base", "salario", "promedio", "acumulado", "diferencia", "tolerancia"]):
                    fmt = money_fmt
                elif any(k in lower for k in ["dias", "días", "meses", "registros", "cambios"]):
                    fmt = int_fmt
                elif "fecha" in lower or "desde" in lower or "hasta" in lower or "inicio" in lower or "fin" in lower:
                    fmt = date_fmt
                elif "%" in lower or "porcentaje" in lower:
                    fmt = pct_fmt
                ws.set_column(col_idx, col_idx, width, fmt)
    return output.getvalue()


def make_param_template() -> bytes:
    df = pd.DataFrame({
        "Concepto": sorted(DEFAULT_BASE_PRESTACIONES),
        "Descripción": "",
        "Base_Prima": "Sí",
        "Base_Cesantías": "Sí",
        "Tipo_Componente": ["Salario histórico" if c in SALARY_CONCEPTS_DEFAULT else "Variable acumulado" for c in sorted(DEFAULT_BASE_PRESTACIONES)],
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Parametrizacion")
    return output.getvalue()


def default_semester_dates(today: date) -> Tuple[date, date]:
    if today.month <= 6:
        return date(today.year, 1, 1), date(today.year, 6, 30)
    return date(today.year, 7, 1), date(today.year, 12, 31)

# -----------------------------
# Interfaz
# -----------------------------

st.title("🦜 Validador de bases de prima y cesantías")
st.caption("Solo acumulados históricos + ejercicio de histórico salarial. No calcula proyección.")

with st.expander("📌 ¿Qué hace esta versión?", expanded=True):
    st.markdown(
        """
        Esta herramienta está pensada para **validar la base prestacional**, no para proyectar nómina.

        **Lógica aplicada:**
        1. Lee los **acumulados históricos** por empleado, concepto y periodo.
        2. Lee el **histórico de salarios** y arma el salario promedio del periodo según vigencias.
        3. Separa los conceptos salariales, como Y010/Y011/Y020/Y050/Y051/Y090, para no duplicarlos si ya se calculan por histórico.
        4. Calcula variables acumuladas promedio con la fórmula: `valor acumulado / días divisor × 30`.
        5. Genera la base final: **salario histórico promedio + variables acumuladas promedio**.
        6. Deja una validación contra lo acumulado en SAP para identificar diferencias.
        """
    )

with st.sidebar:
    st.header("⚙️ Parámetros")
    today = date.today()
    def_prima_ini, def_prima_fin = default_semester_dates(today)
    def_ces_ini = date(today.year, 1, 1)
    def_ces_fin = today

    prima_start = st.date_input("Inicio periodo prima", value=def_prima_ini, format="DD/MM/YYYY")
    prima_end = st.date_input("Fin periodo prima", value=def_prima_fin, format="DD/MM/YYYY")
    ces_start = st.date_input("Inicio periodo cesantías", value=def_ces_ini, format="DD/MM/YYYY")
    ces_end = st.date_input("Fin periodo cesantías", value=def_ces_fin, format="DD/MM/YYYY")

    day_mode = st.selectbox(
        "Días divisor",
        ["Días 360 nómina", "Días calendario reales"],
        index=0,
        help="Para nómina Colombia normalmente se usa base 360. El salario histórico se divide en salario/30 por días del tramo.",
    )
    tolerance = st.number_input("Tolerancia para alerta de diferencia", min_value=0.0, value=1000.0, step=1000.0)

    st.download_button(
        "⬇️ Descargar plantilla parametrización",
        data=make_param_template(),
        file_name="plantilla_parametrizacion_prima_cesantias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.subheader("1) Carga de archivos")
c1, c2, c3, c4 = st.columns(4)
with c1:
    accum_file = st.file_uploader("Acumulados de nómina", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"], key="accum")
with c2:
    salary_file = st.file_uploader("Histórico de salarios", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"], key="salary")
with c3:
    md_file = st.file_uploader("Master Data / empleados (opcional)", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"], key="md")
with c4:
    param_file = st.file_uploader("Parametrización conceptos (opcional)", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"], key="param")

# Selección de hojas
accum_sheet = salary_sheet = md_sheet = param_sheet = None
if accum_file:
    try:
        sheets = get_sheet_names(accum_file)
        accum_sheet = st.selectbox("Hoja acumulados", sheets, key="accum_sheet") if len(sheets) > 1 else sheets[0]
    except Exception as exc:
        st.error(f"No pude leer las hojas de acumulados: {exc}")
if salary_file:
    try:
        sheets = get_sheet_names(salary_file)
        salary_sheet = st.selectbox("Hoja histórico salarios", sheets, key="salary_sheet") if len(sheets) > 1 else sheets[0]
    except Exception as exc:
        st.error(f"No pude leer las hojas del histórico de salarios: {exc}")
if md_file:
    try:
        sheets = get_sheet_names(md_file)
        md_sheet = st.selectbox("Hoja Master Data", sheets, key="md_sheet") if len(sheets) > 1 else sheets[0]
    except Exception as exc:
        st.warning(f"No pude leer las hojas de Master Data: {exc}")
if param_file:
    try:
        sheets = get_sheet_names(param_file)
        param_sheet = st.selectbox("Hoja parametrización", sheets, key="param_sheet") if len(sheets) > 1 else sheets[0]
    except Exception as exc:
        st.warning(f"No pude leer las hojas de parametrización: {exc}")

# Lectura de insumos
accum_std = salary_std = md_std = param_std = None
read_log = []

if accum_file and accum_sheet:
    try:
        raw = read_uploaded_table(accum_file, accum_sheet)
        accum_std, detected = standardize_accumulated(raw)
        read_log.append({"Paso": "Acumulados", "Resultado": f"OK - {len(accum_std):,} registros", "Detalle": str(detected)})
        st.success(f"Acumulados leídos: {len(accum_std):,} registros")
        with st.expander("Columnas detectadas en acumulados"):
            st.json(detected)
    except Exception as exc:
        st.error(f"Error leyendo acumulados: {exc}")

if salary_file and salary_sheet:
    try:
        raw = read_uploaded_table(salary_file, salary_sheet)
        salary_std, detected = standardize_salary_history(raw)
        read_log.append({"Paso": "Histórico salarios", "Resultado": f"OK - {len(salary_std):,} registros", "Detalle": str(detected)})
        st.success(f"Histórico de salarios leído: {len(salary_std):,} registros")
        with st.expander("Columnas detectadas en histórico de salarios"):
            st.json(detected)
    except Exception as exc:
        st.error(f"Error leyendo histórico de salarios: {exc}")

if md_file and md_sheet:
    try:
        raw = read_uploaded_table(md_file, md_sheet)
        md_std, detected = standardize_md(raw)
        read_log.append({"Paso": "Master Data", "Resultado": f"OK - {len(md_std):,} empleados", "Detalle": str(detected)})
        st.success(f"Master Data leído: {len(md_std):,} empleados")
        with st.expander("Columnas detectadas en Master Data"):
            st.json(detected)
    except Exception as exc:
        st.warning(f"No se usará Master Data porque ocurrió un error: {exc}")

if param_file and param_sheet:
    try:
        raw = read_uploaded_table(param_file, param_sheet)
        param_std, detected = standardize_param(raw)
        read_log.append({"Paso": "Parametrización", "Resultado": f"OK - {len(param_std):,} conceptos", "Detalle": str(detected)})
        st.success(f"Parametrización leída: {len(param_std):,} conceptos")
        with st.expander("Columnas detectadas en parametrización"):
            st.json(detected)
    except Exception as exc:
        st.warning(f"No se usará parametrización porque ocurrió un error: {exc}")

if accum_std is not None and salary_std is not None:
    st.subheader("2) Conceptos y tratamiento")

    concept_catalog = accum_std.groupby(["Concepto", "Texto Concepto"], as_index=False).agg(
        Valor_Total=("Valor", "sum"),
        Registros=("Valor", "size"),
        Primer_Periodo=("Periodo_Mes", "min"),
        Ultimo_Periodo=("Periodo_Mes", "max"),
    ).sort_values("Concepto")
    concept_catalog["Primer_Periodo"] = concept_catalog["Primer_Periodo"].dt.strftime("%Y-%m")
    concept_catalog["Ultimo_Periodo"] = concept_catalog["Ultimo_Periodo"].dt.strftime("%Y-%m")

    all_concepts = sorted(concept_catalog["Concepto"].dropna().unique().tolist())

    if param_std is not None and not param_std.empty:
        prima_default = sorted(param_std.loc[param_std["Base_Prima"], "Concepto"].unique().tolist())
        ces_default = sorted(param_std.loc[param_std["Base_Cesantias"], "Concepto"].unique().tolist())
        salary_default = sorted(
            set(param_std.loc[param_std["Tipo_Componente"].apply(lambda x: "salario" in normalize_text(x)), "Concepto"].unique().tolist())
            | (set(all_concepts) & SALARY_CONCEPTS_DEFAULT)
        )
        st.info("Se tomó la parametrización cargada. Puedes ajustar los conceptos antes de generar.")
    else:
        prima_default = [c for c in all_concepts if c in DEFAULT_BASE_PRESTACIONES or c.startswith("Y")]
        ces_default = [c for c in all_concepts if c in DEFAULT_BASE_PRESTACIONES or c.startswith("Y")]
        salary_default = [c for c in all_concepts if c in SALARY_CONCEPTS_DEFAULT]
        st.warning("No cargaste parametrización. Se preseleccionaron conceptos Y y los salariales conocidos para validar contra histórico.")

    a, b = st.columns(2)
    with a:
        prima_concepts = st.multiselect("Conceptos que aplican para base de prima", all_concepts, default=[c for c in prima_default if c in all_concepts])
    with b:
        ces_concepts = st.multiselect("Conceptos que aplican para base de cesantías", all_concepts, default=[c for c in ces_default if c in all_concepts])

    salary_concepts = st.multiselect(
        "Conceptos salariales que se validan con histórico de salarios",
        all_concepts,
        default=[c for c in salary_default if c in all_concepts],
        help="Estos conceptos no se suman como variable si estás usando el histórico salarial. Se usan para comparar contra SAP.",
    )

    with st.expander("Ver catálogo de conceptos detectados"):
        st.dataframe(concept_catalog, use_container_width=True)

    st.subheader("3) Validación de periodos")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Periodo prima", f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}")
    with k2:
        st.metric("Periodo cesantías", f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}")
    with k3:
        st.metric("Conceptos prima", f"{len(prima_concepts):,}")
    with k4:
        st.metric("Conceptos cesantías", f"{len(ces_concepts):,}")

    periods_summary = accum_std.groupby(accum_std["Periodo_Mes"].dt.strftime("%Y-%m"), as_index=True).agg(
        Registros=("Valor", "size"),
        Valor=("Valor", "sum"),
    ).reset_index().rename(columns={"Periodo_Mes": "Periodo"})
    with st.expander("Periodos encontrados en acumulados"):
        st.dataframe(periods_summary, use_container_width=True)

    salary_periods = salary_std.copy()
    salary_periods["Desde"] = salary_periods["Desde"].dt.strftime("%d/%m/%Y")
    salary_periods["Hasta"] = salary_periods["Hasta"].dt.strftime("%d/%m/%Y")
    with st.expander("Vista rápida del histórico de salarios"):
        st.dataframe(salary_periods.head(300), use_container_width=True)

    generate = st.button("🚀 Generar bases de prima y cesantías", type="primary")

    if generate:
        if prima_end < prima_start:
            st.error("Revisa el periodo de prima: la fecha final no puede ser menor que la inicial.")
            st.stop()
        if ces_end < ces_start:
            st.error("Revisa el periodo de cesantías: la fecha final no puede ser menor que la inicial.")
            st.stop()
        if not prima_concepts:
            st.error("Selecciona al menos un concepto para prima.")
            st.stop()
        if not ces_concepts:
            st.error("Selecciona al menos un concepto para cesantías.")
            st.stop()

        progress = st.progress(0)
        status = st.empty()

        status.write("Preparando población...")
        population = build_population(md_std, accum_std, salary_std)
        progress.progress(10)

        prima_salary_concepts = sorted(set(prima_concepts) & set(salary_concepts))
        prima_variable_concepts = sorted(set(prima_concepts) - set(salary_concepts))
        ces_salary_concepts = sorted(set(ces_concepts) & set(salary_concepts))
        ces_variable_concepts = sorted(set(ces_concepts) - set(salary_concepts))

        status.write("Calculando salario histórico para prima...")
        sal_prima, seg_prima = salary_average_for_period(population, salary_std, prima_start, prima_end, day_mode, "Prima")
        progress.progress(25)

        status.write("Calculando salario histórico para cesantías...")
        sal_ces, seg_ces = salary_average_for_period(population, salary_std, ces_start, ces_end, day_mode, "Cesantias")
        salary_segments = pd.concat([seg_prima, seg_ces], ignore_index=True)
        progress.progress(40)

        status.write("Calculando acumulados de prima...")
        acc_prima_var = accum_base_for_period(population, accum_std, prima_variable_concepts, prima_start, prima_end, day_mode, "Prima", "variables acumuladas")
        acc_prima_sal = accum_base_for_period(population, accum_std, prima_salary_concepts, prima_start, prima_end, day_mode, "Prima", "conceptos salariales acumulados")
        acc_prima_total = accum_base_for_period(population, accum_std, prima_concepts, prima_start, prima_end, day_mode, "Prima", "total acumulados")
        progress.progress(55)

        status.write("Calculando acumulados de cesantías...")
        acc_ces_var = accum_base_for_period(population, accum_std, ces_variable_concepts, ces_start, ces_end, day_mode, "Cesantias", "variables acumuladas")
        acc_ces_sal = accum_base_for_period(population, accum_std, ces_salary_concepts, ces_start, ces_end, day_mode, "Cesantias", "conceptos salariales acumulados")
        acc_ces_total = accum_base_for_period(population, accum_std, ces_concepts, ces_start, ces_end, day_mode, "Cesantias", "total acumulados")
        progress.progress(70)

        status.write("Armando bases finales...")
        base_prima = compose_base_sheet(population, sal_prima, acc_prima_var, acc_prima_sal, acc_prima_total, prima_start, prima_end, "Prima", tolerance)
        base_ces = compose_base_sheet(population, sal_ces, acc_ces_var, acc_ces_sal, acc_ces_total, ces_start, ces_end, "Cesantias", tolerance)

        llevar_modelo = base_prima[[
            "SAP", "Cédula", "Nombre", "Área de nómina", "CECO", "Cargo", "Salario actual MD",
            "Salario histórico promedio Prima", "Base promedio variables acumuladas Prima",
            "Base final Prima histórico + acumulados", "Validación base acumulada SAP Prima",
            "Diferencia final vs acumulado SAP Prima", "Estado Prima",
        ]].merge(
            base_ces[[
                "SAP", "Salario histórico promedio Cesantias", "Base promedio variables acumuladas Cesantias",
                "Base final Cesantias histórico + acumulados", "Validación base acumulada SAP Cesantias",
                "Diferencia final vs acumulado SAP Cesantias", "Estado Cesantias",
            ]],
            on="SAP",
            how="outer",
        )
        llevar_modelo["Periodo prima usado"] = f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}"
        llevar_modelo["Periodo cesantías usado"] = f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}"
        llevar_modelo["Método"] = "Base final = salario histórico promedio + acumulados variables promedio"
        llevar_modelo["Observación"] = "Herramienta de validación de base prestacional; no usa proyección."
        progress.progress(82)

        status.write("Construyendo detalle, alertas y log...")
        detail = build_detail_used(accum_std, prima_concepts, ces_concepts, salary_concepts, prima_start, prima_end, ces_start, ces_end)
        concepts_table = build_concepts_table(accum_std, prima_concepts, ces_concepts, salary_concepts)
        alerts = build_alerts(base_prima, base_ces, tolerance)

        log_rows = read_log + [
            {"Paso": "Periodo prima", "Resultado": f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}", "Detalle": f"Conceptos prima: {', '.join(prima_concepts)}"},
            {"Paso": "Periodo cesantías", "Resultado": f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}", "Detalle": f"Conceptos cesantías: {', '.join(ces_concepts)}"},
            {"Paso": "Conceptos salariales", "Resultado": ", ".join(salary_concepts), "Detalle": "Se validan con histórico de salarios y no se duplican como variable."},
            {"Paso": "Días divisor", "Resultado": day_mode, "Detalle": "Se usa para salario histórico y acumulados."},
            {"Paso": "Fórmula base", "Resultado": "Salario histórico promedio + variables acumuladas promedio", "Detalle": "Acumulados promedio = valor acumulado / días divisor * 30."},
            {"Paso": "Tolerancia alertas", "Resultado": f"{tolerance:,.2f}", "Detalle": "Diferencias superiores generan alerta."},
        ]
        log_df = pd.DataFrame(log_rows)
        progress.progress(92)

        status.write("Generando Excel...")
        excel_bytes = make_excel_report(base_prima, base_ces, llevar_modelo, detail, salary_segments, concepts_table, alerts, log_df)
        progress.progress(100)
        status.success("Bases generadas correctamente.")

        st.subheader("✅ Resultado")
        r1, r2, r3, r4 = st.columns(4)
        with r1:
            st.metric("Empleados en salida", f"{len(llevar_modelo):,}")
        with r2:
            st.metric("Alertas", f"{len(alerts):,}")
        with r3:
            st.metric("Base prima total", f"{llevar_modelo['Base final Prima histórico + acumulados'].sum():,.0f}")
        with r4:
            st.metric("Base cesantías total", f"{llevar_modelo['Base final Cesantias histórico + acumulados'].sum():,.0f}")

        st.dataframe(llevar_modelo.head(300), use_container_width=True)
        if not alerts.empty:
            with st.expander("⚠️ Alertas generadas", expanded=True):
                st.dataframe(alerts, use_container_width=True)

        file_name = f"bases_prima_cesantias_validacion_historico_salarios_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.download_button(
            "⬇️ Descargar Excel generado",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

else:
    st.info("Carga como mínimo **Acumulados de nómina** e **Histórico de salarios** para continuar.")

st.divider()
st.caption("🦜 Creado por Andrés Huérfano Dávila - Nómina JMC")
