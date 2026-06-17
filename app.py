import io
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ============================================================
# APP: Validador bases Prima y Cesantías - Nómina JMC
# Enfoque corregido:
# - NO es una herramienta de proyección.
# - La base se valida con ACUMULADOS históricos + HISTÓRICO DE SALARIOS.
# - El salario fijo se calcula desde histórico salarial por vigencias.
# - Los acumulados aportan variables / auxilios / bonos según parametrización.
# - La parametrización por área de nómina define cómo se cuentan los días.
# ============================================================

st.set_page_config(
    page_title="Validador bases prima y cesantías | Nómina JMC",
    page_icon="🦜",
    layout="wide",
)

# -----------------------------
# Parametrizaciones base
# -----------------------------
FIXED_SALARY_CONCEPTS = {"Y010", "Y011", "Y020", "Y050", "Y051", "Y090"}

DEFAULT_CONCEPTS = pd.DataFrame(
    [
        # Concepto, descripción, prima, cesantías, tipo
        ("Y010", "Sueldo básico", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y011", "Part time días", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y020", "Salario integral", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y050", "Apoyo sostenimiento", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y051", "Apoyo practicante", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y090", "Part time horas", True, True, "SALARIO_FIJO_HISTORICO"),
        ("Y200", "Auxilio de transporte", True, True, "VARIABLE_ACUMULADO"),
        ("Y220", "Recargo nocturno", True, True, "VARIABLE_ACUMULADO"),
        ("Y221", "Recargo nocturno festivo", True, True, "VARIABLE_ACUMULADO"),
        ("Y300", "Hora extra diurna", True, True, "VARIABLE_ACUMULADO"),
        ("Y305", "Hora extra nocturna", True, True, "VARIABLE_ACUMULADO"),
        ("Y310", "Hora extra dominical diurna", True, True, "VARIABLE_ACUMULADO"),
        ("Y315", "Hora extra dominical nocturna", True, True, "VARIABLE_ACUMULADO"),
        ("Y350", "Compensatorio", True, True, "VARIABLE_ACUMULADO"),
        ("YM01", "Tiempo suplementario YM01", True, True, "VARIABLE_ACUMULADO"),
        ("Y506", "Bono salarial", True, True, "VARIABLE_ACUMULADO"),
        ("Y610", "Bono salarial", True, True, "VARIABLE_ACUMULADO"),
        ("Y617", "Bono salarial", True, True, "VARIABLE_ACUMULADO"),
        ("Y618", "Bono salarial", True, True, "VARIABLE_ACUMULADO"),
    ],
    columns=["Concepto", "Descripción", "Base_Prima", "Base_Cesantias", "Tipo_Base"],
)

DEFAULT_AREA_RULES = pd.DataFrame(
    [
        ("ZM", "Administrativos / base 360", "DIAS_360", 30, "Salario mensual ponderado por días 360"),
        ("ZL", "Mensual admon 365 / días reales", "DIAS_CALENDARIO", 30, "Salario mensual ponderado por días calendario"),
        ("ZH", "Part time horas", "DIAS_CALENDARIO", 30, "ZH se trata con base 365 para promedio"),
        ("ZP", "Part time días", "DIAS_CALENDARIO", 30, "ZP se trata con base 365 para promedio"),
    ],
    columns=["Área de nómina", "Descripción", "Regla_Dias", "Mensualizar_A", "Observación"],
)

EXCEL_EXTS = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}
CSV_EXTS = {".csv", ".txt"}

SAP_CANDIDATES = [
    "sap", "nº pers", "n° pers", "no pers", "nro pers", "numero personal",
    "número personal", "pernr", "cod sap", "codigo sap", "código sap", "usuario",
    "employee id", "id empleado", "colaborador", "n pers", "personal", "nro personal",
]
CONCEPT_CANDIDATES = [
    "cc-nomina", "cc nomina", "cc-nómina", "cc nómina", "concepto", "codigo concepto",
    "código concepto", "cod concepto", "lgart", "clase de nomina", "clase de nómina",
]
CONCEPT_TEXT_CANDIDATES = [
    "texto", "descripcion", "descripción", "desc concepto", "descripcion concepto",
    "descripción concepto", "texto concepto", "nombre concepto", "concepto texto",
]
VALUE_CANDIDATES = [
    "valor", "importe", "monto", "devengo", "valor pago", "valor pagado", "total",
    "amount", "valor concepto", "importe moneda", "valor ml", "valor acumulado",
]
QTY_CANDIDATES = ["cantidad", "cant", "dias", "días", "horas", "numero", "número"]
PERIOD_CANDIDATES = [
    "periodo", "período", "periodo nomina", "período nómina", "periodo para nomina",
    "periodo para nómina", "fecha de pago", "fecha pago", "mes", "mes pago", "for-period",
    "periodo pago", "período pago", "fecha contabilizacion", "fecha contabilización",
]

MD_NAME_CANDIDATES = ["nombre", "nombres", "nombre completo", "empleado", "trabajador"]
MD_DOC_CANDIDATES = ["cedula", "cédula", "numero id", "número id", "documento", "id", "identificacion", "identificación"]
MD_AREA_CANDIDATES = ["area de nomina", "área de nómina", "area nomina", "área nómina", "area cálculo nómina", "area calculo nomina"]
MD_CECO_CANDIDATES = ["ce.coste", "ce coste", "ceco", "centro de coste", "centro de costo"]
MD_CARGO_CANDIDATES = ["cargo", "funcion", "función", "posicion", "posición"]
MD_SALARY_CANDIDATES = ["salario total", "salario", "sueldo", "sueldo basico", "sueldo básico"]
MD_INGRESO_CANDIDATES = ["fecha ingreso", "fecha de ingreso", "ingreso", "alta", "fecha alta"]
MD_RETIRO_CANDIDATES = ["fecha retiro", "fecha de retiro", "retiro", "baja", "fecha baja"]

SAL_FROM_CANDIDATES = [
    "fecha desde", "desde", "inicio", "fecha inicio", "vigencia desde", "válido desde", "valido desde",
    "fecha inicial", "begda", "inicio vigencia",
]
SAL_TO_CANDIDATES = [
    "fecha hasta", "hasta", "fin", "fecha fin", "vigencia hasta", "válido hasta", "valido hasta",
    "fecha final", "endda", "fin vigencia",
]
SAL_VALUE_CANDIDATES = [
    "salario total", "salario", "sueldo", "sueldo basico", "sueldo básico", "importe", "valor", "valor salario",
]
AREA_RULE_CANDIDATES = ["regla dias", "regla_dias", "regla días", "tipo dias", "tipo días", "metodo", "método"]
MONTHLY_TO_CANDIDATES = ["mensualizar a", "mensualizar_a", "mensualizacion", "mensualización", "dias mensualizar", "días mensualizar"]

# -----------------------------
# Utilidades
# -----------------------------

def normalize_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = text.replace("_", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text)
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
    xls = pd.ExcelFile(io.BytesIO(data), engine=excel_engine_for_name(uploaded_file.name))
    return xls.sheet_names


def detect_header_row(raw_preview: pd.DataFrame) -> int:
    keywords = (
        SAP_CANDIDATES + CONCEPT_CANDIDATES + VALUE_CANDIDATES + PERIOD_CANDIDATES
        + MD_NAME_CANDIDATES + MD_CECO_CANDIDATES + SAL_FROM_CANDIDATES + SAL_TO_CANDIDATES
    )
    norm_keywords = [normalize_text(k) for k in keywords]
    best_row, best_score = 0, -1
    for idx in range(min(len(raw_preview), 40)):
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
            best_row, best_score = idx, score
    return best_row if best_score > 0 else 0


def read_uploaded_table(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    ext = get_extension(uploaded_file.name)
    data = uploaded_file.getvalue()

    if ext in CSV_EXTS:
        best, best_cols = None, -1
        for enc in ["utf-8-sig", "latin1", "cp1252"]:
            for sep in [";", ",", "\t", "|"]:
                try:
                    df = pd.read_csv(io.BytesIO(data), sep=sep, encoding=enc, dtype=str, engine="python")
                    df = clean_columns(df)
                    if len(df.columns) > best_cols:
                        best, best_cols = df, len(df.columns)
                except Exception:
                    pass
        if best is None:
            raise ValueError("No fue posible leer el archivo plano. Revisa separador o codificación.")
        return best

    if ext in EXCEL_EXTS:
        engine = excel_engine_for_name(uploaded_file.name)
        sheet = sheet_name if sheet_name and sheet_name != "Archivo plano" else 0
        preview = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, nrows=40, engine=engine)
        header_row = detect_header_row(preview)
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=header_row, engine=engine, dtype=object)
        return clean_columns(df)

    raise ValueError(f"Extensión no soportada: {ext}.")


def normalize_sap(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    digits = re.sub(r"\D", "", text)
    if digits:
        return digits.lstrip("0") or "0"
    return text.strip()


def normalize_area(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).upper().strip()
    m = re.search(r"\b(ZM|ZL|ZH|ZP)\b", text)
    if m:
        return m.group(1)
    # Si viene como texto largo, buscar la sigla pegada.
    for area in ["ZM", "ZL", "ZH", "ZP"]:
        if area in text:
            return area
    return text


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
    s = s.replace("_", "/").replace("-", "/").replace(".", "/")
    s = re.sub(r"\s+", "", s)
    m = re.fullmatch(r"(\d{6})", s)
    if m:
        val = m.group(1)
        y1, m1 = int(val[:4]), int(val[4:])
        if 1900 <= y1 <= 2100 and 1 <= m1 <= 12:
            return pd.Timestamp(year=y1, month=m1, day=1)
        m2, y2 = int(val[:2]), int(val[2:])
        if 1 <= m2 <= 12 and 1900 <= y2 <= 2100:
            return pd.Timestamp(year=y2, month=m2, day=1)
    m = re.fullmatch(r"(\d{1,4})/(\d{1,4})", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if 1900 <= a <= 2100 and 1 <= b <= 12:
            return pd.Timestamp(year=a, month=b, day=1)
        if 1 <= a <= 12 and 1900 <= b <= 2100:
            return pd.Timestamp(year=b, month=a, day=1)
    dt = pd.to_datetime(str(value), dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return pd.Timestamp(year=dt.year, month=dt.month, day=1)
    return pd.NaT


def month_last_day(d: date) -> date:
    if d.month == 12:
        return date(d.year, 12, 31)
    return date(d.year, d.month + 1, 1) - timedelta(days=1)


def add_months(first_day: date, months: int) -> date:
    y = first_day.year + (first_day.month - 1 + months) // 12
    m = (first_day.month - 1 + months) % 12 + 1
    return date(y, m, 1)


def days360_colombian(start_d: date, end_d: date) -> int:
    """Conteo 360 simple inclusivo: cada mes máximo 30 días."""
    if end_d < start_d:
        return 0
    y1, m1, d1 = start_d.year, start_d.month, min(start_d.day, 30)
    y2, m2, d2 = end_d.year, end_d.month, min(end_d.day, 30)
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1) + 1


def calendar_days(start_d: date, end_d: date) -> int:
    if end_d < start_d:
        return 0
    return (end_d - start_d).days + 1


def area_days(start_d: date, end_d: date, area: str, area_rules: pd.DataFrame) -> int:
    area_norm = normalize_area(area)
    rules = area_rules.copy()
    rules["_area"] = rules["Área de nómina"].apply(normalize_area)
    match = rules[rules["_area"] == area_norm]
    rule = "DIAS_CALENDARIO"
    if not match.empty:
        rule = str(match.iloc[0].get("Regla_Dias", "DIAS_CALENDARIO")).upper().strip()
    if "360" in rule:
        return days360_colombian(start_d, end_d)
    return calendar_days(start_d, end_d)


def month_count_inclusive(start_d: date, end_d: date) -> int:
    if end_d < start_d:
        return 0
    return (end_d.year - start_d.year) * 12 + (end_d.month - start_d.month) + 1


def safe_bool(value) -> bool:
    if value is None or pd.isna(value):
        return False
    s = normalize_text(value)
    return s in {"si", "s", "x", "true", "verdadero", "1", "aplica", "yes", "y"}

# -----------------------------
# Estandarización de archivos
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
    out["Periodo_Mes"] = df[period_col].apply(parse_period_value)
    out["Periodo_Original"] = df[period_col].astype(str)
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
    ingreso_col = find_column(df, MD_INGRESO_CANDIDATES)
    retiro_col = find_column(df, MD_RETIRO_CANDIDATES)

    out = pd.DataFrame()
    out["SAP"] = df[sap_col].apply(normalize_sap)
    out["Cédula"] = df[doc_col].astype(str).str.strip() if doc_col else ""
    out["Nombre"] = df[name_col].astype(str).str.strip() if name_col else ""
    out["Área de nómina"] = df[area_col].apply(normalize_area) if area_col else ""
    out["CECO"] = df[ceco_col].astype(str).str.strip() if ceco_col else ""
    out["Cargo"] = df[cargo_col].astype(str).str.strip() if cargo_col else ""
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
        "Fecha ingreso": ingreso_col or "No detectada",
        "Fecha retiro": retiro_col or "No detectada",
    }
    return out, detected


def standardize_salary_history(df: pd.DataFrame, md_std: Optional[pd.DataFrame] = None) -> Tuple[pd.DataFrame, Dict[str, str]]:
    sap_col = find_column(df, SAP_CANDIDATES)
    area_col = find_column(df, MD_AREA_CANDIDATES)
    sal_col = find_column(df, SAL_VALUE_CANDIDATES)
    from_col = find_column(df, SAL_FROM_CANDIDATES)
    to_col = find_column(df, SAL_TO_CANDIDATES)

    missing = []
    if not sap_col:
        missing.append("SAP / Nº pers.")
    if not sal_col:
        missing.append("Salario")
    if not from_col:
        missing.append("Fecha desde / inicio vigencia")
    if not to_col:
        missing.append("Fecha hasta / fin vigencia")
    if missing:
        raise ValueError("No encontré estas columnas en histórico de salarios: " + ", ".join(missing))

    out = pd.DataFrame()
    out["SAP"] = df[sap_col].apply(normalize_sap)
    out["Área de nómina"] = df[area_col].apply(normalize_area) if area_col else ""
    out["Salario Vigencia"] = to_number_series(df[sal_col])
    out["Desde"] = parse_date_series(df[from_col])
    out["Hasta"] = parse_date_series(df[to_col])
    out = out[(out["SAP"] != "") & (out["Salario Vigencia"] > 0)]
    out = out[pd.notna(out["Desde"])]
    out.loc[pd.isna(out["Hasta"]), "Hasta"] = pd.Timestamp(9999, 12, 31)

    if md_std is not None and not md_std.empty:
        area_map = md_std[["SAP", "Área de nómina"]].drop_duplicates("SAP")
        out = out.merge(area_map.rename(columns={"Área de nómina": "Área MD"}), on="SAP", how="left")
        out["Área de nómina"] = out["Área de nómina"].where(out["Área de nómina"].astype(str).str.strip() != "", out["Área MD"].fillna(""))
        out = out.drop(columns=["Área MD"])

    detected = {
        "SAP": sap_col,
        "Área de nómina": area_col or "No detectada; se completó con MD si existía",
        "Salario": sal_col,
        "Desde": from_col,
        "Hasta": to_col,
    }
    return out, detected


def standardize_concepts_param(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    concept_col = find_column(df, CONCEPT_CANDIDATES + ["codigo", "código"])
    desc_col = find_column(df, CONCEPT_TEXT_CANDIDATES)
    prima_col = find_column(df, ["base prima", "prima", "aplica prima"])
    ces_col = find_column(df, ["base cesantias", "base cesantías", "cesantias", "cesantías", "aplica cesantias", "aplica cesantías"])
    tipo_col = find_column(df, ["tipo base", "tipo_base", "tipo", "clasificacion", "clasificación"])
    if not concept_col:
        raise ValueError("La parametrización de conceptos debe tener Concepto / CC-nómina.")

    out = pd.DataFrame()
    out["Concepto"] = df[concept_col].apply(extract_concept)
    out["Descripción"] = df[desc_col].astype(str).str.strip() if desc_col else ""
    out["Base_Prima"] = df[prima_col].apply(safe_bool) if prima_col else True
    out["Base_Cesantias"] = df[ces_col].apply(safe_bool) if ces_col else True
    out["Tipo_Base"] = df[tipo_col].astype(str).str.upper().str.strip() if tipo_col else "VARIABLE_ACUMULADO"
    out.loc[out["Concepto"].isin(FIXED_SALARY_CONCEPTS), "Tipo_Base"] = "SALARIO_FIJO_HISTORICO"
    out = out[out["Concepto"] != ""].drop_duplicates("Concepto", keep="last")
    detected = {
        "Concepto": concept_col,
        "Descripción": desc_col or "No detectada",
        "Base Prima": prima_col or "No detectada; se asumió Sí",
        "Base Cesantías": ces_col or "No detectada; se asumió Sí",
        "Tipo Base": tipo_col or "No detectada; salario fijo se identificó por Y010/Y011/Y020/Y050/Y051/Y090",
    }
    return out, detected


def standardize_area_rules(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    area_col = find_column(df, MD_AREA_CANDIDATES + ["area", "área"])
    desc_col = find_column(df, ["descripcion", "descripción"])
    rule_col = find_column(df, AREA_RULE_CANDIDATES)
    monthly_col = find_column(df, MONTHLY_TO_CANDIDATES)
    obs_col = find_column(df, ["observacion", "observación", "comentario"])
    if not area_col:
        raise ValueError("La parametrización de áreas debe tener columna Área de nómina.")

    out = pd.DataFrame()
    out["Área de nómina"] = df[area_col].apply(normalize_area)
    out["Descripción"] = df[desc_col].astype(str).str.strip() if desc_col else ""
    out["Regla_Dias"] = df[rule_col].astype(str).str.upper().str.strip() if rule_col else "DIAS_CALENDARIO"
    out["Mensualizar_A"] = to_number_series(df[monthly_col]) if monthly_col else 30
    out["Observación"] = df[obs_col].astype(str).str.strip() if obs_col else ""
    out.loc[out["Regla_Dias"].str.contains("360", na=False), "Regla_Dias"] = "DIAS_360"
    out.loc[~out["Regla_Dias"].str.contains("360", na=False), "Regla_Dias"] = "DIAS_CALENDARIO"
    out.loc[out["Mensualizar_A"] <= 0, "Mensualizar_A"] = 30
    out = out[out["Área de nómina"] != ""].drop_duplicates("Área de nómina", keep="last")
    detected = {
        "Área de nómina": area_col,
        "Descripción": desc_col or "No detectada",
        "Regla días": rule_col or "No detectada; se asumió calendario salvo parametrización default",
        "Mensualizar a": monthly_col or "No detectada; se asumió 30",
        "Observación": obs_col or "No detectada",
    }
    return out, detected

# -----------------------------
# Cálculos principales
# -----------------------------

def build_population(accum: pd.DataFrame, salary_hist: pd.DataFrame, md_std: Optional[pd.DataFrame]) -> pd.DataFrame:
    saps = sorted(set(accum["SAP"].astype(str)) | set(salary_hist["SAP"].astype(str)))
    pop = pd.DataFrame({"SAP": saps})
    if md_std is not None and not md_std.empty:
        pop = pop.merge(md_std, on="SAP", how="left")
    else:
        for col in ["Cédula", "Nombre", "Área de nómina", "CECO", "Cargo"]:
            pop[col] = ""
        pop["Fecha ingreso"] = pd.NaT
        pop["Fecha retiro"] = pd.NaT

    # Completar área desde histórico si no existe en MD.
    area_hist = salary_hist[salary_hist["Área de nómina"].astype(str).str.strip() != ""]
    if not area_hist.empty:
        last_area = area_hist.sort_values("Desde").drop_duplicates("SAP", keep="last")[["SAP", "Área de nómina"]]
        pop = pop.merge(last_area.rename(columns={"Área de nómina": "Área histórico"}), on="SAP", how="left")
        pop["Área de nómina"] = pop["Área de nómina"].fillna("")
        pop["Área de nómina"] = pop["Área de nómina"].where(pop["Área de nómina"].astype(str).str.strip() != "", pop["Área histórico"].fillna(""))
        pop = pop.drop(columns=["Área histórico"])
    return pop


def calc_salary_average(
    salary_hist: pd.DataFrame,
    population: pd.DataFrame,
    area_rules: pd.DataFrame,
    period_start: date,
    period_end: date,
    label: str,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    rows = []
    for _, r in salary_hist.iterrows():
        sap = r["SAP"]
        area = normalize_area(r.get("Área de nómina", ""))
        desde = r["Desde"]
        hasta = r["Hasta"]
        salario = float(r.get("Salario Vigencia", 0) or 0)
        if pd.isna(desde) or salario <= 0:
            continue
        ini = max(period_start, desde.date())
        fin_raw = hasta.date() if pd.notna(hasta) else date(9999, 12, 31)
        fin = min(period_end, fin_raw)
        if fin < ini:
            continue
        dias = area_days(ini, fin, area, area_rules)
        if dias <= 0:
            continue
        rows.append(
            {
                "SAP": sap,
                "Etiqueta": label,
                "Área de nómina salario": area,
                "Desde tramo": ini,
                "Hasta tramo": fin,
                "Salario Vigencia": salario,
                "Días salario": dias,
                "Salario x días": salario * dias,
            }
        )

    detail = pd.DataFrame(rows)
    if detail.empty:
        out = population[["SAP"]].copy()
        out[f"Salario histórico promedio {label}"] = 0.0
        out[f"Días salario histórico {label}"] = 0
        return out, detail

    grouped = detail.groupby("SAP", as_index=False).agg(
        **{
            f"Salario x días {label}": ("Salario x días", "sum"),
            f"Días salario histórico {label}": ("Días salario", "sum"),
            f"Tramos salario {label}": ("Salario Vigencia", "size"),
        }
    )
    grouped[f"Salario histórico promedio {label}"] = (
        grouped[f"Salario x días {label}"] / grouped[f"Días salario histórico {label}"].replace(0, pd.NA)
    ).fillna(0.0)
    return grouped, detail


def calc_accum_average(
    accum: pd.DataFrame,
    population: pd.DataFrame,
    concepts: List[str],
    period_start: date,
    period_end: date,
    label: str,
    area_rules: pd.DataFrame,
    only_variable: bool,
    concept_param: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    concepts_set = set(concepts)
    p_start_ts = pd.Timestamp(period_start.year, period_start.month, 1)
    p_end_ts = pd.Timestamp(period_end.year, period_end.month, 1)

    used = accum[
        (accum["Periodo_Mes"] >= p_start_ts)
        & (accum["Periodo_Mes"] <= p_end_ts)
        & (accum["Concepto"].isin(concepts_set))
    ].copy()

    type_map = concept_param[["Concepto", "Tipo_Base"]].drop_duplicates("Concepto") if not concept_param.empty else pd.DataFrame(columns=["Concepto", "Tipo_Base"])
    used = used.merge(type_map, on="Concepto", how="left")
    used["Tipo_Base"] = used["Tipo_Base"].fillna("VARIABLE_ACUMULADO")
    if only_variable:
        used = used[used["Tipo_Base"] != "SALARIO_FIJO_HISTORICO"].copy()

    if used.empty:
        grouped = pd.DataFrame({"SAP": population["SAP"]})
        grouped[f"Valor acumulado {'variable ' if only_variable else 'total '} {label}"] = 0.0
        grouped[f"Meses con acumulado {'variable ' if only_variable else 'total '} {label}"] = 0
        detail = used
    else:
        grouped = used.groupby("SAP", as_index=False).agg(
            **{
                f"Valor acumulado {'variable ' if only_variable else 'total '} {label}": ("Valor", "sum"),
                f"Meses con acumulado {'variable ' if only_variable else 'total '} {label}": ("Periodo_Mes", lambda s: s.dt.strftime("%Y-%m").nunique()),
            }
        )
        detail = used

    out = population[["SAP", "Área de nómina"]].merge(grouped, on="SAP", how="left")
    value_col = f"Valor acumulado {'variable ' if only_variable else 'total '} {label}"
    months_col = f"Meses con acumulado {'variable ' if only_variable else 'total '} {label}"
    out[value_col] = out[value_col].fillna(0.0)
    out[months_col] = out[months_col].fillna(0).astype(int)

    # Divisor por área de nómina. Para periodos completos full, esto equivale a /6 o /3 con base 30 cuando aplique.
    out[f"Días divisor acumulados {label}"] = out["Área de nómina"].apply(lambda a: area_days(period_start, period_end, a, area_rules))
    out[f"Meses divisor acumulados {label}"] = month_count_inclusive(period_start, period_end)
    out[f"Promedio {'variable' if only_variable else 'total acumulado'} {label}"] = (
        out[value_col] / out[f"Días divisor acumulados {label}"].replace(0, pd.NA) * 30
    ).fillna(0.0)
    return out.drop(columns=["Área de nómina"]), detail


def calc_base_for_label(
    label: str,
    population: pd.DataFrame,
    salary_hist: pd.DataFrame,
    accum: pd.DataFrame,
    concepts: List[str],
    period_start: date,
    period_end: date,
    area_rules: pd.DataFrame,
    concept_param: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    salary_avg, salary_detail = calc_salary_average(salary_hist, population, area_rules, period_start, period_end, label)
    variable_avg, variable_detail = calc_accum_average(
        accum, population, concepts, period_start, period_end, label, area_rules, True, concept_param
    )
    total_accum_avg, total_detail = calc_accum_average(
        accum, population, concepts, period_start, period_end, label, area_rules, False, concept_param
    )

    base = population.merge(salary_avg, on="SAP", how="left").merge(variable_avg, on="SAP", how="left").merge(total_accum_avg, on="SAP", how="left")

    sal_col = f"Salario histórico promedio {label}"
    var_col = f"Promedio variable {label}"
    total_col = f"Promedio total acumulado {label}"
    base[sal_col] = base[sal_col].fillna(0.0)
    base[var_col] = base[var_col].fillna(0.0)
    base[total_col] = base[total_col].fillna(0.0)
    base[f"Base calculada {label}"] = base[sal_col] + base[var_col]
    base[f"Base SAP acumulada estimada {label}"] = base[total_col]
    base[f"Diferencia vs acumulados {label}"] = base[f"Base calculada {label}"] - base[f"Base SAP acumulada estimada {label}"]
    base[f"Periodo inicial {label}"] = period_start.strftime("%d/%m/%Y")
    base[f"Periodo final {label}"] = period_end.strftime("%d/%m/%Y")

    status = []
    for _, r in base.iterrows():
        notes = []
        if float(r.get(sal_col, 0) or 0) == 0:
            notes.append("Sin salario histórico en periodo")
        if str(r.get("Área de nómina", "")).strip() == "":
            notes.append("Sin área de nómina para aplicar regla")
        if float(r.get(f"Días salario histórico {label}", 0) or 0) == 0:
            notes.append("Sin días de salario histórico")
        if not notes:
            notes.append("OK")
        status.append(" | ".join(notes))
    base[f"Estado {label}"] = status

    ordered = [
        "SAP", "Cédula", "Nombre", "Área de nómina", "CECO", "Cargo", "Fecha ingreso", "Fecha retiro",
        f"Periodo inicial {label}", f"Periodo final {label}",
        sal_col, f"Días salario histórico {label}", f"Tramos salario {label}",
        f"Valor acumulado variable  {label}", f"Meses con acumulado variable  {label}",
        f"Promedio variable {label}",
        f"Valor acumulado total  {label}", f"Meses con acumulado total  {label}",
        f"Promedio total acumulado {label}",
        f"Base calculada {label}", f"Base SAP acumulada estimada {label}", f"Diferencia vs acumulados {label}",
        f"Estado {label}",
    ]
    # Nombres generados tienen doble espacio por el sufijo usado; normalizar visualmente.
    rename_map = {
        f"Valor acumulado variable  {label}": f"Valor acumulado variable {label}",
        f"Meses con acumulado variable  {label}": f"Meses con acumulado variable {label}",
        f"Valor acumulado total  {label}": f"Valor acumulado total {label}",
        f"Meses con acumulado total  {label}": f"Meses con acumulado total {label}",
    }
    base = base.rename(columns=rename_map)
    ordered = [rename_map.get(c, c) for c in ordered]
    return base[[c for c in ordered if c in base.columns]], salary_detail, variable_detail


def build_detail_accum(accum: pd.DataFrame, concept_param: pd.DataFrame, prima_concepts: List[str], ces_concepts: List[str], prima_start: date, prima_end: date, ces_start: date, ces_end: date) -> pd.DataFrame:
    detail = accum.copy()
    detail = detail.merge(concept_param[["Concepto", "Base_Prima", "Base_Cesantias", "Tipo_Base"]], on="Concepto", how="left")
    detail["Base_Prima"] = detail["Base_Prima"].fillna(False)
    detail["Base_Cesantias"] = detail["Base_Cesantias"].fillna(False)
    detail["Tipo_Base"] = detail["Tipo_Base"].fillna("NO_PARAMETRIZADO")
    detail["Periodo"] = detail["Periodo_Mes"].dt.strftime("%Y-%m")
    detail["Usado_Prima"] = (
        (detail["Periodo_Mes"] >= pd.Timestamp(prima_start.year, prima_start.month, 1))
        & (detail["Periodo_Mes"] <= pd.Timestamp(prima_end.year, prima_end.month, 1))
        & (detail["Concepto"].isin(prima_concepts))
    )
    detail["Usado_Cesantias"] = (
        (detail["Periodo_Mes"] >= pd.Timestamp(ces_start.year, ces_start.month, 1))
        & (detail["Periodo_Mes"] <= pd.Timestamp(ces_end.year, ces_end.month, 1))
        & (detail["Concepto"].isin(ces_concepts))
    )
    cols = [
        "SAP", "Periodo", "Periodo_Original", "Concepto", "Texto Concepto", "Valor", "Cantidad",
        "Tipo_Base", "Base_Prima", "Base_Cesantias", "Usado_Prima", "Usado_Cesantias",
    ]
    return detail[cols]


def build_alerts(base_prima: pd.DataFrame, base_ces: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for df, label in [(base_prima, "Prima"), (base_ces, "Cesantias")]:
        status_col = f"Estado {label}"
        diff_col = f"Diferencia vs acumulados {label}"
        base_col = f"Base calculada {label}"
        for _, r in df.iterrows():
            estado = str(r.get(status_col, ""))
            diff = float(r.get(diff_col, 0) or 0)
            base = float(r.get(base_col, 0) or 0)
            pct = abs(diff) / abs(base) if base else 0
            if estado != "OK" or abs(diff) >= 1000 and pct >= 0.01:
                rows.append({
                    "SAP": r.get("SAP", ""),
                    "Nombre": r.get("Nombre", ""),
                    "Área de nómina": r.get("Área de nómina", ""),
                    "Base": label,
                    "Estado": estado,
                    "Base calculada": base,
                    "Diferencia": diff,
                    "% Diferencia": pct,
                })
    return pd.DataFrame(rows)


def make_template_bytes() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        DEFAULT_CONCEPTS.to_excel(writer, index=False, sheet_name="Conceptos")
        DEFAULT_AREA_RULES.to_excel(writer, index=False, sheet_name="Areas")
    return output.getvalue()


def make_excel_report(
    llevar_modelo: pd.DataFrame,
    base_prima: pd.DataFrame,
    base_ces: pd.DataFrame,
    detail_accum: pd.DataFrame,
    sal_detail: pd.DataFrame,
    concept_param: pd.DataFrame,
    area_rules: pd.DataFrame,
    alerts: pd.DataFrame,
    log_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        sheets = {
            "Llevar_al_Modelo": llevar_modelo,
            "Base_Prima": base_prima,
            "Base_Cesantias": base_ces,
            "Detalle_Acumulados": detail_accum,
            "Historico_Salarios_Calculo": sal_detail,
            "Param_Conceptos": concept_param,
            "Param_Areas": area_rules,
            "Alertas": alerts,
            "Log": log_df,
        }
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        int_fmt = workbook.add_format({"num_format": "#,##0"})
        pct_fmt = workbook.add_format({"num_format": "0.00%"})
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})

        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            if df is None or df.empty:
                df = pd.DataFrame({"Sin registros": []})
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, header_fmt)
                width = min(max(len(str(col_name)) + 2, 12), 50)
                if not df.empty:
                    # Evita TypeError con pandas/pyarrow cuando hay valores nulos (pd.NA),
                    # fechas fuera de rango, listas u objetos no convertibles por map(len).
                    def _safe_text_len(value):
                        try:
                            if value is None:
                                return 0
                            if pd.isna(value):
                                return 0
                        except Exception:
                            pass
                        try:
                            return len(str(value))
                        except Exception:
                            return 0

                    sample = df[col_name].head(200).map(_safe_text_len).max()
                    if pd.isna(sample):
                        sample = 0
                    width = min(max(width, int(sample) + 2), 50)
                lower = str(col_name).lower()
                fmt = None
                if any(k in lower for k in ["valor", "base", "salario", "promedio", "acumulado", "diferencia"]):
                    fmt = money_fmt
                elif any(k in lower for k in ["dias", "días", "meses", "tramos"]):
                    fmt = int_fmt
                elif "%" in lower:
                    fmt = pct_fmt
                elif "fecha" in lower or "desde" in lower or "hasta" in lower:
                    fmt = date_fmt
                ws.set_column(col_idx, col_idx, width, fmt)
    return output.getvalue()

# -----------------------------
# UI
# -----------------------------

st.title("🦜 Validador de bases de prima y cesantías")
st.caption("Modelo financiero nómina · Acumulados históricos + histórico de salarios · Sin proyección")

with st.expander("📌 Qué hace esta versión", expanded=True):
    st.markdown(
        """
        Esta versión **no toma pagado del mes de proyección** ni calcula proyección.  
        Hace la validación de la base así:

        **Base calculada = salario histórico promedio por vigencias + variables promedio de acumulados.**

        El salario fijo sale del **histórico de salarios**. Por eso los conceptos `Y010`, `Y011`, `Y020`, `Y050`, `Y051` y `Y090` se tratan como salario fijo histórico y no se duplican dentro de las variables acumuladas.
        """
    )

with st.sidebar:
    st.header("⚙️ Parámetros")
    corte = st.date_input("Fecha de corte de validación", value=date.today().replace(day=1) - timedelta(days=1))

    st.caption("Periodos automáticos sugeridos")
    use_auto = st.checkbox("Usar últimos 6 meses para prima y últimos 3 meses para cesantías", value=True)
    if use_auto:
        corte_month_first = date(corte.year, corte.month, 1)
        prima_start = add_months(corte_month_first, -5)
        prima_end = month_last_day(corte_month_first)
        ces_start = add_months(corte_month_first, -2)
        ces_end = month_last_day(corte_month_first)
    else:
        prima_start = st.date_input("Inicio periodo prima", value=add_months(date(corte.year, corte.month, 1), -5))
        prima_end = st.date_input("Fin periodo prima", value=corte)
        ces_start = st.date_input("Inicio periodo cesantías", value=add_months(date(corte.year, corte.month, 1), -2))
        ces_end = st.date_input("Fin periodo cesantías", value=corte)

    st.download_button(
        "⬇️ Descargar plantilla de parametrización",
        data=make_template_bytes(),
        file_name="plantilla_parametrizacion_bases_prima_cesantias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.subheader("1) Carga de archivos")
col1, col2, col3, col4 = st.columns(4)
with col1:
    accum_file = st.file_uploader("Acumulados de nómina", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"])
with col2:
    salary_file = st.file_uploader("Histórico de salarios", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"])
with col3:
    md_file = st.file_uploader("Master Data / empleados (opcional)", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"])
with col4:
    param_file = st.file_uploader("Parametrización conceptos/áreas (opcional)", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "txt"])


def sheet_selector(file, key):
    if not file:
        return None
    try:
        sheets = get_sheet_names(file)
        return st.selectbox(f"Hoja {key}", sheets, key=f"sheet_{key}") if len(sheets) > 1 else sheets[0]
    except Exception as exc:
        st.error(f"No pude leer hojas de {key}: {exc}")
        return None

accum_sheet = sheet_selector(accum_file, "acumulados")
salary_sheet = sheet_selector(salary_file, "histórico salarios")
md_sheet = sheet_selector(md_file, "MD")
param_sheet = sheet_selector(param_file, "parametrización")

read_log = []
accum_std = None
salary_std = None
md_std = None
concept_param = DEFAULT_CONCEPTS.copy()
area_rules = DEFAULT_AREA_RULES.copy()

if md_file and md_sheet:
    try:
        md_raw = read_uploaded_table(md_file, md_sheet)
        md_std, md_detected = standardize_md(md_raw)
        read_log.append({"Paso": "Master Data", "Resultado": f"OK - {len(md_std):,} empleados", "Detalle": str(md_detected)})
        st.success(f"Master Data leído: {len(md_std):,} empleados")
        with st.expander("Columnas detectadas en MD"):
            st.json(md_detected)
    except Exception as exc:
        st.warning(f"No se usará MD: {exc}")
        md_std = None

if accum_file and accum_sheet:
    try:
        accum_raw = read_uploaded_table(accum_file, accum_sheet)
        accum_std, accum_detected = standardize_accumulated(accum_raw)
        read_log.append({"Paso": "Acumulados", "Resultado": f"OK - {len(accum_std):,} registros", "Detalle": str(accum_detected)})
        st.success(f"Acumulados leídos: {len(accum_std):,} registros")
        with st.expander("Columnas detectadas en acumulados"):
            st.json(accum_detected)
    except Exception as exc:
        st.error(f"Error leyendo acumulados: {exc}")

if salary_file and salary_sheet:
    try:
        salary_raw = read_uploaded_table(salary_file, salary_sheet)
        salary_std, salary_detected = standardize_salary_history(salary_raw, md_std)
        read_log.append({"Paso": "Histórico salarios", "Resultado": f"OK - {len(salary_std):,} vigencias", "Detalle": str(salary_detected)})
        st.success(f"Histórico de salarios leído: {len(salary_std):,} vigencias")
        with st.expander("Columnas detectadas en histórico de salarios"):
            st.json(salary_detected)
    except Exception as exc:
        st.error(f"Error leyendo histórico de salarios: {exc}")

if param_file and param_sheet:
    try:
        param_raw = read_uploaded_table(param_file, param_sheet)
        # El mismo archivo puede traer una hoja de conceptos o áreas. Si la hoja elegida parece áreas, carga áreas; si no, conceptos.
        try:
            param_area, area_detected = standardize_area_rules(param_raw)
            if set(param_area["Área de nómina"].dropna()) & {"ZM", "ZL", "ZH", "ZP"}:
                area_rules = pd.concat([DEFAULT_AREA_RULES, param_area], ignore_index=True).drop_duplicates("Área de nómina", keep="last")
                read_log.append({"Paso": "Param áreas", "Resultado": f"OK - {len(area_rules):,} reglas", "Detalle": str(area_detected)})
                st.success("Parametrización de áreas cargada")
        except Exception:
            pass
        try:
            param_concepts, concept_detected = standardize_concepts_param(param_raw)
            if not param_concepts.empty:
                concept_param = pd.concat([DEFAULT_CONCEPTS, param_concepts], ignore_index=True).drop_duplicates("Concepto", keep="last")
                read_log.append({"Paso": "Param conceptos", "Resultado": f"OK - {len(concept_param):,} conceptos", "Detalle": str(concept_detected)})
                st.success("Parametrización de conceptos cargada")
        except Exception:
            pass
    except Exception as exc:
        st.warning(f"No se usó parametrización cargada: {exc}")

st.subheader("2) Parametrización aplicada")
pa, pc = st.columns(2)
with pa:
    st.markdown("**Reglas por área de nómina**")
    area_rules = st.data_editor(area_rules, use_container_width=True, num_rows="dynamic")
with pc:
    st.markdown("**Conceptos para prima y cesantías**")
    concept_param = st.data_editor(concept_param, use_container_width=True, num_rows="dynamic")

if accum_std is not None and salary_std is not None:
    st.subheader("3) Validación antes de generar")
    prima_concepts = sorted(concept_param.loc[concept_param["Base_Prima"].astype(bool), "Concepto"].dropna().apply(extract_concept).unique().tolist())
    ces_concepts = sorted(concept_param.loc[concept_param["Base_Cesantias"].astype(bool), "Concepto"].dropna().apply(extract_concept).unique().tolist())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Periodo prima", f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}")
    with c2:
        st.metric("Periodo cesantías", f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}")
    with c3:
        st.metric("Conceptos prima", len(prima_concepts))
    with c4:
        st.metric("Conceptos cesantías", len(ces_concepts))

    periods_summary = (
        accum_std.groupby(accum_std["Periodo_Mes"].dt.strftime("%Y-%m"), as_index=True)
        .agg(Registros=("Valor", "size"), Valor=("Valor", "sum"))
        .reset_index()
        .rename(columns={"Periodo_Mes": "Periodo"})
    )
    with st.expander("Periodos encontrados en acumulados"):
        st.dataframe(periods_summary, use_container_width=True)
    with st.expander("Conceptos encontrados en acumulados"):
        catalog = accum_std.groupby(["Concepto", "Texto Concepto"], as_index=False).agg(Valor_Total=("Valor", "sum"), Registros=("Valor", "size"))
        catalog = catalog.merge(concept_param[["Concepto", "Base_Prima", "Base_Cesantias", "Tipo_Base"]], on="Concepto", how="left")
        st.dataframe(catalog, use_container_width=True)

    generate = st.button("🚀 Generar validación de bases", type="primary")

    if generate:
        if prima_end < prima_start or ces_end < ces_start:
            st.error("Revisa los periodos: la fecha final no puede ser menor que la inicial.")
            st.stop()
        if not prima_concepts or not ces_concepts:
            st.error("La parametrización debe dejar al menos un concepto para prima y uno para cesantías.")
            st.stop()

        progress = st.progress(0)
        status = st.empty()

        status.write("Preparando población...")
        population = build_population(accum_std, salary_std, md_std)
        progress.progress(15)

        status.write("Calculando prima con histórico salarial + acumulados...")
        base_prima, sal_detail_prima, var_detail_prima = calc_base_for_label(
            "Prima", population, salary_std, accum_std, prima_concepts, prima_start, prima_end, area_rules, concept_param
        )
        progress.progress(40)

        status.write("Calculando cesantías con histórico salarial + acumulados...")
        base_ces, sal_detail_ces, var_detail_ces = calc_base_for_label(
            "Cesantias", population, salary_std, accum_std, ces_concepts, ces_start, ces_end, area_rules, concept_param
        )
        progress.progress(65)

        status.write("Armando salida...")
        llevar_modelo = base_prima[[
            "SAP", "Cédula", "Nombre", "Área de nómina", "CECO", "Cargo",
            "Salario histórico promedio Prima", "Promedio variable Prima", "Base calculada Prima",
            "Base SAP acumulada estimada Prima", "Diferencia vs acumulados Prima", "Estado Prima",
        ]].merge(
            base_ces[[
                "SAP", "Salario histórico promedio Cesantias", "Promedio variable Cesantias", "Base calculada Cesantias",
                "Base SAP acumulada estimada Cesantias", "Diferencia vs acumulados Cesantias", "Estado Cesantias",
            ]],
            on="SAP", how="outer",
        )
        llevar_modelo["Periodo prima usado"] = f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}"
        llevar_modelo["Periodo cesantías usado"] = f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}"
        llevar_modelo["Observación"] = "Validación con acumulados históricos e histórico de salarios; sin proyección."

        detail_accum = build_detail_accum(accum_std, concept_param, prima_concepts, ces_concepts, prima_start, prima_end, ces_start, ces_end)
        sal_detail = pd.concat([sal_detail_prima, sal_detail_ces], ignore_index=True) if not sal_detail_prima.empty or not sal_detail_ces.empty else pd.DataFrame()
        alerts = build_alerts(base_prima, base_ces)
        progress.progress(80)

        log_df = pd.DataFrame(read_log + [
            {"Paso": "Periodo prima", "Resultado": f"{prima_start:%d/%m/%Y} - {prima_end:%d/%m/%Y}", "Detalle": "Últimos 6 meses si se dejó automático."},
            {"Paso": "Periodo cesantías", "Resultado": f"{ces_start:%d/%m/%Y} - {ces_end:%d/%m/%Y}", "Detalle": "Últimos 3 meses si se dejó automático."},
            {"Paso": "Reglas área", "Resultado": f"{len(area_rules):,} reglas", "Detalle": area_rules.to_dict(orient="records")},
            {"Paso": "Conceptos salario fijo", "Resultado": ", ".join(sorted(FIXED_SALARY_CONCEPTS)), "Detalle": "Se calculan desde histórico salarial, no se duplican como variable."},
            {"Paso": "Fórmula", "Resultado": "Base calculada = salario histórico promedio + promedio variable acumulado", "Detalle": "La base SAP acumulada estimada se deja como comparación."},
        ])

        excel_bytes = make_excel_report(
            llevar_modelo, base_prima, base_ces, detail_accum, sal_detail, concept_param, area_rules, alerts, log_df
        )
        progress.progress(100)
        status.success("Validación generada correctamente.")

        st.subheader("✅ Resultado")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.metric("Empleados", f"{len(llevar_modelo):,}")
        with r2:
            st.metric("Alertas", f"{len(alerts):,}")
        with r3:
            st.metric("Tramos salariales calculados", f"{len(sal_detail):,}")

        st.dataframe(llevar_modelo.head(300), use_container_width=True)
        st.download_button(
            "⬇️ Descargar Excel generado",
            data=excel_bytes,
            file_name=f"validacion_bases_prima_cesantias_{corte:%Y_%m_%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
else:
    st.info("Carga mínimo acumulados de nómina e histórico de salarios para continuar.")

st.divider()
st.caption("🦜 Creado por Andrés Huérfano Dávila - Nómina JMC")
