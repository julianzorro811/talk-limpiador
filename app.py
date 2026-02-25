import io
import os
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Limpieza Talk", page_icon="🧼", layout="wide")

st.title("🧼 Limpieza para carga en Talk")
st.caption(
    "Sube Excel/CSV/TXT (SIN celdas combinadas y con encabezados en la primera fila). "
    "Limpieza inteligente: numéricos, indicativos país, correos y truncamiento a 30 con reporte."
)

# ======================================================
# Lectura: Excel / CSV / TXT (autodetecta delimitador)
# ======================================================
def read_any_table(file_name: str, file_bytes: bytes) -> pd.DataFrame:
    ext = os.path.splitext(file_name.lower())[1]

    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(io.BytesIO(file_bytes), header=0, dtype=object)

    if ext in [".csv", ".txt"]:
        for sep in [";", ",", "\t", "|"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=sep, header=0, dtype=object, engine="python")
                if df.shape[1] == 1 and sep != "|":
                    continue
                return df
            except Exception:
                pass

        text = file_bytes.decode("utf-8", errors="replace").splitlines()
        return pd.DataFrame({"raw": text})

    raise ValueError(f"Extensión no soportada: {ext}")

# ======================================================
# Helpers base
# ======================================================
# Para TEXTO (Regla 3): se eliminan estos signos
DISALLOWED_SIGNS = r",\.\-\$\%\#\(\)\/"
INVISIBLE_CHARS_PATTERN = re.compile(r"[\uFEFF\u200B\u200C\u200D\u2060\u00AD]")

def remove_invisibles(s: str) -> str:
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    return INVISIBLE_CHARS_PATTERN.sub("", s)

def strip_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

def collapse_spaces(s: str) -> str:
    return re.sub(r"\s{2,}", " ", s)

def nompropio_like_excel(s: str) -> str:
    return s.lower().title()

# ======================================================
# Detección de tipo de columna (numérica / indicativo país)
# ======================================================
NUM_LIKE_RE = re.compile(r"^[\s\+\-]?[0-9\s\.,\*%$#()\/-]+$")  # permisivo para detectar

def is_numeric_like_value(v: object) -> bool:
    if pd.isna(v):
        return False
    s = remove_invisibles(str(v)).strip()
    if s == "":
        return False
    return bool(NUM_LIKE_RE.match(s))

def detect_numeric_columns(df: pd.DataFrame, sample_size: int = 80, threshold: float = 0.85) -> set:
    numeric_cols = set()
    for c in df.columns:
        series = df[c].dropna().astype(str)
        if series.empty:
            continue
        sample = series.head(sample_size)

        good, total = 0, 0
        for v in sample:
            v = str(v).strip()
            if v == "":
                continue
            total += 1
            if is_numeric_like_value(v):
                good += 1

        if total > 0 and (good / total) >= threshold:
            numeric_cols.add(c)
    return numeric_cols

def looks_like_country_code_value(v: object) -> bool:
    if pd.isna(v):
        return False
    s = remove_invisibles(str(v)).strip()
    if s == "":
        return False
    s2 = re.sub(r"\s+", "", s)
    return bool(re.fullmatch(r"\+?\d{1,4}", s2))

def detect_country_code_columns(df: pd.DataFrame, sample_size: int = 80, threshold: float = 0.75) -> set:
    cc_cols = set()
    for c in df.columns:
        name = str(c).strip().lower()

        # Heurística por nombre
        name_hit = any(k in name for k in [
            "indicativo", "codigo pais", "código pais", "código país",
            "pais", "country code", "prefijo", "prefijo pais"
        ])

        series = df[c].dropna().astype(str)
        if series.empty:
            continue
        sample = series.head(sample_size)

        good, total = 0, 0
        for v in sample:
            v = str(v).strip()
            if v == "":
                continue
            total += 1
            if looks_like_country_code_value(v):
                good += 1

        value_hit = (total > 0 and (good / total) >= threshold)

        if name_hit or value_hit:
            cc_cols.add(c)

    return cc_cols

# ======================================================
# Limpieza por tipo
# ======================================================
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def clean_email_one(val: object) -> tuple[str, str | None]:
    """
    Regla 5 robusta:
    - Encuentra TODOS los correos dentro del texto (aunque haya basura)
    - Toma el primero
    - Si hay más de uno, lo reporta como advertencia
    """
    if pd.isna(val) or str(val).strip() == "":
        return ("", "Correo vacío")

    raw = remove_invisibles(str(val)).strip()

    # Busca correos dentro del texto (aunque haya caracteres raros alrededor)
    found = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", raw)

    if not found:
        return ("", "No se encontró ningún correo válido en la celda")

    chosen = found[0].lower()

    # Advertencias
    warnings = []
    if len(found) > 1:
        warnings.append(f"Había {len(found)} correos; se tomó el primero")

    # Si la celda tiene más texto además del correo (basura), lo avisamos
    # (esto ayuda a detectar casos como '*!!\"#!')
    raw_compact = re.sub(r"\s+", " ", raw).strip()
    if raw_compact.lower() != chosen and len(raw_compact) > len(chosen):
        warnings.append("La celda contenía texto extra; se ignoró")

    return (chosen, " | ".join(warnings) if warnings else None)

def clean_text_general(val: object) -> str:
    """
    Reglas 1-3 para TEXTO (siempre NOMPROPIO):
    - trim
    - sin tildes
    - sin caracteres especiales (incluye , . - $ % # ( / ))
    - conserva espacios internos (colapsa múltiples a 1)
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val)).strip()
    if s == "":
        return ""

    s = strip_accents(s)
    s = re.sub(f"[{DISALLOWED_SIGNS}]", " ", s)  # quita signos listados
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)        # quita otros especiales
    s = collapse_spaces(s).strip()

    if s:
        s = nompropio_like_excel(s)

    return s

def clean_numeric_general(val: object) -> str:
    """
    Limpieza para columnas NUMÉRICAS:
    - trim
    - quita espacios internos
    - convierte coma a punto para decimal
    - elimina basura (*, $, etc)
    - conserva '.' como decimal
    Ej: ' 3 ,45* ' -> '3.45'
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val)).strip()
    if s == "":
        return ""

    s = re.sub(r"\s+", "", s)
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.\+\-]", "", s)

    if s.count(".") > 1:
        parts = s.split(".")
        s = parts[0] + "." + "".join(parts[1:])

    if s in ["+", "-", ".", "+.", "-."]:
        return ""

    return s

def clean_country_code(val: object) -> str:
    """
    Indicativo país:
    - conserva '+' si venía
    - deja solo + y dígitos
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val)).strip()
    if s == "":
        return ""
    s = re.sub(r"\s+", "", s)
    has_plus = s.startswith("+")
    digits = re.sub(r"\D", "", s)
    if digits == "":
        return ""
    return ("+" + digits) if has_plus else digits

def clean_apto_keep_inner_spaces(val: object) -> str:
    """
    Regla 4 (Apto):
    - quita espacios al inicio/fin
    - mantiene espacios internos intactos
    - quita tildes
    - elimina caracteres especiales (deja letras/números/espacios)
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val)).strip()
    if s == "":
        return ""
    s = strip_accents(s)
    s = re.sub(f"[{DISALLOWED_SIGNS}]", "", s)
    s = re.sub(r"[^A-Za-z0-9\s]", "", s)
    # NO colapsamos espacios internos
    return s

def validate_len(s: str, max_len: int = 30) -> str | None:
    if len(s) > max_len:
        return f"Supera {max_len} caracteres (tiene {len(s)})"
    return None

# ======================================================
# UI
# ======================================================
uploaded = st.file_uploader("📤 Sube archivo (xlsx/xls/csv/txt)", type=["xlsx", "xls", "csv", "txt"])
if not uploaded:
    st.stop()

try:
    df = read_any_table(uploaded.name, uploaded.getvalue())
except Exception as e:
    st.error(f"No pude leer el archivo: {e}")
    st.stop()

st.success(f"Archivo cargado. Filas: {len(df):,} | Columnas: {df.shape[1]:,}")
st.subheader("Vista previa (original)")
st.dataframe(df.head(20), use_container_width=True)

cols = list(df.columns)

st.divider()
st.subheader("✅ Configuración rápida")

c1, c2, c3 = st.columns(3)
with c1:
    apt_col = st.selectbox("Columna Apartamento (Limpia caracteres especiales)", ["(Ninguna)"] + cols, index=0)
with c2:
    email_col = st.selectbox("Columna Correo)", ["(Ninguna)"] + cols, index=0)
with c3:
    max30_cols = st.multiselect("Columnas con máximo 30 caracteres", cols, default=[])

st.caption("NOMPROPIO siempre activo. Truncamiento a 30 siempre activo (solo en columnas seleccionadas) y se reporta en ERRORES.")

st.divider()

if st.button("🚀 Limpiar y generar archivos"):
    numeric_cols = detect_numeric_columns(df)
    country_code_cols = detect_country_code_columns(df)

    df_clean = df.copy()
    error_rows = []

    for idx, row in df.iterrows():
        row_errs = []

        for c in df_clean.columns:
            raw_val = row.get(c)

            # Indicativo país (prioridad)
            if c in country_code_cols:
                df_clean.at[idx, c] = clean_country_code(raw_val)
                continue

            # Apartamento seleccionado (Regla 4)
            if apt_col != "(Ninguna)" and c == apt_col:
                df_clean.at[idx, c] = clean_apto_keep_inner_spaces(raw_val)
                continue

            # Correo seleccionado (Regla 5)
            if email_col != "(Ninguna)" and c == email_col:
                email, e = clean_email_one(raw_val)
                df_clean.at[idx, c] = email
                if e:
                    row_errs.append(f"{c}: {e}")
                continue

            # Numéricas: conservar decimales con '.'
            if c in numeric_cols:
                df_clean.at[idx, c] = clean_numeric_general(raw_val)
                continue

            # Texto: Reglas 1-3
            df_clean.at[idx, c] = clean_text_general(raw_val)

        # Regla 6: máximo 30 (TRUNCA y reporta)
        for c in max30_cols:
            s = "" if pd.isna(df_clean.at[idx, c]) else str(df_clean.at[idx, c])
            len_err = validate_len(s, 30)
            if len_err:
                original_len = len(s)
                df_clean.at[idx, c] = s[:30]
                row_errs.append(f"{c}: {len_err} -> TRUNCADO a 30")

        if row_errs:
            # fila_origen excel: header fila 1, datos desde fila 2
            error_rows.append({"fila_origen": idx + 2, "errores": " | ".join(row_errs)})

    err_df = pd.DataFrame(error_rows)

    st.success("✅ Limpieza terminada.")
    st.subheader("Vista previa (limpio)")
    st.dataframe(df_clean.head(20), use_container_width=True)

    st.subheader("Errores / Advertencias")
    st.write(f"Total filas con notas (errores o truncamientos): {len(err_df)}")
    st.dataframe(err_df.head(100), use_container_width=True)

    out_xlsx = io.BytesIO()
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="LIMPIO")
        err_df.to_excel(writer, index=False, sheet_name="ERRORES")
    out_xlsx.seek(0)

    out_csv = df_clean.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    base = os.path.splitext(uploaded.name)[0]
    st.download_button(
        "⬇️ Descargar Excel (LIMPIO + ERRORES)",
        data=out_xlsx,
        file_name=f"{base}_LIMPIO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "⬇️ Descargar CSV limpio",
        data=out_csv,
        file_name=f"{base}_LIMPIO.csv",
        mime="text/csv",
    )


