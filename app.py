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
    "Limpieza inteligente: detecta columnas numéricas, indicativos país y aplica reglas Talk."
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
DISALLOWED_SIGNS = r",\.\-\$\%\#\(\)\/"  # para texto (Regla 3) - ojo: en numérico NO quitamos '.'
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
    s = str(v)
    s = remove_invisibles(s).strip()
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
        good = 0
        total = 0
        for v in sample:
            v = v.strip()
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
    s = str(v)
    s = remove_invisibles(s).strip()
    if s == "":
        return False
    # +507, 507, +57, 57 etc, cortos
    if re.fullmatch(r"\+?\d{1,4}", s):
        return True
    # a veces viene "+507 " o " +507"
    s2 = re.sub(r"\s+", "", s)
    return bool(re.fullmatch(r"\+?\d{1,4}", s2))

def detect_country_code_columns(df: pd.DataFrame, numeric_cols: set, sample_size: int = 80, threshold: float = 0.75) -> set:
    cc_cols = set()
    for c in df.columns:
        name = str(c).strip().lower()

        # Heurística por nombre
        name_hit = any(k in name for k in [
            "indicativo", "codigo pais", "código pais", "pais", "country code", "código país"
        ])

        series = df[c].dropna().astype(str)
        if series.empty:
            continue

        sample = series.head(sample_size)
        total = 0
        good = 0
        for v in sample:
            v = v.strip()
            if v == "":
                continue
            total += 1
            if looks_like_country_code_value(v):
                good += 1

        value_hit = (total > 0 and (good / total) >= threshold)

        # Importante: si la col es numérica no la excluimos, porque indicativo también se ve "numérico".
        # Preferimos marcarla como country_code si cumple.
        if name_hit or value_hit:
            cc_cols.add(c)

    return cc_cols

# ======================================================
# Limpieza por tipo
# ======================================================
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def clean_email_one(val: object) -> tuple[str, str | None]:
    if pd.isna(val) or str(val).strip() == "":
        return ("", "Correo vacío")
    raw = remove_invisibles(str(val)).strip()

    # si hay separadores típicos de múltiples correos
    if any(sep in raw for sep in [";", ",", "/", "|"]):
        return ("", "Más de un correo o separadores detectados")

    email = raw.replace(" ", "")
    if not EMAIL_RE.match(email):
        return ("", "Correo inválido")
    return (email.lower(), None)

def clean_text_general(val: object, apply_nompropio: bool = True) -> str:
    """
    Reglas 1-3 para TEXTO:
    - trim
    - sin tildes
    - sin caracteres especiales (incluye , . - $ % # ( / ))
    - mantiene espacios (colapsa múltiple a 1)
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val))
    s = s.strip()
    if s == "":
        return ""
    s = strip_accents(s)
    # reemplaza signos prohibidos por espacio
    s = re.sub(f"[{DISALLOWED_SIGNS}]", " ", s)
    # quita otros especiales, deja letras/números/espacios
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)
    s = collapse_spaces(s).strip()
    if apply_nompropio and s:
        s = nompropio_like_excel(s)
    return s

def clean_numeric_general(val: object) -> str:
    """
    Limpieza para columnas NUMÉRICAS:
    - trim
    - quita espacios internos
    - convierte coma a punto para decimal
    - elimina caracteres basura (como '*', '$', '%', etc)
    - conserva '.' como decimal
    Ej: ' 3 ,45* ' -> '3.45'
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val)).strip()
    if s == "":
        return ""
    # quita espacios internos
    s = re.sub(r"\s+", "", s)
    # convierte coma a punto
    s = s.replace(",", ".")
    # deja solo dígitos, punto, signo +/-
    s = re.sub(r"[^0-9\.\+\-]", "", s)

    # si quedan múltiples puntos, intentamos quedarnos con el primero como decimal y borrar los demás
    if s.count(".") > 1:
        parts = s.split(".")
        s = parts[0] + "." + "".join(parts[1:])

    # casos extremos: solo '+' o '-' o '.'
    if s in ["+", "-", ".", "+.", "-."]:
        return ""
    return s

def clean_country_code(val: object) -> str:
    """
    Indicativo país:
    - conserva '+' si venía
    - deja solo + y dígitos
    Ej '+507' -> '+507', '507' -> '507'
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
    - limpieza general de especiales, pero:
      * NO colapsa ni elimina espacios internos
      * solo quita espacios al inicio/fin
    - quita tildes
    - elimina caracteres especiales (dejando letras/números/espacios)
    """
    if pd.isna(val):
        return ""
    s = remove_invisibles(str(val))
    s = s.strip()  # quita bordes, mantiene internos
    if s == "":
        return ""
    s = strip_accents(s)
    # elimina signos prohibidos SIN convertir a un solo espacio (ponemos "")
    s = re.sub(f"[{DISALLOWED_SIGNS}]", "", s)
    # elimina otros especiales, dejando letras/números/espacios
    s = re.sub(r"[^A-Za-z0-9\s]", "", s)
    # NO tocamos espacios internos
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
    apt_col = st.selectbox("Columna Apartamento (Regla 4)", ["(Ninguna)"] + cols, index=0)
with c2:
    email_col = st.selectbox("Columna Correo (Regla 5)", ["(Ninguna)"] + cols, index=0)
with c3:
    max30_cols = st.multiselect("Columnas con máximo 30 caracteres (Regla 6)", cols, default=[])

apply_proper = st.checkbox("Aplicar NOMPROPIO / Capitalización (Regla 1)", value=True)
truncate_long = st.checkbox("✂️ Truncar automáticamente si pasa 30 (si no, solo reporta error)", value=False)

st.caption(
    "La app detecta automáticamente columnas numéricas (para no dañar decimales) y columnas de indicativo país (para conservar '+')."
)

st.divider()

if st.button("🚀 Limpiar y generar archivos"):
    # Detectar tipos de columnas una sola vez
    numeric_cols = detect_numeric_columns(df)
    country_code_cols = detect_country_code_columns(df, numeric_cols)

    # Nota: si una col está en country_code_cols, se trata como indicativo (prioridad)
    df_clean = df.copy()
    error_rows = []

    for idx, row in df.iterrows():
        row_errs = []

        for c in df_clean.columns:
            raw_val = row.get(c)

            # Prioridad: indicativo país
            if c in country_code_cols:
                df_clean.at[idx, c] = clean_country_code(raw_val)
                continue

            # Columna apartamento seleccionada (regla 4)
            if apt_col != "(Ninguna)" and c == apt_col:
                df_clean.at[idx, c] = clean_apto_keep_inner_spaces(raw_val)
                continue

            # Columna email seleccionada (regla 5)
            if email_col != "(Ninguna)" and c == email_col:
                email, e = clean_email_one(raw_val)
                df_clean.at[idx, c] = email
                if e:
                    row_errs.append(f"{c}: {e}")
                continue

            # Numéricas: mantener decimales y limpiar basura
            if c in numeric_cols:
                df_clean.at[idx, c] = clean_numeric_general(raw_val)
                continue

            # Texto general: reglas 1-3
            df_clean.at[idx, c] = clean_text_general(raw_val, apply_nompropio=apply_proper)

        # Regla 6: máximo 30 en columnas seleccionadas
        for c in max30_cols:
            s = "" if pd.isna(df_clean.at[idx, c]) else str(df_clean.at[idx, c])
            len_err = validate_len(s, 30)
            if len_err:
                if truncate_long:
                    df_clean.at[idx, c] = s[:30]
                row_errs.append(f"{c}: {len_err}")

        if row_errs:
            # fila_origen en Excel: header es fila 1, datos inician en fila 2
            error_rows.append({"fila_origen": idx + 2, "errores": " | ".join(row_errs)})

    err_df = pd.DataFrame(error_rows)

    st.success("✅ Limpieza terminada.")
    st.subheader("Vista previa (limpio)")
    st.dataframe(df_clean.head(20), use_container_width=True)

    st.subheader("Errores encontrados")
    st.write(f"Total filas con error: {len(err_df)}")
    st.dataframe(err_df.head(50), use_container_width=True)

    # Export Excel con 2 hojas: LIMPIO + ERRORES
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
