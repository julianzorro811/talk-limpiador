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
    "Aplica reglas Talk y genera archivo limpio + reporte de errores."
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
                # si quedó en 1 columna, probablemente no era el separador correcto
                if df.shape[1] == 1 and sep != "|":
                    continue
                return df
            except Exception:
                pass

        # fallback: una sola columna raw
        text = file_bytes.decode("utf-8", errors="replace").splitlines()
        return pd.DataFrame({"raw": text})

    raise ValueError(f"Extensión no soportada: {ext}")

# ======================================================
# Limpieza según reglas Talk (1-3) + validaciones (4-6)
# ======================================================
DISALLOWED_SIGNS = r",\.\-\$\%\#\(\)\/"
INVISIBLE_CHARS_PATTERN = re.compile(r"[\uFEFF\u200B\u200C\u200D\u2060\u00AD]")
MULTISPACE_RE = re.compile(r"\s{2,}")
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

def remove_invisibles(s: str) -> str:
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    return INVISIBLE_CHARS_PATTERN.sub("", s)

def strip_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

def normalize_spaces(s: str) -> str:
    s = re.sub(r"[\t\r\n]+", " ", s)
    s = MULTISPACE_RE.sub(" ", s)
    return s

def to_proper_case_like_excel(s: str) -> str:
    # Similar a NOMPROPIO
    return s.lower().title()

def clean_text_talk(val: object, proper: bool = True) -> str:
    """
    Reglas 1-3:
    1) Nompropio (opcional)
    2) Sin espacios al inicio/fin
    3) Sin caracteres especiales, sin tildes, sin comas/puntos/-$%#( / )
    """
    if pd.isna(val):
        return ""
    s = str(val)
    s = remove_invisibles(s).strip()
    s = normalize_spaces(s)
    s = strip_accents(s)                        # sin tildes
    s = re.sub(f"[{DISALLOWED_SIGNS}]", " ", s)  # quita signos listados
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)        # quita otros especiales
    s = normalize_spaces(s).strip()
    if proper and s:
        s = to_proper_case_like_excel(s)
    return s

def clean_apt_number(val: object) -> str:
    """
    Regla 4: Apartamento sin caracteres especiales.
    Deja solo letras/números, sin espacios.
    """
    if pd.isna(val):
        return ""
    s = str(val)
    s = remove_invisibles(s).strip()
    s = strip_accents(s)
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s

def extract_single_email(val: object) -> tuple[str, str | None]:
    """
    Regla 5: solo 1 correo por fila y válido.
    """
    if pd.isna(val) or str(val).strip() == "":
        return ("", "Correo vacío")

    raw = str(val).strip()

    # si hay separadores típicos de "más de un correo"
    if any(sep in raw for sep in [";", ",", "/", "|"]):
        return ("", "Más de un correo o separadores detectados")

    email = raw.replace(" ", "")  # quita espacios internos
    if not EMAIL_RE.match(email):
        return ("", "Correo inválido")

    return (email.lower(), None)

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
st.subheader("✅ Configuración rápida (sin mapeo por campo)")

c1, c2, c3 = st.columns(3)
with c1:
    apt_col = st.selectbox("Columna de Apartamento (Regla 4)", ["(Ninguna)"] + cols, index=0)
with c2:
    email_col = st.selectbox("Columna de Correo (Regla 5)", ["(Ninguna)"] + cols, index=0)
with c3:
    max30_cols = st.multiselect("Columnas con máximo 30 chars (Regla 6)", cols, default=[])

apply_proper = st.checkbox("Aplicar NOMPROPIO / Capitalización (Regla 1)", value=True)
truncate_long = st.checkbox("✂️ Truncar automáticamente si pasa 30 (si no, solo reporta error)", value=False)

with st.expander("📌 Notas", expanded=False):
    st.write(
        "- Regla 2 (trim) y Regla 3 (sin tildes/especiales) se aplican a todas las columnas como texto.\n"
        "- Regla 4 se aplica solo a la columna que selecciones como Apartamento.\n"
        "- Regla 5 se aplica solo a la columna que selecciones como Correo.\n"
        "- Regla 6 se aplica a las columnas que marques (Nombre/Apellido/Complemento, etc.).\n"
        "- Si tu archivo tiene números que no quieres convertir, igual se guardan como texto (recomendado para cargas)."
    )

st.divider()

if st.button("🚀 Limpiar y generar archivos"):
    df_clean = df.copy()
    error_rows = []

    for idx, row in df.iterrows():
        row_errs = []

        # Reglas 1-3: limpiar texto en todas las columnas
        for c in df_clean.columns:
            df_clean.at[idx, c] = clean_text_talk(row.get(c), proper=apply_proper)

        # Regla 4: apartamento sin especiales
        if apt_col != "(Ninguna)":
            df_clean.at[idx, apt_col] = clean_apt_number(row.get(apt_col))

        # Regla 5: correo único y válido
        if email_col != "(Ninguna)":
            email, e = extract_single_email(row.get(email_col))
            df_clean.at[idx, email_col] = email
            if e:
                row_errs.append(f"{email_col}: {e}")

        # Regla 6: máximo 30 caracteres
        for c in max30_cols:
            s = "" if pd.isna(df_clean.at[idx, c]) else str(df_clean.at[idx, c])
            len_err = validate_len(s, 30)
            if len_err:
                if truncate_long:
                    df_clean.at[idx, c] = s[:30]
                row_errs.append(f"{c}: {len_err}")

        if row_errs:
            error_rows.append({"fila_origen": idx + 2, "errores": " | ".join(row_errs)})
            # idx+2 porque excel: fila 1 header, datos empiezan fila 2

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

    # Export CSV limpio (opcional)
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
