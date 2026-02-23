import io
import os
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Transformador Talk", page_icon="🧼", layout="wide")
st.title("🧼 Transformador + Validador para Talk")
st.caption("Sube cualquier archivo (Excel/CSV/TXT), mapea columnas, aplicamos reglas Talk y descargas el Excel final + reporte de errores.")

# ======================================================
# Utilidades: lectura flexible
# ======================================================
def read_any_table(file_name: str, file_bytes: bytes, header_row: int | None) -> pd.DataFrame:
    ext = os.path.splitext(file_name.lower())[1]

    if ext in [".xlsx", ".xls"]:
        # header_row: None => no header, devuelve filas crudas
        return pd.read_excel(io.BytesIO(file_bytes), header=header_row, dtype=object)

    if ext in [".csv", ".txt"]:
        # intentos de separador
        for sep in [",", ";", "\t", "|"]:
            try:
                return pd.read_csv(io.BytesIO(file_bytes), sep=sep, header=header_row, dtype=object, engine="python")
            except Exception:
                pass
        # fallback raw
        text = file_bytes.decode("utf-8", errors="replace").splitlines()
        return pd.DataFrame({"raw": text})

    raise ValueError(f"Extensión no soportada: {ext}")

def guess_header_row(df_raw: pd.DataFrame, max_rows: int = 30) -> int:
    """
    Heurística simple: busca la fila con más celdas tipo string útiles.
    """
    best_i, best_score = 0, -1
    top = df_raw.head(max_rows)

    for i in range(len(top)):
        row = top.iloc[i]
        # score: cantidad de strings no vacíos
        score = 0
        for v in row.tolist():
            if pd.isna(v):
                continue
            s = str(v).strip()
            if len(s) >= 2:
                score += 1
        if score > best_score:
            best_score = score
            best_i = i
    return best_i

# ======================================================
# Limpieza según reglas Talk
# ======================================================
DISALLOWED_SIGNS = r",\.\-\$\%\#\(\)\/"  # signos a remover

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
    # Similar a NOMPROPIO: title() pero cuidando espacios múltiples
    # (ya normalizamos antes). title() en Python convierte "de"->"De", igual que Excel.
    return s.lower().title()

def remove_specials_talk(s: str) -> str:
    # Quita tildes + signos específicos + cualquier cosa rara dejando letras/números/espacios
    s = strip_accents(s)
    s = re.sub(f"[{DISALLOWED_SIGNS}]", " ", s)      # quita signos listados
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)            # quita otros especiales
    s = normalize_spaces(s)
    return s.strip()

def clean_text_talk(val: object, proper: bool = True, max_len: int | None = None) -> str:
    if pd.isna(val):
        return ""
    s = str(val)
    s = remove_invisibles(s)
    s = s.strip()
    s = normalize_spaces(s)
    s = remove_specials_talk(s)
    if proper:
        s = to_proper_case_like_excel(s)
    if max_len is not None and len(s) > max_len:
        # NO truncamos aquí; lo dejamos para validación/decisión
        pass
    return s

def clean_apt_number(val: object) -> str:
    # Regla 4: sin caracteres especiales => dejamos solo letras y números (sin espacios)
    if pd.isna(val):
        return ""
    s = str(val)
    s = remove_invisibles(s).strip()
    s = strip_accents(s)
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s

def extract_single_email(val: object) -> tuple[str, str | None]:
    """
    Devuelve (email_limpio, error_o_None)
    - si hay más de un email o separadores, marca error
    """
    if pd.isna(val):
        return ("", "Correo vacío")
    raw = str(val).strip()

    # si contiene separadores típicos de "varios correos"
    if any(sep in raw for sep in [";", ",", " ", "/", "|"]):
        # puede ser un email con espacios; intentamos compactar
        compact = raw.replace(" ", "")
        if any(sep in compact for sep in [";", ",", "/", "|"]):
            return ("", "Más de un correo o separadores detectados")
        raw = compact

    email = raw.strip()
    if not EMAIL_RE.match(email):
        return ("", "Correo inválido")
    return (email.lower(), None)

def validate_len(s: str, max_len: int) -> str | None:
    if len(s) > max_len:
        return f"Supera {max_len} caracteres (tiene {len(s)})"
    return None

# ======================================================
# UI: subir + detectar header + mapping
# ======================================================
uploaded = st.file_uploader("📤 Sube archivo crudo (xlsx/xls/csv/txt)", type=["xlsx", "xls", "csv", "txt"])

if not uploaded:
    st.stop()

file_bytes = uploaded.getvalue()

# Leemos sin header para adivinar dónde está
df_no_header = read_any_table(uploaded.name, file_bytes, header_row=None)
guess = guess_header_row(df_no_header)

col1, col2, col3 = st.columns([1,1,2])
with col1:
    header_row = st.number_input("Fila de encabezado (0-index)", min_value=0, max_value=200, value=int(guess))
with col2:
    preview_rows = st.number_input("Filas preview", min_value=5, max_value=50, value=15)
with col3:
    st.info("Si tu archivo tiene títulos arriba (como tu ejemplo), ajusta la fila de encabezado hasta que veas nombres de columnas correctos.")

df = read_any_table(uploaded.name, file_bytes, header_row=int(header_row))

st.subheader("Vista previa del archivo (ya con encabezado)")
st.dataframe(df.head(int(preview_rows)), use_container_width=True)

# Columnas disponibles
cols = list(df.columns)

st.divider()
st.subheader("Mapeo de columnas hacia plantilla Talk")

# Campos Talk (en orden)
TALK_FIELDS = [
    "Tipo de propiedad  *Obligatorio",
    "Nombre de la propiedad  *Obligatorio",
    "Nombre Propietario   *Obligatorio",
    "Apellido Propietario  *Obligatorio",
    "Correo Propietario *Obligatorio",
    "Complemento de la propiedad Opcional ",
    "Código País Teléfono Propietario",
    "Teléfono Propietario",
    "Coeficiente #1  *Obligatorio",
    "Coeficiente #2 Opcional ",
    "Coeficiente #3 Opcional ",
    "Estado del pago *Obligatorio",
]

DEFAULTS = {
    "Código País Teléfono Propietario": "",
    "Coeficiente #2 Opcional ": "",
    "Coeficiente #3 Opcional ": "",
}

st.caption("Selecciona de qué columna viene cada campo Talk. Si un campo no existe, puedes dejarlo en '(Valor fijo)' y escribir un valor.")

def select_source(field: str):
    options = ["(Vacío)", "(Valor fijo)"] + cols
    return st.selectbox(field, options, index=0)

mapping = {}
fixed_values = {}
left, right = st.columns(2)

for i, f in enumerate(TALK_FIELDS):
    target = left if i % 2 == 0 else right
    with target:
        sel = select_source(f)
        mapping[f] = sel
        if sel == "(Valor fijo)":
            fixed_values[f] = st.text_input(f"Valor fijo para: {f}", value=DEFAULTS.get(f, ""))
        elif sel == "(Vacío)":
            fixed_values[f] = DEFAULTS.get(f, "")
        else:
            fixed_values[f] = None

st.divider()
truncate_long = st.checkbox("✂️ Truncar automáticamente campos de 30 caracteres (Nombre/Apellido/Complemento) en vez de solo reportar error", value=False)

if st.button("🚀 Generar archivo Talk + reporte de errores"):
    out = pd.DataFrame(columns=TALK_FIELDS)
    errors = []

    # Construcción + limpieza por reglas
    for idx, row in df.iterrows():
        rec = {}
        row_errors = []

        for field in TALK_FIELDS:
            src = mapping[field]
            if src in ["(Vacío)", "(Valor fijo)"]:
                raw_val = fixed_values.get(field, "")
            else:
                raw_val = row.get(src)

            # Limpiezas por tipo de campo
            if field == "Correo Propietario *Obligatorio":
                email, e = extract_single_email(raw_val)
                rec[field] = email
                if e:
                    row_errors.append(f"{field}: {e}")

            elif field == "Complemento de la propiedad Opcional ":
                # regla 6 aplica aquí (max 30) + regla 3
                txt = clean_text_talk(raw_val, proper=True, max_len=30)
                len_err = validate_len(txt, 30)
                if len_err:
                    if truncate_long:
                        txt = txt[:30]
                    row_errors.append(f"{field}: {len_err}")
                rec[field] = txt

            elif field in ["Nombre Propietario   *Obligatorio", "Apellido Propietario  *Obligatorio"]:
                txt = clean_text_talk(raw_val, proper=True, max_len=30)
                len_err = validate_len(txt, 30)
                if len_err:
                    if truncate_long:
                        txt = txt[:30]
                    row_errors.append(f"{field}: {len_err}")
                rec[field] = txt

            elif field == "Nombre de la propiedad  *Obligatorio":
                # Regla 1,2,3. (No limit aquí)
                rec[field] = clean_text_talk(raw_val, proper=True)

            elif field == "Tipo de propiedad  *Obligatorio":
                rec[field] = clean_text_talk(raw_val, proper=True)

            elif field == "Coeficiente #1  *Obligatorio":
                # Lo dejamos como viene, pero trim; Talk suele aceptar número
                val = "" if pd.isna(raw_val) else str(raw_val).strip()
                rec[field] = val
                if val == "":
                    row_errors.append(f"{field}: vacío")

            elif field == "Estado del pago *Obligatorio":
                val = clean_text_talk(raw_val, proper=True)
                rec[field] = val
                if val == "":
                    row_errors.append(f"{field}: vacío")

            elif field == "Teléfono Propietario":
                # Solo trim y normaliza espacios
                val = "" if pd.isna(raw_val) else normalize_spaces(str(raw_val).strip())
                rec[field] = val

            elif field == "Código País Teléfono Propietario":
                val = "" if pd.isna(raw_val) else str(raw_val).strip()
                rec[field] = val

            elif field in ["Coeficiente #2 Opcional ", "Coeficiente #3 Opcional "]:
                rec[field] = "" if pd.isna(raw_val) else str(raw_val).strip()

            else:
                rec[field] = clean_text_talk(raw_val, proper=True)

        # Regla 4 (apto sin especiales) — asumiendo que tu "Número apto" va en "Complemento"
        # Si tu apto realmente es otra columna, mapéala en Complemento o en Nombre de la propiedad según tu proceso.
        # Aquí solo aplicamos a Complemento como “Apto” típico.
        # Si NO quieres esto, desactívalo y crea un campo específico.
        if rec["Complemento de la propiedad Opcional "]:
            rec["Complemento de la propiedad Opcional "] = clean_apt_number(rec["Complemento de la propiedad Opcional "])

        if row_errors:
            errors.append({
                "fila_origen": idx,
                "errores": " | ".join(row_errors),
            })

        out.loc[len(out)] = rec

    st.success("Generación terminada.")
    st.subheader("Preview archivo Talk")
    st.dataframe(out.head(20), use_container_width=True)

    err_df = pd.DataFrame(errors)
    st.subheader("Errores encontrados")
    st.write(f"Total filas con error: {len(err_df)}")
    st.dataframe(err_df.head(50), use_container_width=True)

    # Exportar Excel con 2 hojas: Talk + Errores
    out_xlsx = io.BytesIO()
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Talk")
        err_df.to_excel(writer, index=False, sheet_name="Errores")
    out_xlsx.seek(0)

    base = os.path.splitext(uploaded.name)[0]
    st.download_button(
        "⬇️ Descargar Excel final (Talk + Errores)",
        data=out_xlsx,
        file_name=f"{base}_TALK_GENERADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )