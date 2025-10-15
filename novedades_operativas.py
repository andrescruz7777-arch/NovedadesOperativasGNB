# ===============================
# ⚖️ Dependiente Judicial de Novedades – Streamlit
# ===============================
import os
import io
import re
import json
import base64
import datetime as dt
from typing import List, Dict, Any, Optional

import streamlit as st
import pandas as pd

# Intentaremos importar extract_msg; si no está, hacemos fallback.
try:
    import extract_msg  # para .msg de Outlook
    HAS_MSG = True
except Exception:
    HAS_MSG = False

from email import policy
from email.parser import BytesParser

# ======================
# Configuración Streamlit
# ======================
st.set_page_config(page_title="⚖️ Dependiente Judicial – Novedades", layout="wide")
st.title("⚖️ Dependiente Judicial – Novedades Operativas")
st.caption("Carga correos .msg/.eml, la IA clasifica, sugiere acción y redacta respuesta. Persistencia todo el día + guardado incremental en Excel.")

# ======================
# OpenAI Client
# ======================
OAI_READY = False
try:
    from openai import OpenAI
    OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY"))
    if OPENAI_API_KEY:
        os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY
        oai_client = OpenAI()
        OAI_READY = True
    else:
        st.warning("⚠️ Falta OPENAI_API_KEY en st.secrets o variables de entorno. El análisis IA devolverá 'VALIDAR MANUALMENTE'.")
except Exception as e:
    st.warning(f"⚠️ OpenAI SDK no disponible ({e}). El análisis IA devolverá 'VALIDAR MANUALMENTE'.")

# ======================
# Persistencia de sesión (no reinicia en el día)
# ======================
if "novedades" not in st.session_state:
    st.session_state["novedades"] = []  # lista de dicts
if "categorias" not in st.session_state:
    st.session_state["categorias"] = set()  # para autoaprendizaje
if "ultimo_guardado" not in st.session_state:
    st.session_state["ultimo_guardado"] = None

BITACORA_FILENAME = "bitacora_novedades.xlsx"

# ======================
# Utilidades
# ======================
def parse_eml(file_bytes: bytes) -> Dict[str, Any]:
    """Parsea .eml a dict: from, to, subject, date, body."""
    try:
        msg = BytesParser(policy=policy.default).parsebytes(file_bytes)
        subject = msg.get("subject", "") or ""
        sender = msg.get("from", "") or ""
        to = msg.get("to", "") or ""
        date = msg.get("date", "") or ""
        # Obtener cuerpo preferiblemente en texto
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/plain":
                    body += part.get_content() or ""
        else:
            body = msg.get_content() or ""
        return {
            "from": sender,
            "to": to,
            "subject": subject,
            "date": date,
            "body": body.strip(),
        }
    except Exception as e:
        return {"from": "", "to": "", "subject": "", "date": "", "body": f"ERROR parse_eml: {e}"}

def parse_msg(file_bytes: bytes) -> Dict[str, Any]:
    """Parsea .msg de Outlook a dict. Requiere extract_msg."""
    if not HAS_MSG:
        return {"from": "", "to": "", "subject": "", "date": "", "body": "extract_msg no disponible. Suba .eml o instale extract_msg.", "parse_error": True}
    try:
        # extract_msg espera ruta o stream-like con nombre; usamos NamedTemporaryFile
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".msg", delete=True) as tmp:
            tmp.write(file_bytes)
            tmp.flush()
            msg = extract_msg.Message(tmp.name)
            msg_sender = msg.sender or msg.sender_email or ""
            msg_to = msg.to or ""
            subject = msg.subject or ""
            date = msg.date or ""
            body = (msg.body or "").strip()
            return {
                "from": msg_sender,
                "to": msg_to,
                "subject": subject,
                "date": date,
                "body": body,
            }
    except Exception as e:
        return {"from": "", "to": "", "subject": "", "date": "", "body": f"ERROR parse_msg: {e}", "parse_error": True}

def clean_text(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r'\r', '\n', s)
    s = re.sub(r'\n+', '\n', s)
    s = s.strip()
    return s

def safe_json_loads(s: str) -> Optional[dict]:
    try:
        return json.loads(s)
    except Exception:
        s2 = s.strip()
        if s2.startswith("```") and s2.endswith("```"):
            try:
                return json.loads("\n".join(s2.splitlines()[1:-1]))
            except Exception:
                return None
        return None

# ======================
# Prompt IA (clasificación + acción + respuesta)
# ======================
PROMPT_CLASIF = (
    "Eres un dependiente judicial operativo experto. Analizarás el contenido de un correo de NOVEDAD OPERATIVA "
    "en contexto de procesos judiciales (mandamiento de pago, autos, cargues en aplicativo del banco, radicados, embargos, etc.).\n\n"
    "OBJETIVO:\n"
    "- Clasificar la novedad en una categoría existente o crear una nueva si no encaja claramente.\n"
    "- Detectar subcategoría específica.\n"
    "- Asignar impacto operativo: Alto / Medio / Bajo.\n"
    "- Resumir la novedad en 2-4 líneas.\n"
    "- Indicar la acción operativa concreta que debe ejecutar el analista (pasos precisos).\n"
    "- Redactar un correo profesional y conciso para responder al banco.\n\n"
    "FORMATO DE SALIDA (JSON ESTRICTO, sin texto adicional):\n"
    "{\n"
    "  \"categoria\": \"...\",\n"
    "  \"subcategoria\": \"...\",\n"
    "  \"impacto\": \"Alto\" | \"Medio\" | \"Bajo\",\n"
    "  \"resumen\": \"...\",\n"
    "  \"accion_recomendada\": \"...\", \n"
    "  \"respuesta_sugerida\": \"Asunto: ...\\nCuerpo: ...\"\n"
    "}\n\n"
    "REGLAS IMPORTANTES:\n"
    "- Si observas que la novedad parece de tipo 'Auto o mandamiento no cargado/visible', 'Desfase procesal (rama vs banco)', "
    "'Error de identificación (radicado/cédula)', 'Falta de soporte/adjuntos', 'Falla de aplicativo', 'Reiteración/seguimiento', etc., usa esa categoría; "
    "si no encaja, crea una categoría nueva y sé claro en su nombre.\n"
    "- La 'accion_recomendada' deben ser pasos OBJETIVOS (ej: verificar portal Rama Judicial, descargar PDF, cargar a banco con nombre X, confirmar cargue, responder)."
)

def analyze_email_with_ai(subject: str, body: str, model: str = "gpt-4o-mini") -> Dict[str, Any]:
    """Clasifica y genera acción + respuesta. Si IA no disponible, retorna placeholders."""
    text = f"Asunto: {subject}\n\nCuerpo:\n{body}"
    if not OAI_READY:
        return {
            "categoria": "NO CLASIFICADO – VALIDAR MANUALMENTE",
            "subcategoria": "NO CLASIFICADO – VALIDAR MANUALMENTE",
            "impacto": "Medio",
            "resumen": "IA no disponible. Revisar manualmente el contenido del correo.",
            "accion_recomendada": "Validar en Rama Judicial y aplicativo del banco; generar respuesta estándar.",
            "respuesta_sugerida": "Asunto: Acuse de recibo – Novedad operativa\nCuerpo: Hemos recibido su observación. Procederemos a validar y cargar la documentación correspondiente. Cordial saludo."
        }

    try:
        from openai import OpenAI
        client = OpenAI()
        content_blocks = [
            {"type": "text", "text": PROMPT_CLASIF},
            {"type": "text", "text": text},
        ]
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": content_blocks}],
            temperature=0,
        )
        raw = resp.choices[0].message.content
        data = safe_json_loads(raw) or {}
        # Normalización mínima
        for k in ["categoria", "subcategoria", "impacto", "resumen", "accion_recomendada", "respuesta_sugerida"]:
            if k not in data or not str(data[k]).strip():
                data[k] = "NO SE APORTÓ – VALIDAR MANUALMENTE"
        return data
    except Exception as e:
        return {
            "categoria": "ERROR IA – VALIDAR MANUALMENTE",
            "subcategoria": str(e)[:180],
            "impacto": "Medio",
            "resumen": "Falla al procesar con IA.",
            "accion_recomendada": "Revisar manualmente y responder con plantilla estándar.",
            "respuesta_sugerida": "Asunto: Acuse de recibo – Novedad operativa\nCuerpo: Se revisará la novedad y se dará respuesta formal en el menor tiempo posible."
        }

# ======================
# Guardado incremental local
# ======================
def guardar_excel_incremental(rows: List[Dict[str, Any]], filename: str = BITACORA_FILENAME):
    """Guarda/actualiza un Excel local con acumulado. No borra lo anterior."""
    try:
        # Si existe, concatenamos; si no, creamos
        if os.path.exists(filename):
            prev = pd.read_excel(filename)
            df_new = pd.DataFrame(rows)
            df_full = pd.concat([prev, df_new], ignore_index=True)
        else:
            df_full = pd.DataFrame(rows)
        # Guardar
        df_full.to_excel(filename, index=False, engine="openpyxl")
        st.session_state["ultimo_guardado"] = dt.datetime.now()
    except Exception as e:
        st.error(f"Error guardando Excel: {e}")

# ======================
# UI – Carga y procesamiento
# ======================
st.subheader("📥 Cargar correos (.msg / .eml)")
uploads = st.file_uploader(
    "Arrastra o selecciona 1..N archivos",
    type=["msg", "eml"],
    accept_multiple_files=True
)

colA, colB, colC = st.columns([1,1,2])
with colA:
    modelo = st.selectbox("Modelo IA", ["gpt-4o-mini", "gpt-4o"], index=0)
with colB:
    btn_proc = st.button("🚀 Procesar cargados")
with colC:
    st.write("")

# Procesar
if btn_proc and uploads:
    progress = st.progress(0.0)
    nuevos_registros = []

    for i, up in enumerate(uploads, start=1):
        name = up.name
        ext = os.path.splitext(name)[1].lower()
        data = up.read()

        # Parse
        if ext == ".eml":
            parsed = parse_eml(data)
        else:  # .msg
            parsed = parse_msg(data)

        subject = clean_text(parsed.get("subject", "")) or name
        sender = clean_text(parsed.get("from", ""))
        date_raw = clean_text(parsed.get("date", ""))
        body = clean_text(parsed.get("body", ""))

        # Radicado probable (heurística: 10+ dígitos o patrón usual)
        rad_match = re.search(r'\b\d{15,23}\b', subject + " " + body)  # ajusta si tu patrón es distinto
        radicado = rad_match.group(0) if rad_match else ""

        # IA clasificación + acción + respuesta
        result = analyze_email_with_ai(subject, body, model=modelo)
        categoria = result.get("categoria", "")
        subcat = result.get("subcategoria", "")
        impacto = result.get("impacto", "")
        resumen = result.get("resumen", "")
        accion = result.get("accion_recomendada", "")
        respuesta = result.get("respuesta_sugerida", "")

        # Autoaprendizaje categorías
        if categoria:
            st.session_state["categorias"].add(categoria)

        registro = {
            "Fecha_Registro": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Archivo": name,
            "Remitente": sender,
            "Fecha_Correo": date_raw,
            "Asunto": subject,
            "Radicado": radicado if radicado else "NO DETECTADO",
            "Categoria": categoria,
            "Subcategoria": subcat,
            "Impacto": impacto,
            "Resumen": resumen,
            "Accion_Recomendada": accion,
            "Respuesta_Sugerida": respuesta,
            "Estado": "Pendiente",
            "Observaciones": "",
        }

        st.session_state["novedades"].append(registro)
        nuevos_registros.append(registro)

        # Guardado incremental tras cada correo (persistencia durante el día)
        guardar_excel_incremental([registro])

        progress.progress(i / len(uploads))

    st.success(f"✔️ Procesados {len(nuevos_registros)} correos. Guardados incrementalmente en {BITACORA_FILENAME}.")

# ======================
# Panel de control y exportación
# ======================
st.subheader("📊 Novedades del día (acumulado en sesión)")
df = pd.DataFrame(st.session_state["novedades"]) if st.session_state["novedades"] else pd.DataFrame(columns=[
    "Fecha_Registro","Archivo","Remitente","Fecha_Correo","Asunto","Radicado","Categoria","Subcategoria","Impacto","Resumen","Accion_Recomendada","Respuesta_Sugerida","Estado","Observaciones"
])

st.dataframe(df, use_container_width=True, height=420)

col1, col2, col3 = st.columns([1,1,2])
with col1:
    if not df.empty:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Novedades", index=False)
        st.download_button(
            "📥 Descargar Excel (sesión)",
            data=out.getvalue(),
            file_name="novedades_sesion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
with col2:
    if os.path.exists(BITACORA_FILENAME):
        with open(BITACORA_FILENAME, "rb") as f:
            st.download_button(
                "📦 Descargar Bitácora Acumulada",
                data=f.read(),
                file_name=BITACORA_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
with col3:
    if st.session_state["ultimo_guardado"]:
        st.caption(f"🕒 Último guardado: {st.session_state['ultimo_guardado'].strftime('%Y-%m-%d %H:%M:%S')}")

st.markdown("---")

# ======================
# Cierre del día
# ======================
st.subheader("🧾 Cerrar y Consolidar Día")
st.caption("Guarda definitivamente y limpia la sesión para iniciar un nuevo día de trabajo (la bitácora local NO se borra).")

colx, coly = st.columns([1,3])
with colx:
    if st.button("✅ Cerrar y consolidar"):
        # Guardar lo que haya en sesión (por si no se guardó algo)
        if st.session_state["novedades"]:
            guardar_excel_incremental(st.session_state["novedades"])
        st.session_state["novedades"] = []
        st.success("Día consolidado. La sesión se ha limpiado. La bitácora acumulada permanece en el archivo local.")

with coly:
    st.write("")
    if st.session_state["categorias"]:
        st.info(f"📚 Categorías detectadas/creadas hoy: {', '.join(sorted(st.session_state['categorias']))}")
