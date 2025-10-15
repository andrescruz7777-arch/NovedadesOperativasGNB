# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: IA compatible (OpenAI SDK >= 1.40)
# =========================

import streamlit as st
import pandas as pd
import openai
import io
import os
import pdfplumber
from docx import Document
import extract_msg

# =========================
# ⚙️ CONFIGURACIÓN INICIAL
# =========================
st.set_page_config(page_title="📬 Novedades Operativas GNB", layout="wide")
st.title("🤖 Analizador de Novedades Operativas - Contacto Solutions ⚖️")

# =========================
# 🔐 CONFIGURACIÓN DE API
# =========================
try:
    client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    IA_DISPONIBLE = True
except Exception as e:
    IA_DISPONIBLE = False
    st.warning(f"⚠️ OpenAI SDK no disponible ({e}). El análisis IA devolverá 'VALIDAR MANUALMENTE'.")

# =========================
# 📁 VARIABLES GLOBALES
# =========================
if "novedades_data" not in st.session_state:
    st.session_state["novedades_data"] = []
if "procesando" not in st.session_state:
    st.session_state["procesando"] = False

# =========================
# 🎨 ESTILOS
# =========================
st.markdown("""
<style>
    body { color: #1B168C; background-color: #FFFFFF; }
    .stApp { background-color: #FFFFFF; }
    .block-container { padding-top: 1rem; }
    .uploadedFile { color: #1B168C; }
</style>
""", unsafe_allow_html=True)

# =========================
# 🧩 FUNCIONES AUXILIARES
# =========================
def leer_archivo_msg(archivo):
    msg = extract_msg.Message(archivo)
    return f"De: {msg.sender}\nAsunto: {msg.subject}\n\n{msg.body}"

def leer_archivo_pdf(archivo):
    texto = ""
    with pdfplumber.open(archivo) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"
    return texto.strip()

def leer_archivo_docx(archivo):
    doc = Document(archivo)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def analizar_novedad(texto):
    if not IA_DISPONIBLE:
        return {
            "categoria": "VALIDAR MANUALMENTE",
            "accion_recomendada": "Revisar el contenido manualmente. No se pudo usar la IA.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE"
        }

    prompt = f"""
Analiza el siguiente correo o documento de novedad operativa del banco y clasifícalo según su naturaleza.

Texto:
{texto}

Responde en formato JSON con las siguientes claves:
- categoria: tipo general del requerimiento (ej: 'Mandamiento no visible', 'Auto pendiente de carga', 'Error en documento', 'Revisión de medidas', etc.)
- accion_recomendada: qué debe hacer el usuario para subsanar o resolver la novedad
- respuesta_sugerida: texto redactado para responder al correo del banco profesionalmente
"""
    try:
        respuesta = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "system", "content": "Eres un abogado experto en operaciones judiciales bancarias."},
                      {"role": "user", "content": prompt}],
            temperature=0.3
        )
        contenido = respuesta.choices[0].message.content
        return eval(contenido) if "{" in contenido else {
            "categoria": "ERROR DE FORMATO",
            "accion_recomendada": "La IA no devolvió un JSON válido.",
            "respuesta_sugerida": contenido
        }
    except Exception as e:
        return {
            "categoria": "ERROR DE PROCESAMIENTO",
            "accion_recomendada": f"Validar manualmente. Error: {e}",
            "respuesta_sugerida": "VALIDAR MANUALMENTE"
        }

# =========================
# 🧾 INTERFAZ DE CARGA
# =========================
st.subheader("📂 Cargar correos o documentos (.msg, .pdf, .docx)")
archivos = st.file_uploader("Selecciona uno o varios archivos", type=["msg", "pdf", "docx"], accept_multiple_files=True)

if archivos:
    if st.button("🚀 Analizar Novedades"):
        st.session_state.procesando = True
        resultados = []

        for archivo in archivos:
            nombre = archivo.name
            extension = nombre.split(".")[-1].lower()

            try:
                if extension == "msg":
                    texto = leer_archivo_msg(archivo)
                elif extension == "pdf":
                    texto = leer_archivo_pdf(archivo)
                elif extension == "docx":
                    texto = leer_archivo_docx(archivo)
                else:
                    texto = ""

                analisis = analizar_novedad(texto)
                resultados.append({
                    "ARCHIVO": nombre,
                    "CATEGORIA": analisis.get("categoria", ""),
                    "ACCION_RECOMENDADA": analisis.get("accion_recomendada", ""),
                    "RESPUESTA_SUGERIDA": analisis.get("respuesta_sugerida", "")
                })
            except Exception as e:
                resultados.append({
                    "ARCHIVO": nombre,
                    "CATEGORIA": "ERROR DE LECTURA",
                    "ACCION_RECOMENDADA": f"Revisar manualmente ({e})",
                    "RESPUESTA_SUGERIDA": "VALIDAR MANUALMENTE"
                })

        st.session_state.novedades_data.extend(resultados)
        st.session_state.procesando = False
        st.success("✅ Análisis completado")

# =========================
# 📊 RESULTADOS
# =========================
if st.session_state.novedades_data:
    df = pd.DataFrame(st.session_state.novedades_data)
    st.subheader("📋 Resultado consolidado")
    st.dataframe(df, use_container_width=True)

    # 📥 Descargar Excel
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    st.download_button(
        label="⬇️ Descargar resultados en Excel",
        data=buffer,
        file_name="Novedades_Operativas_Resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================
# 🔄 OPCIÓN PARA LIMPIAR SESIÓN
# =========================
if st.button("🧹 Limpiar sesión"):
    st.session_state.novedades_data = []
    st.session_state.procesando = False
    st.experimental_rerun()
