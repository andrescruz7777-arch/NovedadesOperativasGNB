# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: Compatible con openai>=1.0.0
# =========================

import streamlit as st
import pandas as pd
import io
import pdfplumber
from docx import Document
import extract_msg
import json
from openai import OpenAI

# =========================
# ⚙️ CONFIGURACIÓN INICIAL
# =========================
st.set_page_config(page_title="Novedades Operativas GNB", layout="wide")
st.title("Analizador de Novedades Operativas - Contacto Solutions")

# =========================
# 🔐 CONFIGURACIÓN DE API
# =========================
IA_DISPONIBLE = False
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    IA_DISPONIBLE = True
    st.info("✅ Conexión con OpenAI establecida correctamente.")
except Exception as e:
    st.warning(f"⚠️ No se pudo inicializar la IA. Error: {e}")

# =========================
# 📁 VARIABLES GLOBALES
# =========================
if "novedades_data" not in st.session_state:
    st.session_state["novedades_data"] = []
if "procesando" not in st.session_state:
    st.session_state["procesando"] = False

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
            page_text = pagina.extract_text()
            if page_text:
                texto += page_text + "\n"
    return texto.strip()

def leer_archivo_docx(archivo):
    doc = Document(archivo)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def analizar_novedad(texto):
    if not IA_DISPONIBLE:
        return {
            "categoria": "VALIDAR MANUALMENTE",
            "accion_recomendada": "Revisar manualmente el contenido. La IA no está disponible.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE"
        }

    prompt = f"""
Analiza el siguiente correo o documento de novedad operativa y clasifícalo según su naturaleza.

Texto:
{texto}

Responde estrictamente en formato JSON con las siguientes claves:
- categoria
- accion_recomendada
- respuesta_sugerida
"""

    try:
        respuesta = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Eres un abogado experto en operaciones judiciales bancarias."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3
        )

        contenido = respuesta.choices[0].message.content

        try:
            datos = json.loads(contenido)
        except Exception:
            datos = {
                "categoria": "ERROR DE FORMATO",
                "accion_recomendada": "La IA no devolvió un JSON válido.",
                "respuesta_sugerida": contenido
            }

        return datos

    except Exception as e:
        return {
            "categoria": "ERROR DE PROCESAMIENTO",
            "accion_recomendada": f"Validar manualmente. Error: {e}",
            "respuesta_sugerida": "VALIDAR MANUALMENTE"
        }

# =========================
# 🧾 INTERFAZ DE CARGA
# =========================
st.subheader("Cargar correos o documentos (.msg, .pdf, .docx)")
archivos = st.file_uploader(
    "Selecciona uno o varios archivos",
    type=["msg", "pdf", "docx"],
    accept_multiple_files=True
)

if archivos:
    if st.button("Analizar Novedades"):
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
        st.success("✅ Análisis completado correctamente.")

# =========================
# 📊 RESULTADOS
# =========================
if st.session_state.novedades_data:
    df = pd.DataFrame(st.session_state.novedades_data)
    st.subheader("Resultado consolidado")
    st.dataframe(df, use_container_width=True)

    buffer = io.BytesIO()
    try:
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            label="Descargar resultados en Excel",
            data=buffer,
            file_name="Novedades_Operativas_Resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Descargar resultados en CSV",
            data=csv_data,
            file_name="Novedades_Operativas_Resultados.csv",
            mime="text/csv"
        )

# =========================
# 🔄 LIMPIAR SESIÓN
# =========================
if st.button("Limpiar sesión"):
    st.session_state.novedades_data = []
    st.session_state.procesando = False
    st.rerun()
