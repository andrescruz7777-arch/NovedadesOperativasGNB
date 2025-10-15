# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: IA experta (detalle completo + acción recomendada sin asunto)
# =========================

import streamlit as st
import pandas as pd
import io, re, json
from datetime import datetime
import pdfplumber
from docx import Document
import extract_msg
from openai import OpenAI

# =========================
# ⚙️ CONFIGURACIÓN INICIAL
# =========================
st.set_page_config(page_title="Novedades Operativas GNB", layout="wide")
st.title("Analizador de Novedades Operativas - Contacto Solutions ⚖️")

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
    asunto = msg.subject or ""
    cuerpo = msg.body or ""
    remitente = msg.sender or ""
    return asunto, f"De: {remitente}\n\n{cuerpo}"

def leer_archivo_pdf(archivo):
    texto = ""
    with pdfplumber.open(archivo) as pdf:
        for pagina in pdf.pages:
            if pagina.extract_text():
                texto += pagina.extract_text() + "\n"
    return texto.strip()

def leer_archivo_docx(archivo):
    doc = Document(archivo)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def extraer_cc_y_nombre(texto):
    cc = ""
    nombre = ""
    cc_match = re.search(r"(?:CC[_\s:]*|CÉDULA[_\s:]*)?([0-9]{5,12})", texto, re.IGNORECASE)
    if cc_match:
        cc = cc_match.group(1).strip()
    nombre_match = re.search(r"([A-ZÁÉÍÓÚÑ ]{3,})\s*(?:CC|CÉDULA)", texto)
    if nombre_match:
        nombre = nombre_match.group(1).title().strip()
    return cc, nombre

# =========================
# 🧠 FUNCIÓN PRINCIPAL IA
# =========================
def analizar_novedad(texto):
    if not IA_DISPONIBLE:
        return {
            "categoria": "VALIDAR MANUALMENTE",
            "detalle_novedad": "No se pudo analizar el contenido. La IA no está disponible.",
            "accion_recomendada": "Revisar manualmente el correo.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "validado_ia": "No"
        }

    prompt = f"""
Eres un **abogado judicial colombiano senior con MBA**, especializado en **procesos ejecutivos bancarios, riesgos procesales y gestión de back office judicial**.

Tu tarea es analizar un correo o documento de novedad operativa del Banco GNB Sudameris y producir un informe completo, preciso y operativo para el equipo de back office.

🎯 INSTRUCCIONES CLAVE:
1. Usa solo **normas procesales reales de Colombia** (Ley 1564 de 2012 - CGP, Ley 2213 de 2022, etc.). No inventes leyes ni artículos.
2. Clasifica la novedad dentro de una **categoría general** (por ejemplo: “Errores de cargue documental”, “Desfase procesal”, “Fallas del sistema”, “Solicitud de actualización”, etc.).
3. En **DETALLE_NOVEDAD**, describe claramente todo lo que menciona el correo, sin omitir ninguna solicitud, comentario o detalle.
4. En **ACCION_RECOMENDADA**, explica paso a paso qué debe hacer el back office para resolver la novedad, con un lenguaje claro y práctico.
5. En **RESPUESTA_SUGERIDA**, redacta una respuesta formal, empática y profesional dirigida al banco.
6. Si la novedad no encaja en ninguna categoría conocida, **crea una nueva categoría** coherente con el contexto procesal y de riesgo operativo.

Responde estrictamente en formato JSON con las siguientes claves:
{{
  "categoria": "",
  "detalle_novedad": "",
  "accion_recomendada": "",
  "respuesta_sugerida": ""
}}

Texto a analizar:
{texto}
"""

    try:
        respuesta = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Eres un abogado colombiano con MBA, experto en litigio bancario y gestión judicial. Escribes de manera clara, exacta y con rigor procesal."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.25
        )

        contenido = respuesta.choices[0].message.content.strip()
        contenido = contenido.replace("```json", "").replace("```", "").strip()

        try:
            datos = json.loads(contenido)
        except Exception:
            try:
                datos = json.loads(contenido.replace('""', '"'))
            except Exception:
                datos = {
                    "categoria": "ERROR DE FORMATO",
                    "detalle_novedad": "La IA no devolvió un JSON válido.",
                    "accion_recomendada": "Validar manualmente.",
                    "respuesta_sugerida": contenido
                }

        datos["validado_ia"] = "Sí" if "ERROR" not in datos.get("categoria", "").upper() else "No"
        return datos

    except Exception as e:
        return {
            "categoria": "ERROR DE PROCESAMIENTO",
            "detalle_novedad": f"Error: {e}",
            "accion_recomendada": "Validar manualmente.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "validado_ia": "No"
        }

# =========================
# 🧾 INTERFAZ STREAMLIT
# =========================
st.subheader("📂 Cargar correos o documentos (.msg, .pdf, .docx)")
archivos = st.file_uploader(
    "Selecciona uno o varios archivos para analizar",
    type=["msg", "pdf", "docx"],
    accept_multiple_files=True
)

if archivos and st.button("🚀 Analizar Novedades"):
    st.session_state.procesando = True
    resultados = []

    for archivo in archivos:
        nombre = archivo.name
        extension = nombre.split(".")[-1].lower()
        try:
            if extension == "msg":
                asunto, texto = leer_archivo_msg(archivo)
                texto_completo = asunto + "\n\n" + texto
            elif extension == "pdf":
                texto_completo = leer_archivo_pdf(archivo)
            elif extension == "docx":
                texto_completo = leer_archivo_docx(archivo)
            else:
                texto_completo = ""

            cc, nombre_cli = extraer_cc_y_nombre(asunto)
            if not cc and not nombre_cli:
                cc, nombre_cli = extraer_cc_y_nombre(texto_completo)

            analisis = analizar_novedad(texto_completo)
            fecha_analisis = datetime.now().strftime("%Y-%m-%d %H:%M")

            resultados.append({
                "ARCHIVO": nombre,
                "CC": cc,
                "NOMBRE_CLIENTE": nombre_cli,
                "CATEGORIA": analisis.get("categoria", ""),
                "DETALLE_NOVEDAD": analisis.get("detalle_novedad", ""),
                "ACCION_RECOMENDADA": analisis.get("accion_recomendada", ""),
                "RESPUESTA_SUGERIDA": analisis.get("respuesta_sugerida", ""),
                "VALIDADO_IA": analisis.get("validado_ia", ""),
                "FECHA_ANALISIS": fecha_analisis
            })
        except Exception as e:
            resultados.append({
                "ARCHIVO": nombre,
                "CC": "",
                "NOMBRE_CLIENTE": "",
                "CATEGORIA": "ERROR DE LECTURA",
                "DETALLE_NOVEDAD": f"Revisar manualmente ({e})",
                "ACCION_RECOMENDADA": "VALIDAR MANUALMENTE",
                "RESPUESTA_SUGERIDA": "VALIDAR MANUALMENTE",
                "VALIDADO_IA": "No",
                "FECHA_ANALISIS": datetime.now().strftime("%Y-%m-%d %H:%M")
            })

    st.session_state.novedades_data.extend(resultados)
    st.session_state.procesando = False
    st.success("✅ Análisis completado correctamente.")

# =========================
# 📊 RESULTADOS Y RESUMEN EJECUTIVO
# =========================
if st.session_state.novedades_data:
    df = pd.DataFrame(st.session_state.novedades_data)
    st.subheader("📋 Resultado consolidado")
    st.dataframe(df, use_container_width=True)

    st.subheader(f"📊 Resumen ejecutivo del análisis preliminar ({len(df)} correos)")
    resumen = df.groupby("CATEGORIA").size().reset_index(name="Frecuencia")
    resumen["% del total"] = (resumen["Frecuencia"] / len(df) * 100).round(1)
    resumen["Impacto operativo"] = resumen["CATEGORIA"].map({
        "Errores de cargue documental": "🔴 Alto",
        "Desfase procesal (estado rama vs banco)": "🔴 Alto",
        "Errores de identificación del demandado": "🟠 Medio",
        "Duplicidad / cruces inconsistentes": "🟠 Medio",
        "Fallas en aplicativo o reportería": "🟡 Bajo–Medio",
        "Errores de notificación / comunicación": "🟡 Bajo",
        "Demoras de gestión / sin movimiento": "🟢 Medio–Alto",
    }).fillna("🟢 Bajo")

    st.dataframe(resumen, use_container_width=True)

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    st.download_button("⬇️ Descargar resultados en Excel", buffer, "Novedades_Operativas_Resultados.xlsx")

# =========================
# 🔄 LIMPIAR SESIÓN
# =========================
if st.button("🧹 Limpiar sesión"):
    st.session_state.novedades_data = []
    st.session_state.procesando = False
    st.rerun()
