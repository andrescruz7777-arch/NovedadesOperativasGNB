# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: IA con razonamiento autónomo (abogado judicial colombiano + MBA)
# =========================

import streamlit as st
import pandas as pd
import io, re, json, os
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
    return asunto, f"De: {remitente}\nAsunto: {asunto}\n\n{cuerpo}"

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
            "accion_recomendada": "Revisar manualmente. La IA no está disponible.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "accion_automatizada": "Sin IA disponible",
            "validado_ia": "No"
        }

    prompt = f"""
Actúa como un **abogado judicial colombiano senior con MBA**, especializado en **procesos ejecutivos bancarios, riesgos procesales y gestión de back office judicial**.
Tu papel es analizar un correo o documento de novedad operativa del **Banco GNB Sudameris** y emitir un análisis claro, completo y ejecutable.

📋 **Reglas:**
1. Usa solo normas procesales **reales y vigentes en Colombia** (CGP - Ley 1564 de 2012, Ley 2213 de 2022, etc.).  
2. No inventes leyes ni artículos.
3. Si la novedad dice que el **sistema está desactualizado**, instruye que se busque en la carpeta compartida del cliente `/mnt/shared/clientes/[CC]/` y se cargue el soporte en el aplicativo.
4. Si menciona que el **juzgado está incorrecto o no coincide**, indica validar en la página oficial de la **Rama Judicial** y actualizar los datos.
5. Si hay **más de una solicitud en el correo**, incluye **todas** sin omitir ninguna.
6. Si la novedad **no encaja en las categorías anteriores**, **razona y crea una acción y categoría nuevas**, de acuerdo con tu criterio jurídico-operativo y perfil profesional.

🎯 Tu salida debe ser clara y estructurada en formato JSON con los siguientes campos:
{{
  "categoria": "tipo de novedad o incidencia procesal detectada",
  "accion_recomendada": "qué debe hacer el back office, paso a paso",
  "respuesta_sugerida": "texto formal y empático para responder al banco",
  "accion_automatizada": "instrucción práctica para ejecutar en el sistema o la carpeta compartida"
}}

Texto del correo o documento:
{texto}
"""

    try:
        respuesta = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Eres un abogado judicial colombiano con MBA, experto en riesgo procesal y operaciones judiciales bancarias. Tu tono es técnico, empático y orientado a la acción."},
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
                    "accion_recomendada": "La IA no devolvió un JSON válido.",
                    "respuesta_sugerida": contenido,
                    "accion_automatizada": "Sin acción detectada"
                }

        datos["validado_ia"] = "Sí" if "ERROR" not in datos.get("categoria", "").upper() else "No"
        return datos

    except Exception as e:
        return {
            "categoria": "ERROR DE PROCESAMIENTO",
            "accion_recomendada": f"Validar manualmente. Error: {e}",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "accion_automatizada": "Sin acción detectada",
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
        asunto = ""
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
                "ASUNTO": asunto,
                "CC": cc,
                "NOMBRE_CLIENTE": nombre_cli,
                "CATEGORIA": analisis.get("categoria", ""),
                "ACCION_RECOMENDADA": analisis.get("accion_recomendada", ""),
                "RESPUESTA_SUGERIDA": analisis.get("respuesta_sugerida", ""),
                "ACCION_AUTOMATIZADA": analisis.get("accion_automatizada", ""),
                "VALIDADO_IA": analisis.get("validado_ia", ""),
                "FECHA_ANALISIS": fecha_analisis
            })
        except Exception as e:
            resultados.append({
                "ARCHIVO": nombre,
                "ASUNTO": asunto,
                "CC": "",
                "NOMBRE_CLIENTE": "",
                "CATEGORIA": "ERROR DE LECTURA",
                "ACCION_RECOMENDADA": f"Revisar manualmente ({e})",
                "RESPUESTA_SUGERIDA": "VALIDAR MANUALMENTE",
                "ACCION_AUTOMATIZADA": "Sin acción detectada",
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

    impacto_map = {
        "Errores de cargue documental": "🔴 Alto",
        "Desfase procesal (estado rama vs banco)": "🔴 Alto",
        "Errores de identificación del demandado": "🟠 Medio",
        "Duplicidad / cruces inconsistentes": "🟠 Medio",
        "Fallas en aplicativo o reportería": "🟡 Bajo–Medio",
        "Errores de notificación / comunicación": "🟡 Bajo",
        "Demoras de gestión / sin movimiento": "🟢 Medio–Alto",
    }
    resumen["Impacto operativo"] = resumen["CATEGORIA"].map(impacto_map).fillna("🟢 Bajo")

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
