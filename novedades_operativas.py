# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: Compatible con openai>=1.0.0 + Resumen ejecutivo y CC/nombre/fecha
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

def extraer_cc_y_nombre(texto):
    """Busca CC y nombre dentro del texto"""
    cc = ""
    nombre = ""

    # Buscar cédula (CC 12345678 o 1.234.567.890)
    cc_match = re.search(r"CC[:\s_]*([0-9\.\-]+)", texto, re.IGNORECASE)
    if cc_match:
        cc = cc_match.group(1).replace(".", "").replace("-", "").strip()

    # Buscar posible nombre (antes o después de la cédula)
    nombre_match = re.search(r"([A-ZÁÉÍÓÚÑ ]{3,})\s*CC", texto)
    if nombre_match:
        nombre = nombre_match.group(1).title().strip()

    return cc, nombre

def analizar_novedad(texto):
    """Analiza el texto de la novedad con IA"""
    if not IA_DISPONIBLE:
        return {
            "categoria": "VALIDAR MANUALMENTE",
            "accion_recomendada": "Revisar manualmente el contenido. La IA no está disponible.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "validado_ia": "No"
        }

    prompt = f"""
Analiza el siguiente correo o documento de novedad operativa y clasifícalo según su naturaleza.

Texto:
{texto}

Responde estrictamente en formato JSON con las siguientes claves:
- categoria
- accion_recomendada
- respuesta_sugerida

No incluyas comillas triples ni bloques de código (no uses ```json ni ```).
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

        contenido = respuesta.choices[0].message.content.strip()
        contenido = contenido.replace("```json", "").replace("```", "").strip()

        try:
            datos = json.loads(contenido)
        except Exception:
            try:
                contenido_corr = contenido.replace('""', '"')
                datos = json.loads(contenido_corr)
            except Exception:
                datos = {
                    "categoria": "ERROR DE FORMATO",
                    "accion_recomendada": "La IA no devolvió un JSON válido.",
                    "respuesta_sugerida": contenido
                }

        if datos.get("categoria", "").upper() not in ["ERROR DE FORMATO", "ERROR DE PROCESAMIENTO"]:
            datos["validado_ia"] = "Sí"
        else:
            datos["validado_ia"] = "No"

        return datos

    except Exception as e:
        return {
            "categoria": "ERROR DE PROCESAMIENTO",
            "accion_recomendada": f"Validar manualmente. Error: {e}",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "validado_ia": "No"
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

                cc, nombre_cli = extraer_cc_y_nombre(texto)
                analisis = analizar_novedad(texto)
                fecha_analisis = datetime.now().strftime("%Y-%m-%d %H:%M")

                resultados.append({
                    "ARCHIVO": nombre,
                    "CC": cc,
                    "NOMBRE_CLIENTE": nombre_cli,
                    "CATEGORIA": analisis.get("categoria", ""),
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
                    "ACCION_RECOMENDADA": f"Revisar manualmente ({e})",
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

    # =========================
    # 📊 RESUMEN EJECUTIVO
    # =========================
    st.subheader(f"📊 Resumen ejecutivo del análisis preliminar ({len(df)} correos)")
    resumen = df.groupby("CATEGORIA").size().reset_index(name="Frecuencia")
    resumen["% del total"] = (resumen["Frecuencia"] / len(df) * 100).round(1)

    # Asignar impacto operativo (puedes ajustar estos valores)
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

    # Descargar Excel/CSV
    buffer = io.BytesIO()
    try:
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            label="⬇️ Descargar resultados en Excel",
            data=buffer,
            file_name="Novedades_Operativas_Resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Descargar resultados en CSV",
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
