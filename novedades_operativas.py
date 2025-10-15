# =========================
# 📬 ANALIZADOR DE NOVEDADES OPERATIVAS GNB
# Autor: Andrés Cruz - Contacto Solutions
# Versión: Detección de CC y nombre desde asunto + precisión legal colombiana
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
    """Lee correos .msg de Outlook e incluye asunto y cuerpo."""
    msg = extract_msg.Message(archivo)
    asunto = msg.subject or ""
    cuerpo = msg.body or ""
    remitente = msg.sender or ""
    return asunto, f"De: {remitente}\nAsunto: {asunto}\n\n{cuerpo}"

def leer_archivo_pdf(archivo):
    """Lee texto de archivos PDF"""
    texto = ""
    with pdfplumber.open(archivo) as pdf:
        for pagina in pdf.pages:
            page_text = pagina.extract_text()
            if page_text:
                texto += page_text + "\n"
    return texto.strip()

def leer_archivo_docx(archivo):
    """Lee texto de archivos Word .docx"""
    doc = Document(archivo)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def extraer_cc_y_nombre(texto):
    """Detecta número de cédula y nombre si aparecen en el texto o asunto."""
    cc = ""
    nombre = ""
    # Buscar CC
    cc_match = re.search(r"CC[:\s_]*([0-9\.\-]+)", texto, re.IGNORECASE)
    if cc_match:
        cc = cc_match.group(1).replace(".", "").replace("-", "").strip()
    # Buscar nombre antes de CC (en mayúsculas)
    nombre_match = re.search(r"([A-ZÁÉÍÓÚÑ ]{3,})\s*CC", texto)
    if nombre_match:
        nombre = nombre_match.group(1).title().strip()
    return cc, nombre

def analizar_novedad(texto):
    """Analiza la novedad con base jurídica colombiana real y explicación operativa clara."""
    if not IA_DISPONIBLE:
        return {
            "categoria": "VALIDAR MANUALMENTE",
            "accion_recomendada": "Revisar manualmente. La IA no está disponible.",
            "respuesta_sugerida": "VALIDAR MANUALMENTE",
            "validado_ia": "No"
        }

    prompt = f"""
Actúa como un **abogado judicial colombiano senior**, con formación en **MBA, gestión de riesgos procesales y dirección jurídica**.
Analiza un correo o documento remitido por el **Banco GNB Sudameris** como una **novedad operativa (PQR)** dirigida al área jurídica o back office judicial.

🧠 Tu perfil:
- Experto en procesos ejecutivos bancarios bajo el **Código General del Proceso (Ley 1564 de 2012)** y la **Ley 2213 de 2022** sobre medios electrónicos.
- NO inventes leyes ni artículos. Usa solo normas procesales **reales y vigentes de Colombia**.
- Si no aplica citar norma, explica en lenguaje operativo lo que debe hacerse (no técnico).
- El objetivo es guiar al **back office**, no emitir conceptos jurídicos.

🎯 Objetivo:
1. Clasifica la novedad en una **categoría principal**. Usa una de las siguientes si aplica:
   - Errores de cargue documental
   - Desfase procesal (estado rama vs banco)
   - Errores de identificación del demandado
   - Duplicidad / cruces inconsistentes
   - Fallas en aplicativo o reportería
   - Errores de notificación / comunicación
   - Demoras de gestión / sin movimiento
   - Otras (especificar)
2. Si ninguna encaja exactamente, **crea una nueva categoría breve y clara**.
3. Indica una **acción recomendada** simple y ejecutable por el back office (qué revisar, cómo corregir o a quién escalar).
4. Redacta una **respuesta sugerida profesional y empática**, como si respondieras al correo del banco, clara y sin tecnicismos.
5. No cites leyes extranjeras ni inventadas.

Responde solo en formato JSON con esta estructura:
{{
  "categoria": "",
  "accion_recomendada": "",
  "respuesta_sugerida": ""
}}

No uses bloques de código (no ```json ni ```).

Texto a analizar:
{texto}
"""

    try:
        respuesta = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Eres un abogado colombiano con MBA, especializado en litigio bancario, riesgos procesales y liderazgo de back office. Explicas con precisión legal real y lenguaje claro."},
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
st.subheader("📂 Cargar correos o documentos (.msg, .pdf, .docx)")
archivos = st.file_uploader(
    "Selecciona uno o varios archivos para analizar",
    type=["msg", "pdf", "docx"],
    accept_multiple_files=True
)

if archivos:
    if st.button("🚀 Analizar Novedades"):
        st.session_state.procesando = True
        resultados = []

        for archivo in archivos:
            nombre = archivo.name
            extension = nombre.split(".")[-1].lower()
            asunto = ""
            try:
                if extension == "msg":
                    asunto, texto = leer_archivo_msg(archivo)
                elif extension == "pdf":
                    asunto = ""
                    texto = leer_archivo_pdf(archivo)
                elif extension == "docx":
                    asunto = ""
                    texto = leer_archivo_docx(archivo)
                else:
                    texto = ""

                # Buscar CC y nombre primero en el asunto, luego en el cuerpo
                cc, nombre_cli = extraer_cc_y_nombre(asunto)
                if not cc and not nombre_cli:
                    cc, nombre_cli = extraer_cc_y_nombre(texto)

                analisis = analizar_novedad(texto)
                fecha_analisis = datetime.now().strftime("%Y-%m-%d %H:%M")

                resultados.append({
                    "ARCHIVO": nombre,
                    "ASUNTO": asunto,
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
                    "ASUNTO": asunto,
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
if st.button("🧹 Limpiar sesión"):
    st.session_state.novedades_data = []
    st.session_state.procesando = False
    st.rerun()
