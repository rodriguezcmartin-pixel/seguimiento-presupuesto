import streamlit as st
import pandas as pd
from datetime import datetime
import os
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ---------------------------
# CONFIGURACIÓN
# ---------------------------
st.set_page_config(page_title="Seguimiento Presupuesto", layout="centered")

# ---------------------------
# LOGO
# ---------------------------
st.image("logo.png", width=180)

st.title("💰 Seguimiento de Presupuesto")

# ---------------------------
# FORMULARIO
# ---------------------------
with st.form("presupuesto_form"):

    numero_albaran = st.text_input("Número de albarán")

    fecha = st.date_input("Fecha", datetime.today())

    trabajador = st.text_input("Trabajador")

    partida = st.selectbox(
        "Partida del presupuesto",
        [
            "Instalación eléctrica",
            "Domótica",
            "Cuadro eléctrico",
            "Cableado",
            "Iluminación",
            "Automatismos",
            "Otros"
        ]
    )

    gasto = st.number_input(
        "Gasto de la partida (€)",
        min_value=0.0,
        step=0.01
    )

    comentarios = st.text_area("Comentarios")

    # FOTO ALBARÁN
    foto = st.file_uploader(
        "Subir foto del albarán",
        type=["jpg", "jpeg", "png"]
    )

    guardar = st.form_submit_button("Guardar registro")

# ---------------------------
# DATAFRAME
# ---------------------------
if "datos" not in st.session_state:
    st.session_state.datos = pd.DataFrame(columns=[
        "Albarán",
        "Fecha",
        "Trabajador",
        "Partida",
        "Gasto (€)",
        "Comentarios"
    ])

# ---------------------------
# GUARDAR DATOS
# ---------------------------
if guardar:

    nuevo = pd.DataFrame({
        "Albarán": [numero_albaran],
        "Fecha": [fecha],
        "Trabajador": [trabajador],
        "Partida": [partida],
        "Gasto (€)": [gasto],
        "Comentarios": [comentarios]
    })

    st.session_state.datos = pd.concat(
        [st.session_state.datos, nuevo],
        ignore_index=True
    )

    # GUARDAR FOTO
    if foto is not None:

        os.makedirs("fotos_albaranes", exist_ok=True)

        ruta = os.path.join(
            "fotos_albaranes",
            foto.name
        )

        with open(ruta, "wb") as f:
            f.write(foto.getbuffer())

    st.success("Registro guardado correctamente")

# ---------------------------
# MOSTRAR TABLA
# ---------------------------
st.subheader("📊 Registros")

st.dataframe(st.session_state.datos)

# ---------------------------
# TOTAL GASTOS
# ---------------------------
if not st.session_state.datos.empty:

    total = st.session_state.datos["Gasto (€)"].sum()

    st.metric("💵 Total gastos", f"{total:.2f} €")

# ---------------------------
# GENERAR EXCEL
# ---------------------------
excel_file = "seguimiento_presupuesto.xlsx"

if not st.session_state.datos.empty:

    st.session_state.datos.to_excel(
        excel_file,
        index=False
    )

    with open(excel_file, "rb") as f:

        st.download_button(
            label="⬇️ Descargar Excel",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------------------
# ENVIAR EMAIL
# ---------------------------
st.subheader("📧 Enviar Excel por Email")

correo_origen = st.text_input("Correo Gmail")

password = st.text_input(
    "Contraseña de aplicación",
    type="password"
)

correo_destino = st.text_input("Correo destinatario")

if st.button("Enviar correo"):

    try:

        mensaje = MIMEMultipart()

        mensaje["From"] = correo_origen
        mensaje["To"] = correo_destino
        mensaje["Subject"] = "Seguimiento presupuesto obra"

        archivo = open(excel_file, "rb")

        parte = MIMEBase("application", "octet-stream")

        parte.set_payload(archivo.read())

        encoders.encode_base64(parte)

        parte.add_header(
            "Content-Disposition",
            f"attachment; filename={excel_file}"
        )

        mensaje.attach(parte)

        servidor = smtplib.SMTP("smtp.gmail.com", 587)

        servidor.starttls()

        servidor.login(correo_origen, password)

        servidor.send_message(mensaje)

        servidor.quit()

        st.success("Correo enviado correctamente")

    except Exception as e:

        st.error(f"Error al enviar correo: {e}")