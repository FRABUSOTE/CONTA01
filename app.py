import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("🧹 Limpieza de SERIE y NÚMERO de Documentos")

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

def extraer_serie_numero(nrodoc):
    if pd.isna(nrodoc):
        return pd.Series(["", ""])
    
    nrodoc = str(nrodoc).strip()
    nrodoc = re.sub(r'\s+', ' ', nrodoc)

    if '-' in nrodoc and not nrodoc.startswith(('FTF', 'FT', 'LT', 'CC', 'BV', 'PA', 'VR')):
        partes = nrodoc.split('-')
        if len(partes) == 2:
            letras = re.sub(r'\d', '', partes[0])
            numero = partes[1].strip()
            return pd.Series([letras, numero])

    match = re.match(r'^([A-Z]{3})(\d{3})(\d+)(-\d+)?$', nrodoc)
    if match:
        letras = match.group(1)
        serie_num = match.group(2)
        correlativo = match.group(3)
        sufijo = match.group(4) if match.group(4) else ''
        return pd.Series([letras[1] + serie_num, correlativo + sufijo])

    partes = nrodoc.split()
    if len(partes) == 2:
        serie = partes[0]
        numero = partes[1].lstrip('0')
        return pd.Series([serie, numero])

    match = re.match(r'^([A-Z]{2})(\d+)$', nrodoc)
    if match:
        letras = match.group(1)
        numero = match.group(2)
        return pd.Series([letras, str(int(numero))])

    return pd.Series(["", nrodoc])


if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if "NRODOC" not in df.columns:
        st.error("El archivo debe contener la columna 'NRODOC'.")
    else:
        st.success("Archivo cargado correctamente 🙌")

        df[['SERIE', 'NUMERO']] = df['NRODOC'].apply(extraer_serie_numero)

        st.write("### Vista previa del archivo procesado")
        st.dataframe(df.head())

        # Convertir a Excel para descarga
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="⬇ Descargar archivo limpio",
            data=buffer,
            file_name="reporte_limpio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
