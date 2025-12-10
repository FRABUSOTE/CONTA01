import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 Limpieza y Ordenamiento de Libro Mayor – Cuenta 60 Importaciones")

uploaded_file = st.file_uploader("Sube tu archivo Excel (Libro Mayor)", type=["xlsx"])


# ---------------- FUNCIÓN: extraer número de PED ----------------
def extraer_pedido(glosa):
    match = re.search(r'PED\.?\s*(\d+)', str(glosa), re.IGNORECASE)
    if match:
        return int(match.group(1))
    return float('inf')  # Para filas sin PED


# ---------------- FUNCIÓN: extraer número y nombre de cuenta ----------------
def extraer_cuenta(sub):
    sub = str(sub)
    if "Cuenta:" in sub:
        texto = sub.split("Cuenta:")[-1].strip()
        texto = texto.replace("Saldo mes anterior:", "").strip()
        partes = texto.split(maxsplit=1)
        if len(partes) == 2:
            return partes[0], partes[1]
        elif len(partes) == 1:
            return partes[0], None
    return None, None


# ---------------- PROCESAR ARCHIVO ----------------
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Crear columnas
    df["PED_NUM"] = df["Glosa"].apply(extraer_pedido)
    df[["Cuenta_Numero", "Cuenta_Nombre"]] = df["Sub."].apply(
        lambda x: pd.Series(extraer_cuenta(x))
    )

    # Propagar cuenta
    cuenta_num_actual = None
    cuenta_nom_actual = None

    for i, valor in enumerate(df["Sub."]):
        if isinstance(valor, str) and valor.strip().startswith("Cuenta:"):
            cuenta_num_actual, cuenta_nom_actual = extraer_cuenta(valor)
        else:
            df.at[i, "Cuenta_Numero"] = cuenta_num_actual
            df.at[i, "Cuenta_Nombre"] = cuenta_nom_actual

    # Eliminar filas basura
    def fila_valida(row):
        if pd.isna(row["Glosa"]):
            return False
        if str(row["Glosa"]).strip() == "":
            return False
        if "Total" in str(row["Glosa"]) or "Saldo" in str(row["Glosa"]):
            return False
        return True

    df = df[df.apply(fila_valida, axis=1)].reset_index(drop=True)

    # Orden final
    df = df.sort_values(by=["PED_NUM", "Fec.Reg."], ascending=[True, True])
    df = df.drop(columns=["PED_NUM"])

    st.success("Archivo procesado correctamente ✔")
    st.write("### Vista previa del archivo limpio")
    st.dataframe(df.head(50))

    # Convertir a Excel para descarga
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        label="⬇ Descargar archivo limpio",
        data=buffer,
        file_name="resultado_final_limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
