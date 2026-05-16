
import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


st.set_page_config(
    page_title="Lista de Chequeo GLOBAL GAP",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_BASE = "norma_globalgap.xlsx"


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    mapa = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if c in ["Sección", "seccion"]:
            mapa[col] = "seccion"
        elif c in ["Principio"]:
            mapa[col] = "principio"
        elif c in ["Criterios", "criterio"]:
            mapa[col] = "criterio"
        elif c in ["Nivel"]:
            mapa[col] = "nivel"

    df = df.rename(columns=mapa)

    requeridas = ["seccion", "principio", "criterio", "nivel"]
    faltantes = [c for c in requeridas if c not in df.columns]

    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas: {faltantes}. "
            "El Excel debe tener: Sección, Principio, Criterios y Nivel."
        )

    return df[requeridas].copy()


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


@st.cache_data
def cargar_base() -> pd.DataFrame:
    raw_df = pd.read_excel(EXCEL_BASE)
    df = normalizar_columnas(raw_df)

    registros = []
    actual = None

    patron_criterio = re.compile(r"^FV-GFS\s+\d{2}\.\d{2}$")
    patron_capitulo = re.compile(r"^FV-GFS\s+\d{2}(\.0)?$")

    for _, row in df.iterrows():
        seccion = limpiar_texto(row["seccion"])
        principio = limpiar_texto(row["principio"])
        criterio = limpiar_texto(row["criterio"])
        nivel = limpiar_texto(row["nivel"])

        if patron_criterio.match(seccion):
            if actual is not None:
                registros.append(actual)

            actual = {
                "seccion": seccion,
                "principio": principio,
                "criterio": criterio,
                "nivel": nivel,
                "cumplimiento": "",
                "estado": "",
                "metodo de auditoria": "",
                "evidencia": "",
                "comentarios": "",
            }

        elif actual is not None and seccion == "" and principio == "" and criterio != "":
            actual["criterio"] = (actual["criterio"] + "\n" + criterio).strip()
            if nivel and not actual["nivel"]:
                actual["nivel"] = nivel

        elif patron_capitulo.match(seccion):
            continue

    if actual is not None:
        registros.append(actual)

    return pd.DataFrame(registros)


def calcular_resumen(df: pd.DataFrame):
    total = len(df)
    si = (df["estado"] == "Sí lo cumplo").sum()
    no = (df["estado"] == "No lo cumplo").sum()
    na = (df["estado"] == "No aplica").sum()
    pendientes = total - si - no - na
    avance = 0 if total == 0 else round(((si + no + na) / total) * 100, 1)
    cumplimiento = 0 if total == 0 else round((si / max(total - na, 1)) * 100, 1)

    return total, si, no, na, pendientes, avance, cumplimiento


def exportar_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    columnas = [
        "seccion",
        "principio",
        "criterio",
        "nivel",
        "cumplimiento",
        "estado",
        "metodo de auditoria",
        "evidencia",
        "comentarios",
    ]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[columnas].to_excel(writer, index=False, sheet_name="Auditoría")

    output.seek(0)
    wb = load_workbook(output)
    ws = wb["Auditoría"]

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9E2F3")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)

    widths = {
        "A": 16,
        "B": 45,
        "C": 90,
        "D": 20,
        "E": 18,
        "F": 18,
        "G": 35,
        "H": 35,
        "I": 45,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)

        estado = str(row[5].value).strip().lower()

        if estado == "sí lo cumplo":
            row[5].fill = PatternFill("solid", fgColor="C6EFCE")
        elif estado == "no lo cumplo":
            row[5].fill = PatternFill("solid", fgColor="FFC7CE")
        elif estado == "no aplica":
            row[5].fill = PatternFill("solid", fgColor="D9D9D9")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 95

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output.getvalue()


st.title("Sistema de Auditoría GLOBALG.A.P.")
st.caption("Plantilla digital para completar criterios de auditoría y exportar resultados.")

if "auditoria_df" not in st.session_state:
    st.session_state.auditoria_df = cargar_base()

df = st.session_state.auditoria_df

with st.sidebar:
    st.header("Filtros")

    capitulos = sorted(df["seccion"].str.extract(r"(FV-GFS\s+\d{2})")[0].dropna().unique())
    niveles = sorted(df["nivel"].dropna().unique())

    capitulo = st.selectbox("Capítulo", ["Todos"] + capitulos)
    nivel = st.selectbox("Nivel", ["Todos"] + niveles)
    estado = st.selectbox(
        "Estado",
        ["Todos", "Pendiente", "Sí lo cumplo", "No lo cumplo", "No aplica"]
    )
    busqueda = st.text_input("Buscar")

    st.markdown("---")

    if st.button("Reiniciar formulario"):
        st.session_state.auditoria_df = cargar_base()
        st.rerun()

filtro = pd.Series(True, index=df.index)

if capitulo != "Todos":
    filtro &= df["seccion"].str.startswith(capitulo, na=False)

if nivel != "Todos":
    filtro &= df["nivel"].eq(nivel)

if estado == "Pendiente":
    filtro &= df["estado"].eq("")
elif estado != "Todos":
    filtro &= df["estado"].eq(estado)

if busqueda.strip():
    b = busqueda.strip().lower()
    filtro &= (
        df["seccion"].str.lower().str.contains(b, na=False)
        | df["principio"].str.lower().str.contains(b, na=False)
        | df["criterio"].str.lower().str.contains(b, na=False)
    )

df_filtrado = df[filtro]

total, si, no, na, pendientes, avance, cumplimiento = calcular_resumen(df)

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Criterios", total)
m2.metric("Sí cumple", si)
m3.metric("No cumple", no)
m4.metric("No aplica", na)
m5.metric("Pendientes", pendientes)
m6.metric("Cumplimiento", f"{cumplimiento}%")

st.progress(avance / 100 if avance else 0)
st.caption(f"Avance de llenado: {avance}%")

st.markdown("### Criterios de auditoría")

if df_filtrado.empty:
    st.warning("No hay criterios que coincidan con los filtros.")
else:
    for idx, row in df_filtrado.iterrows():
        titulo = f"{row['seccion']} — {row['principio']}"

        with st.expander(titulo, expanded=False):
            st.markdown(f"**Nivel:** {row['nivel']}")

            st.markdown("**Criterio:**")
            st.text_area(
                label="criterio",
                value=row["criterio"],
                height=220,
                disabled=True,
                label_visibility="collapsed",
                key=f"criterio_{idx}"
            )

            c1, c2 = st.columns([1, 2])

            with c1:
                df.at[idx, "cumplimiento"] = st.selectbox(
                    "Cumplimiento",
                    ["", "Mayor", "Menor"],
                    index=["", "Mayor", "Menor"].index(df.at[idx, "cumplimiento"])
                    if df.at[idx, "cumplimiento"] in ["", "Mayor", "Menor"] else 0,
                    key=f"cumplimiento_{idx}"
                )

            with c2:
                df.at[idx, "estado"] = st.radio(
                    "Estado",
                    ["", "Sí lo cumplo", "No lo cumplo", "No aplica"],
                    index=["", "Sí lo cumplo", "No lo cumplo", "No aplica"].index(df.at[idx, "estado"])
                    if df.at[idx, "estado"] in ["", "Sí lo cumplo", "No lo cumplo", "No aplica"] else 0,
                    horizontal=True,
                    key=f"estado_{idx}"
                )

            df.at[idx, "metodo de auditoria"] = st.text_area(
                "Método de auditoría",
                value=df.at[idx, "metodo de auditoria"],
                height=90,
                key=f"metodo_{idx}"
            )

            df.at[idx, "evidencia"] = st.text_area(
                "Evidencia",
                value=df.at[idx, "evidencia"],
                height=90,
                key=f"evidencia_{idx}"
            )

            df.at[idx, "comentarios u observaciones"] = st.text_area(
                "Comentarios u observaciones",
                value=df.at[idx, "comentarios u observaciones"],
                height=90,
                key=f"obs_{idx}"
            )

st.session_state.auditoria_df = df

st.markdown("---")

excel_bytes = exportar_excel(st.session_state.auditoria_df)
fecha = datetime.now().strftime("%Y%m%d_%H%M")

st.download_button(
    label="Descargar auditoría completada en Excel",
    data=excel_bytes,
    file_name=f"auditoria_globalgap_completada_{fecha}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
