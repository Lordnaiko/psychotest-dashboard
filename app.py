import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="PSYCHOTEST SpA | Inteligencia Comercial",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados (Alto contraste en modo oscuro, estilo tarjetas de la 3ra imagen)
st.markdown("""
    <style>
    /* Fondo principal y texto base */
    .stApp, [data-testid="stAppViewContainer"] {
        background-color: #0B1120;
        color: #FFFFFF;
    }
    [data-testid="stHeader"] {
        background-color: #0B1120;
    }

    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1E293B;
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #F8FAFC !important;
    }

    /* Zona y botón Upload del file_uploader */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: #1E293B;
        border: 1px dashed #38BDF8;
        color: #F8FAFC !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border: 1px solid #38BDF8 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {
        color: #FFFFFF !important;
    }

    /* Tarjeta de estado / confirmación con el mismo estilo de las tarjetas KPI */
    .status-card {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-top: 4px solid #10B981;
        padding: 16px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.35);
        margin-bottom: 16px;
    }
    .status-card-title {
        font-size: 14px;
        font-weight: 600;
        color: #6EE7B7 !important;
        margin: 0 0 6px 0;
    }
    .status-card-value {
        font-size: 20px;
        font-weight: 700;
        color: #FFFFFF !important;
        margin: 0 0 6px 0;
    }
    .status-card-sub {
        font-size: 13px;
        color: #E2E8F0 !important;
        margin: 0;
    }

    /* Subtítulos y encabezados */
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] p,
    .stMarkdown p,
    h1, h2, h3, h4 {
        color: #F8FAFC !important;
    }

    /* Tarjetas KPI */
    div[data-testid="stMetric"] {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-top: 4px solid #38BDF8;
        padding: 18px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.35);
    }
    [data-testid="stMetricLabel"],
    [data-testid="stMetricLabel"] *,
    [data-testid="stMetricLabel"] p,
    [data-testid="stMetricLabel"] span,
    [data-testid="stMetricLabel"] div {
        font-size: 15px !important;
        font-weight: 600 !important;
        color: #7DD3FC !important;
    }
    [data-testid="stMetricValue"],
    [data-testid="stMetricValue"] *,
    [data-testid="stMetricValue"] div {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
    }

    /* Cajas de recomendaciones */
    .insight-box {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-left: 5px solid #38BDF8;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        margin-bottom: 16px;
        color: #F8FAFC;
    }
    .insight-box h4 {
        color: #38BDF8 !important;
        margin-top: 0;
        margin-bottom: 10px;
    }
    .insight-box p, .insight-box b {
        color: #F1F5F9 !important;
    }
    .insight-box-green {
        border-left-color: #10B981;
    }
    .insight-box-green h4 {
        color: #34D399 !important;
    }

    /* Pestañas */
    button[data-baseweb="tab"] p,
    button[data-baseweb="tab"] span {
        color: #E2E8F0 !important;
        font-weight: 600;
        font-size: 15px;
    }
    button[data-baseweb="tab"][aria-selected="true"] p,
    button[data-baseweb="tab"][aria-selected="true"] span {
        color: #38BDF8 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        border-bottom-color: #38BDF8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Diccionario de nombres de meses en español
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def formato_clp(valor):
    return f"${valor:,.0f}".replace(",", ".")

# Encabezado corporativo
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("logo.png", width=130)
    except:
        st.markdown("### PSYCHOTEST SpA")
with col_titulo:
    st.title("PSYCHOTEST SpA — Inteligencia Comercial")
    st.caption("Panel Ejecutivo de Rendimiento de Ventas y Planificación Estratégica")

st.divider()

# Barra lateral para carga y filtros
with st.sidebar:
    st.header("Configuración de Datos")
    archivos = st.file_uploader(
        "Cargar reportes de ventas (.xlsx o .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )

def generar_analisis_ia_proyectivo(df_filtrado, col_neto, col_boleta):
    """Analiza estacionalidad por mes calendario y proyecta recomendaciones a futuro."""
    estacionalidad = df_filtrado.groupby("Mes_Num", as_index=False).agg(
        Ventas_Totales=(col_neto, "sum"),
        Boletas_Totales=(col_boleta, "count")
    )
    estacionalidad["Mes_Nombre"] = estacionalidad["Mes_Num"].map(MESES_ES)
    estacionalidad["Ticket_Mes"] = estacionalidad["Ventas_Totales"] / estacionalidad["Boletas_Totales"]

    promedio_mensual = estacionalidad["Ventas_Totales"].mean()
    total_global = estacionalidad["Ventas_Totales"].sum()

    est_sorted = estacionalidad.sort_values(by="Ventas_Totales", ascending=False).reset_index(drop=True)

    mejor_mes_num = int(est_sorted.loc[0, "Mes_Num"])
    mejor_mes_nombre = est_sorted.loc[0, "Mes_Nombre"]
    monto_mejor = est_sorted.loc[0, "Ventas_Totales"]
    boletas_mejor = int(est_sorted.loc[0, "Boletas_Totales"])
    ticket_mejor = est_sorted.loc[0, "Ticket_Mes"]

    pct_sobre_promedio = ((monto_mejor - promedio_mensual) / promedio_mensual) * 100 if promedio_mensual > 0 else 0
    participacion_pct = (monto_mejor / total_global) * 100 if total_global > 0 else 0

    mes_previo_num = 12 if mejor_mes_num == 1 else mejor_mes_num - 1
    mes_previo_nombre = MESES_ES[mes_previo_num]

    hoy = datetime.now()
    anio_objetivo_previo = hoy.year if mes_previo_num >= hoy.month else hoy.year + 1
    anio_objetivo_pico = anio_objetivo_previo + 1 if (mes_previo_num == 12 and mejor_mes_num == 1) else anio_objetivo_previo

    peor_mes_nombre = est_sorted.loc[len(est_sorted) - 1, "Mes_Nombre"]
    monto_peor = est_sorted.loc[len(est_sorted) - 1, "Ventas_Totales"]

    top_3_nombres = est_sorted.head(3)["Mes_Nombre"].tolist()

    return {
        "mejor_mes_nombre": mejor_mes_nombre,
        "anio_objetivo_pico": anio_objetivo_pico,
        "mes_previo_nombre": mes_previo_nombre,
        "anio_objetivo_previo": anio_objetivo_previo,
        "monto_mejor": monto_mejor,
        "boletas_mejor": boletas_mejor,
        "ticket_mejor": ticket_mejor,
        "pct_sobre_promedio": pct_sobre_promedio,
        "participacion_pct": participacion_pct,
        "peor_mes_nombre": peor_mes_nombre,
        "monto_peor": monto_peor,
        "top_3": ", ".join(top_3_nombres)
    }

if archivos:
    lista_dfs = []
    resumen_archivos = []

    for archivo in archivos:
        try:
            if archivo.name.endswith('.csv'):
                temp_df = pd.read_csv(archivo, sep=None, engine='python')
            else:
                temp_df = pd.read_excel(archivo)

            col_fecha_tmp = next((col for col in temp_df.columns if 'fecha' in col.lower()), None)
            col_neto_tmp = next((col for col in temp_df.columns if 'neto' in col.lower() or 'total' in col.lower() or 'venta' in col.lower()), temp_df.columns[1] if len(temp_df.columns) > 1 else None)

            monto_arch = 0
            filas_arch = 0
            if col_fecha_tmp and col_neto_tmp is not None:
                f_proc = pd.to_datetime(temp_df[col_fecha_tmp], errors='coerce', dayfirst=True)
                n_proc = pd.to_numeric(temp_df[col_neto_tmp], errors='coerce')
                valido = f_proc.notna() & n_proc.notna()
                monto_arch = n_proc[valido].sum()
                filas_arch = int(valido.sum())

            resumen_archivos.append((archivo.name, filas_arch, monto_arch))
            lista_dfs.append(temp_df)
        except Exception as e:
            st.error(f"Error al procesar {archivo.name}: {e}")

    if lista_dfs:
        df = pd.concat(lista_dfs, ignore_index=True)

        col_fecha = next((col for col in df.columns if 'fecha' in col.lower()), None)
        col_boleta = next((col for col in df.columns if 'boleta' in col.lower() or 'folio' in col.lower()), df.columns[0])
        col_neto = next((col for col in df.columns if 'neto' in col.lower() or 'total' in col.lower() or 'venta' in col.lower()), df.columns[1])

        if col_fecha and col_neto is not None:
            df["Fecha_Proc"] = pd.to_datetime(df[col_fecha], errors='coerce', dayfirst=True)
            df[col_neto] = pd.to_numeric(df[col_neto], errors='coerce')
            df = df.dropna(subset=["Fecha_Proc", col_neto])

            df["Año"] = df["Fecha_Proc"].dt.year
            df["Mes_Num"] = df["Fecha_Proc"].dt.month
            df["Mes_Nombre"] = df["Mes_Num"].map(MESES_ES)
            df["Año_Mes"] = df["Fecha_Proc"].dt.to_period("M").astype(str)

            años_disponibles = sorted([int(a) for a in df["Año"].unique()])
            detalle_archivos_html = "<br>".join(
                [f"• {nom}: <b>{formato_clp(m)}</b> ({n} reg.)" for nom, n, m in resumen_archivos]
            )

            with st.sidebar:
                año_sel = st.multiselect("Filtrar por Año", años_disponibles, default=años_disponibles)
                st.markdown(f"""
                <div class="status-card">
                    <p class="status-card-title">Montos cargados con éxito</p>
                    <p class="status-card-value">{formato_clp(df[col_neto].sum())}</p>
                    <p class="status-card-sub">Registros válidos activos: <b>{len(df)}</b><br>{detalle_archivos_html}</p>
                </div>
                """, unsafe_allow_html=True)

            df_filtrado = df[df["Año"].isin(año_sel)] if año_sel else df

            total_ventas = df_filtrado[col_neto].sum()
            total_boletas = df_filtrado[col_boleta].count()
            ticket_promedio = total_ventas / total_boletas if total_boletas > 0 else 0

            # Tarjeta de confirmación principal con el mismo estilo de tarjeta oscura
            st.markdown(f"""
            <div class="status-card">
                <p class="status-card-title">Carga de Datos Verificada</p>
                <p class="status-card-value">Montos cargados con éxito: {formato_clp(total_ventas)}</p>
                <p class="status-card-sub">Total neto procesado en <b>{f'{total_boletas:,}'.replace(',', '.')}</b> boletas a través de <b>{len(resumen_archivos)}</b> archivo(s).</p>
            </div>
            """, unsafe_allow_html=True)

            # Tarjetas KPI
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Monto Total (Neto)", formato_clp(total_ventas))
            kpi2.metric("Boletas Emitidas", f"{total_boletas:,}".replace(",", "."))
            kpi3.metric("Boleta Promedio", formato_clp(ticket_promedio))

            st.write("")

            tab_resumen, tab_datos = st.tabs(["Analisis y Recomendaciones", "Base de Datos Detallada"])

            with tab_resumen:
                ventas_mes = df_filtrado.groupby(["Año_Mes", "Mes_Num", "Mes_Nombre"], as_index=False).agg(
                    Monto_Neto=(col_neto, "sum"),
                    Boletas=(col_boleta, "count")
                ).sort_values("Año_Mes")

                if not ventas_mes.empty:
                    ventas_mes["Boleta_Promedio"] = ventas_mes["Monto_Neto"] / ventas_mes["Boletas"]
                    ventas_mes["Participacion_%"] = (ventas_mes["Monto_Neto"] / total_ventas) * 100
                    promedio_val = ventas_mes["Monto_Neto"].mean()

                    col_g1, col_g2 = st.columns([3, 2])

                    with col_g1:
                        fig_bar = px.bar(
                            ventas_mes,
                            x="Año_Mes",
                            y="Monto_Neto",
                            title="Evolución Mensual de Ventas vs. Promedio",
                            color_discrete_sequence=["#38BDF8"],
                            labels={"Año_Mes": "Período", "Monto_Neto": "Monto Neto ($)"},
                            text_auto=".2s"
                        )
                        fig_bar.add_hline(
                            y=promedio_val,
                            line_dash="dash",
                            line_color="#34D399",
                            annotation_text=f"Promedio: {formato_clp(promedio_val)}",
                            annotation_position="top left",
                            annotation_font_color="#34D399"
                        )
                        fig_bar.update_layout(
                            template="plotly_dark",
                            title_font=dict(color="#FFFFFF", size=17),
                            font=dict(color="#F1F5F9"),
                            plot_bgcolor="#1E293B",
                            paper_bgcolor="#1E293B",
                            xaxis=dict(showgrid=False, color="#E2E8F0"),
                            yaxis=dict(showgrid=True, gridcolor="#334155", color="#E2E8F0"),
                            hovermode="x unified",
                            margin=dict(t=55, b=25, l=25, r=25)
                        )
                        st.plotly_chart(fig_bar, use_container_width=True)

                    with col_g2:
                        ventas_por_nombre_mes = df_filtrado.groupby("Mes_Nombre", as_index=False)[col_neto].sum()
                        fig_pie = px.pie(
                            ventas_por_nombre_mes,
                            names="Mes_Nombre",
                            values=col_neto,
                            title="Participación Porcentual por Mes",
                            hole=0.45,
                            color_discrete_sequence=px.colors.qualitative.Pastel
                        )
                        fig_pie.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            textfont=dict(color="#0F172A", size=12)
                        )
                        fig_pie.update_layout(
                            template="plotly_dark",
                            title_font=dict(color="#FFFFFF", size=17),
                            font=dict(color="#F1F5F9"),
                            plot_bgcolor="#1E293B",
                            paper_bgcolor="#1E293B",
                            showlegend=False,
                            margin=dict(t=55, b=25, l=25, r=25)
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)

                    st.subheader("Tabla Comparativa de Rendimiento por Período")
                    tabla_comp = ventas_mes[["Año_Mes", "Mes_Nombre", "Boletas", "Monto_Neto", "Boleta_Promedio", "Participacion_%"]].copy()
                    tabla_comp.columns = ["Período", "Mes", "Boletas Emitidas", "Monto Total (Neto)", "Boleta Promedio", "Participación (%)"]

                    tabla_mostrar = tabla_comp.copy()
                    tabla_mostrar["Monto Total (Neto)"] = tabla_mostrar["Monto Total (Neto)"].apply(formato_clp)
                    tabla_mostrar["Boleta Promedio"] = tabla_mostrar["Boleta Promedio"].apply(formato_clp)
                    tabla_mostrar["Participación (%)"] = tabla_mostrar["Participación (%)"].apply(lambda x: f"{x:.1f}%")

                    st.dataframe(tabla_mostrar, use_container_width=True, hide_index=True)

                    st.divider()

                    st.subheader("Recomendaciones Estratégicas de Inversión Publicitaria")
                    ia = generar_analisis_ia_proyectivo(df_filtrado, col_neto, col_boleta)

                    col_ia1, col_ia2 = st.columns(2)

                    with col_ia1:
                        st.markdown(f"""
                        <div class="insight-box">
                            <h4>1. Cuándo Invertir (Proyección Próximo Ciclo)</h4>
                            <p>El comportamiento histórico demuestra que <b>{ia['mejor_mes_nombre']}</b> es el mes más fuerte del año, concentrando el <b>{ia['participacion_pct']:.1f}%</b> de las ventas totales (<b>{formato_clp(ia['monto_mejor'])}</b> netos en <b>{ia['boletas_mejor']}</b> boletas, superando en un <b>{ia['pct_sobre_promedio']:.1f}%</b> el promedio mensual).</p>
                            <p><b>Estrategia recomendada:</b> Iniciar la inversión fuerte en publicidad durante <b>{ia['mes_previo_nombre']} de {ia['anio_objetivo_previo']}</b> (fase de calentamiento y captación de leads) para maximizar las conversiones durante el pico proyectado de <b>{ia['mejor_mes_nombre']} de {ia['anio_objetivo_pico']}</b>.</p>
                        </div>
                        """, unsafe_allow_html=True)

                    with col_ia2:
                        st.markdown(f"""
                        <div class="insight-box insight-box-green">
                            <h4>2. Por Qué Invertir Así y Optimización de Presupuesto</h4>
                            <p><b>Justificación comercial:</b> En temporada alta ({ia['top_3']}) la intención de compra orgánica crece y la boleta promedio alcanza los <b>{formato_clp(ia['ticket_mejor'])}</b>, lo que garantiza un mayor retorno por cada peso invertido en anuncios (ROAS).</p>
                            <p><b>Contención en meses bajos:</b> Durante <b>{ia['peor_mes_nombre']}</b> (<b>{formato_clp(ia['monto_peor'])}</b>), se recomienda evitar pauta masiva de adquisición fría y redirigir ese presupuesto hacia campañas de remarketing o fidelización de cartera actual.</p>
                        </div>
                        """, unsafe_allow_html=True)

            with tab_datos:
                st.dataframe(
                    df_filtrado.drop(columns=["Fecha_Proc", "Mes_Num"], errors="ignore"),
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.error("No se encontraron las columnas necesarias (Fecha y Monto Neto) en el archivo.")
else:
    st.info("Sube uno o más archivos de ventas desde el menú lateral izquierdo para iniciar el análisis.")