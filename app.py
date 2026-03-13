import io
from datetime import date
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit_authenticator as stauth

# ---------- Configuration ----------
st.set_page_config(
    page_title="PMO Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
DATA_FILE = Path("proyectos.xlsx")
DATE_COLUMNS = ["Fecha Inicio", "Fecha fin Planificada", "Fecha de Cierre"]
TODAY = pd.Timestamp.today().normalize()
REQUIRED_COLUMNS = [
    "Title",
    "Prioridad",
    "Numero de OP",
    "Cliente",
    "Torre",
    "Especialitas",
    "Fecha de Cierre",
    "Fecha fin Planificada",
    "Estado",
    "Fecha Inicio",
    "Avance Planificado",
    "Avance Real",
]

# Map column synonyms (normalized) to canonical names
COLUMN_ALIASES = {
    "Title": ["title", "titulo", "título", "resumen", "proyecto", "nombre", "nombreproyecto", "descripcion", "descripción"],
    "Prioridad": ["prioridad", "priority"],
    "Numero de OP": ["numeroop", "nroop", "n_op", "op", "orden", "ordenproduccion", "ordenproducción"],
    "Cliente": ["cliente", "client", "customer"],
    "Torre": ["torre", "tower"],
    "Especialitas": ["especialitas", "especialistas", "especialidad", "especialidades", "disciplina"],
    "Fecha de Cierre": ["fechadecierre", "fechacierre", "cierre", "cierrefecha"],
    "Fecha fin Planificada": ["fechafinplanificada", "fechafin", "finplanificada", "planfin", "fechafinplaneada"],
    "Estado": ["estado", "status", "etapa", "fase", "situacion", "situación"],
    "Fecha Inicio": ["fechainicio", "inicio", "start", "fechaini"],
    "Avance Planificado": ["avanceplanificado", "avanceplan", "plan"],
    "Avance Real": ["avancereal", "real", "avance"],
}

STATE_ALIASES = {
    "inicio": "Inicio",
    "en ejecucion": "En ejecución",
    "en ejecución": "En ejecución",
    "ejecucion": "En ejecución",
    "ejecución": "En ejecución",
    "ejec": "En ejecución",
    "enproceso": "En ejecución",
    "proceso": "En ejecución",
    "enprogreso": "En ejecución",
    "progreso": "En ejecución",
    "en curso": "En ejecución",
    "stand by": "Stand by",
    "standby": "Stand by",
    "on hold": "Stand by",
    "pausa": "Stand by",
    "pausado": "Stand by",
    "finalizado": "Finalizado",
    "cerrado": "Finalizado",
    "completado": "Finalizado",
    "finalizado": "Finalizado",
}

PRIORIDAD_ALIASES = {
    "critica": "Crítica",
    "crítica": "Crítica",
    "alta": "Alta",
    "media": "Media",
    "baja": "Baja",
    "high": "Alta",
    "medium": "Media",
    "low": "Baja",
    "critical": "Crítica",
    "urgent": "Crítica",
}

# ---------- Authentication ----------
# streamlit-authenticator v0.4+ uses Hasher().hash() per password
hasher = stauth.Hasher()
hashed_passwords = hasher.hash_list(["admin123", "user123"])
credentials = {
    "usernames": {
        "admin": {
            "name": "Administrador",
            "password": hashed_passwords[0],
        },
        "usuario": {
            "name": "Usuario",
            "password": hashed_passwords[1],
        },
    }
}
authenticator = stauth.Authenticate(
    credentials,
    "pmo_dashboard",
    "auth",
    cookie_expiry_days=1,
)

login_fields = {"Form name": "Iniciar sesión", "Username": "Usuario", "Password": "Contraseña"}
authenticator.login(location="main", fields=login_fields, key="login", clear_on_submit=True)
auth_status = st.session_state.get("authentication_status")
name = st.session_state.get("name")
username = st.session_state.get("username")

if auth_status is False:
    st.error("Usuario o contraseña incorrectos")
elif auth_status is None:
    st.warning("Ingresa tus credenciales")
    st.stop()
else:
    authenticator.logout("Cerrar sesión", "sidebar")

# ---------- Data helpers ----------

def build_sample_data() -> pd.DataFrame:
    """Return empty frame with required columns when no Excel is present."""
    empty = pd.DataFrame(columns=REQUIRED_COLUMNS)
    for col in DATE_COLUMNS:
        if col in empty.columns:
            empty[col] = pd.to_datetime(empty[col])
    for col in ["Avance Planificado", "Avance Real"]:
        if col in empty.columns:
            empty[col] = pd.to_numeric(empty[col])
    return empty


def coerce_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ["Avance Planificado", "Avance Real"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _normalize_key(name: str) -> str:
    """Normalize a column name for matching: lowercase, strip spaces, remove accents."""
    normalized = unicodedata.normalize("NFKD", str(name))
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return "".join(ch for ch in ascii_only.lower() if ch.isalnum())


def harmonize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename columns using known aliases so charts/tables align even with varied headers."""
    rename_map = {}
    reverse_index = {alias: canonical for canonical, aliases in COLUMN_ALIASES.items() for alias in aliases}
    for col in df.columns:
        key = _normalize_key(col)
        if key in reverse_index:
            rename_map[col] = reverse_index[key]
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def normalize_estado_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Estado" not in df.columns:
        return df
    df = df.copy()
    def _norm(value: str) -> str:
        key = _normalize_key(value)
        if key in STATE_ALIASES:
            return STATE_ALIASES[key]
        if "ejec" in key or "proce" in key or "progr" in key:
            return "En ejecución"
        if "inic" in key or key == "start":
            return "Inicio"
        if "stand" in key or "hold" in key or "paus" in key:
            return "Stand by"
        if "final" in key or "cierre" in key or "cerr" in key:
            return "Finalizado"
        return value.strip().title()

    df["Estado"] = df["Estado"].astype(str).apply(_norm)
    return df


def normalize_prioridad_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Prioridad" not in df.columns:
        return df
    df = df.copy()
    def _norm(value: str) -> str:
        key = _normalize_key(value)
        if key in PRIORIDAD_ALIASES:
            return PRIORIDAD_ALIASES[key]
        if "crit" in key or "urg" in key:
            return "Crítica"
        if "alt" in key or "high" in key:
            return "Alta"
        if "med" in key:
            return "Media"
        if "baj" in key or "low" in key:
            return "Baja"
        return value.strip().title()

    df["Prioridad"] = df["Prioridad"].astype(str).apply(_norm)
    return df


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Guarantee all required columns exist to avoid KeyError."""
    df = df.copy()
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def load_data(file: Path) -> pd.DataFrame:
    if file.exists():
        try:
            df = pd.read_excel(file)
        except Exception as exc:  # pragma: no cover - defensive
            st.error(f"No se pudo leer el archivo: {exc}")
            df = build_sample_data()
    else:
        df = build_sample_data()

    df = harmonize_columns(df)
    df = normalize_estado_column(df)
    df = normalize_prioridad_column(df)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing and file.exists():
        st.warning(
            "El Excel no trae todas las columnas esperadas. Se rellenarán vacías: "
            + ", ".join(missing)
        )

    return coerce_types(ensure_columns(df))


def save_data(df: pd.DataFrame, file: Path) -> None:
    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df.copy()
    sidebar = st.sidebar
    sidebar.header("Filtros")

    cliente = sidebar.multiselect(
        "Cliente",
        options=sorted(df["Cliente"].dropna().unique()) if "Cliente" in df else [],
        default=None,
    )
    estado = sidebar.multiselect(
        "Estado",
        options=sorted(df["Estado"].dropna().unique()) if "Estado" in df else [],
        default=None,
    )
    prioridad = sidebar.multiselect(
        "Prioridad",
        options=sorted(df["Prioridad"].dropna().unique()) if "Prioridad" in df else [],
        default=None,
    )
    torre = sidebar.multiselect(
        "Torre",
        options=sorted(df["Torre"].dropna().unique()) if "Torre" in df else [],
        default=None,
    )
    especial = sidebar.multiselect(
        "Especialitas",
        options=sorted(df["Especialitas"].dropna().unique()) if "Especialitas" in df else [],
        default=None,
    )

    if "Fecha Inicio" in df and "Fecha fin Planificada" in df and not df.empty:
        min_date = pd.to_datetime(df["Fecha Inicio"].min()).date()
        max_date = pd.to_datetime(df["Fecha fin Planificada"].max()).date()
        date_range = sidebar.date_input("Rango de fechas (Inicio)", (min_date, max_date))
    else:
        date_range = None

    if cliente:
        filtered = filtered[filtered["Cliente"].isin(cliente)]
    if estado:
        filtered = filtered[filtered["Estado"].isin(estado)]
    if prioridad:
        filtered = filtered[filtered["Prioridad"].isin(prioridad)]
    if torre:
        filtered = filtered[filtered["Torre"].isin(torre)]
    if especial:
        filtered = filtered[filtered["Especialitas"].isin(especial)]
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        if "Fecha Inicio" in filtered:
            filtered = filtered[
                (filtered["Fecha Inicio"].dt.date >= start) & (filtered["Fecha Inicio"].dt.date <= end)
            ]
    return filtered


def tag_status(df: pd.DataFrame) -> pd.DataFrame:
    """Add helper flags for retraso and proximidad de vencimiento."""
    df = df.copy()
    estado = df["Estado"].str.lower() if "Estado" in df else pd.Series([], dtype=str)
    finalizado = estado == "finalizado"
    fecha_fin = df["Fecha fin Planificada"] if "Fecha fin Planificada" in df else pd.Series([], dtype="datetime64[ns]")
    fecha_cierre = df["Fecha de Cierre"] if "Fecha de Cierre" in df else pd.Series([], dtype="datetime64[ns]")
    df["Retrasado"] = (fecha_fin < TODAY) & ~finalizado if not fecha_fin.empty else False
    df["ProximoVencer"] = (
        (fecha_fin >= TODAY) & (fecha_fin <= TODAY + pd.Timedelta(days=7)) & ~finalizado
        if not fecha_fin.empty
        else False
    )
    df["BrechaAvance"] = (
        (df["Avance Real"] < (df["Avance Planificado"] - 15))
        if ("Avance Real" in df and "Avance Planificado" in df)
        else False
    )
    df["CierreProximo"] = (
        (fecha_cierre >= TODAY) & (fecha_cierre <= TODAY + pd.Timedelta(days=5)) & ~finalizado
        if not fecha_cierre.empty
        else False
    )
    return df


def compute_metrics(df: pd.DataFrame) -> dict:
    df = normalize_prioridad_column(normalize_estado_column(df))
    df = tag_status(df)

    if df.empty:
        return {
            "total": 0,
            "finalizados": 0,
            "ejecucion": 0,
            "inicio": 0,
            "avance_prom": 0,
            "retrasados": 0,
            "retrasados_pct": 0,
            "proximos": 0,
            "top_cliente": "N/A",
            "top_torre": "N/A",
            "avance_prioridad_str": "N/A",
        }

    total = len(df)
    estado_series = df["Estado"].str.lower() if "Estado" in df else pd.Series([], dtype=str)
    finalizados = estado_series.eq("finalizado").sum()
    ejecucion = estado_series.str.contains("ejecuci", na=False).sum()
    inicio = estado_series.str.contains("inicio", na=False).sum()
    avance_prom = df["Avance Real"].mean() if "Avance Real" in df else 0

    retrasados = df["Retrasado"].sum() if "Retrasado" in df else 0
    retrasados_pct = (retrasados / total * 100) if total else 0
    proximos = df["ProximoVencer"].sum() if "ProximoVencer" in df else 0
    top_cliente = df["Cliente"].mode().iloc[0] if "Cliente" in df and not df["Cliente"].dropna().empty else "N/A"
    top_torre = df["Torre"].mode().iloc[0] if "Torre" in df and not df["Torre"].dropna().empty else "N/A"
    avance_prioridad = (
        df.groupby("Prioridad")["Avance Real"].mean().round(1)
        if "Prioridad" in df and "Avance Real" in df
        else pd.Series(dtype=float)
    )
    avance_prioridad_str = " | ".join([f"{p}: {v:.1f}%" for p, v in avance_prioridad.items()]) if not avance_prioridad.empty else "N/A"

    return {
        "total": total,
        "finalizados": finalizados,
        "ejecucion": ejecucion,
        "inicio": inicio,
        "avance_prom": avance_prom,
        "retrasados": retrasados,
        "retrasados_pct": retrasados_pct,
        "proximos": proximos,
        "top_cliente": top_cliente,
        "top_torre": top_torre,
        "avance_prioridad_str": avance_prioridad_str,
    }


def build_alerts(df: pd.DataFrame) -> pd.DataFrame:
    df = tag_status(normalize_prioridad_column(normalize_estado_column(df)))
    alerts = []
    for _, row in df.iterrows():
        if row.get("Retrasado"):
            alerts.append({
                "Proyecto": row.get("Title"),
                "Alerta": "Fin planificado vencido y no finalizado",
                "Fecha fin": row.get("Fecha fin Planificada").date() if pd.notnull(row.get("Fecha fin Planificada")) else "-",
            })
        if row.get("BrechaAvance"):
            alerts.append({
                "Proyecto": row.get("Title"),
                "Alerta": "Avance real muy por debajo del plan",
                "Brecha (%)": f"{row.get('Avance Planificado', 0) - row.get('Avance Real', 0):.0f}",
            })
        if row.get("CierreProximo"):
            alerts.append({
                "Proyecto": row.get("Title"),
                "Alerta": "Fecha de cierre próxima (<=5 días)",
                "Fecha fin": row.get("Fecha de Cierre").date() if pd.notnull(row.get("Fecha de Cierre")) else "-",
            })
    return pd.DataFrame(alerts)


def kpi_cards(df: pd.DataFrame) -> None:
    m = compute_metrics(df)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de proyectos", f"{m['total']}")
    c2.metric("Proyectos finalizados", f"{m['finalizados']}")
    c3.metric("En ejecución", f"{m['ejecucion']}")
    c4.metric("En inicio", f"{m['inicio']}")
    c5.metric("Promedio avance real", f"{m['avance_prom']:.1f}%")

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("% retrasados", f"{m['retrasados_pct']:.1f}%", help=f"Retrasados: {m['retrasados']}")
    c7.metric("Top cliente", m["top_cliente"])
    c8.metric("Top torre", m["top_torre"])
    c9.metric("Próximos a vencer (7d)", f"{m['proximos']}")
    c10.metric("Avance prom. por prioridad", m["avance_prioridad_str"])
    st.caption(
        "Tarjetas: volumen total, completados, en curso, en arranque, avance promedio, retraso %, top cliente/torre, proximidad a vencimiento (7d) y avance medio por prioridad."
    )


def plot_section(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("No hay datos para graficar.")
        return

    # Normaliza para que los gráficos usen valores alineados
    df_plot = normalize_prioridad_column(normalize_estado_column(df))

    col1, col2, col3 = st.columns(3)
    with col1:
        if "Prioridad" in df_plot:
            fig = px.pie(df_plot, names="Prioridad", title="Proyectos por prioridad", hole=0.4,
                         category_orders={"Prioridad": ["Crítica", "Alta", "Media", "Baja"]})
            st.plotly_chart(fig, width="stretch")
            st.caption("Prioridades: porcentaje de proyectos por nivel (Crítica/Alta/Media/Baja).")
        else:
            st.info("Falta columna Prioridad para este gráfico.")
    with col2:
        if "Especialitas" in df_plot:
            fig = px.bar(df_plot, x="Especialitas", color="Especialitas", title="Proyectos por especialidad")
            st.plotly_chart(fig, width="stretch")
            st.caption("Especialidades: cuántos proyectos hay por disciplina.")
        else:
            st.info("Falta columna Especialitas para este gráfico.")
    with col3:
        if "Cliente" in df_plot:
            fig = px.bar(df_plot, x="Cliente", color="Cliente", title="Proyectos por cliente")
            st.plotly_chart(fig, width="stretch")
            st.caption("Clientes: quién concentra más proyectos en el filtro actual.")
        else:
            st.info("Falta columna Cliente para este gráfico.")

    col4, col5 = st.columns(2)
    with col4:
        if "Estado" in df_plot and "Avance Real" in df_plot:
            avg_state = df_plot.groupby("Estado")["Avance Real"].mean().reset_index()
            fig = px.bar(avg_state, x="Estado", y="Avance Real", color="Estado", title="Avance real promedio por estado")
            st.plotly_chart(fig, width="stretch")
            st.caption("Desempeño: avance real promedio según estado (Inicio, En ejecución, Finalizado, Stand by si aplica).")
        else:
            st.info("Faltan columnas Estado o Avance Real para este gráfico.")
    with col5:
        if "Fecha Inicio" in df_plot:
            timeline = (
                df_plot.assign(Fecha=df_plot["Fecha Inicio"].dt.to_period("M").dt.to_timestamp())
                .groupby("Fecha")
                .size()
                .reset_index(name="Proyectos")
            )
            fig = px.line(timeline, x="Fecha", y="Proyectos", markers=True, title="Evolución de proyectos")
            st.plotly_chart(fig, width="stretch")
            st.caption("Tendencia: proyectos que arrancaron por mes, según los filtros aplicados.")
        else:
            st.info("Falta columna Fecha Inicio para la tendencia.")

    st.markdown("---")
    st.subheader("Estado de proyectos")
    c_state = st.columns(1)[0]
    if "Estado" in df_plot:
        estado_order = ["Inicio", "En ejecución", "Stand by", "Finalizado"]
        counts = df_plot.groupby("Estado").size().reindex(estado_order, fill_value=0).reset_index(name="Proyectos")
        fig = px.bar(
            counts,
            x="Estado",
            y="Proyectos",
            color="Estado",
            category_orders={"Estado": estado_order},
            title="Distribución por estado (incluye Stand by)",
        )
        c_state.plotly_chart(fig, width="stretch")
    else:
        c_state.info("Falta columna Estado para este gráfico.")

    st.markdown("---")
    st.subheader("Gráficos adicionales")
    c11, c12 = st.columns(2)
    with c11:
        if "Title" in df_plot and {"Avance Planificado", "Avance Real"}.issubset(df_plot.columns):
            melted = df_plot.melt(id_vars=["Title"], value_vars=["Avance Planificado", "Avance Real"], var_name="Tipo", value_name="Avance")
            fig = px.bar(
                melted,
                x="Title",
                y="Avance",
                color="Tipo",
                barmode="group",
                title="Avance real vs planificado",
            )
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("Faltan columnas Title/Avance Planificado/Avance Real para comparar avances.")
    with c12:
        if "Torre" in df_plot:
            fig = px.bar(df_plot, x="Torre", color="Torre", title="Proyectos por torre")
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("Falta columna Torre para este gráfico.")

    c13, c14 = st.columns(2)
    with c13:
        if "Cliente" in df_plot:
            carga = df_plot.groupby("Cliente").size().reset_index(name="Proyectos")
            fig = px.bar(carga, x="Cliente", y="Proyectos", color="Cliente", title="Carga de trabajo por cliente")
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("Falta columna Cliente para carga de trabajo.")
    with c14:
        if {"Prioridad", "Estado", "Title"}.issubset(df_plot.columns):
            priorities = ["Crítica", "Alta", "Media", "Baja"]
            states = ["Inicio", "En ejecución", "Stand by", "Finalizado"]

            heat = (
                df_plot.dropna(subset=["Prioridad", "Estado"])
                .assign(_cnt=1)
                .pivot_table(index="Prioridad", columns="Estado", values="_cnt", aggfunc="sum", fill_value=0)
                .reindex(index=priorities, columns=states, fill_value=0)
            )

            if heat.to_numpy().sum() == 0:
                st.info("No hay datos para el heatmap con los filtros actuales.")
            else:
                fig = px.imshow(
                    heat,
                    text_auto=True,
                    aspect="auto",
                    title="Heatmap Prioridad vs Estado",
                    labels={"x": "Estado", "y": "Prioridad", "color": "Proyectos"},
                )
                st.plotly_chart(fig, width="stretch")
        else:
            st.info("Faltan columnas Prioridad/Estado/Title para el heatmap.")

    st.subheader("Diagrama de Gantt")
    df_gantt = tag_status(df_plot)
    if not df_gantt.empty:
        df_gantt = df_gantt.copy()
        df_gantt["Fecha Inicio"] = pd.to_datetime(df_gantt["Fecha Inicio"], errors="coerce")
        df_gantt["Fecha fin Planificada"] = pd.to_datetime(df_gantt["Fecha fin Planificada"], errors="coerce")
        df_gantt = df_gantt.dropna(subset=["Fecha Inicio", "Fecha fin Planificada"]).reset_index(drop=True)
        if df_gantt.empty:
            st.info("No hay fechas válidas para el diagrama de Gantt.")
            return

        df_gantt["Duración (días)"] = (df_gantt["Fecha fin Planificada"] - df_gantt["Fecha Inicio"]).dt.days
        df_gantt["Días restantes"] = (df_gantt["Fecha fin Planificada"] - TODAY).dt.days
        df_gantt["Retraso (días)"] = (
            (TODAY - df_gantt["Fecha fin Planificada"]).dt.days.where(df_gantt["Retrasado"], other=0)
        )
        df_gantt["Estado Gantt"] = df_gantt.apply(
            lambda r: "Retrasado" if r["Retrasado"] else ("Próximo a vencer" if r.get("ProximoVencer") else r["Estado"]),
            axis=1,
        )
        ordered_titles = df_gantt.sort_values("Fecha Inicio")["Title"].dropna().unique().tolist() if "Title" in df_gantt else []
        options = ordered_titles
        selected_titles = st.multiselect(
            "Selecciona proyectos para ver en el Gantt",
            options=options,
            default=options,
        )
        gantt_df = df_gantt[df_gantt["Title"].isin(selected_titles)] if selected_titles else df_gantt.head(0)
        if gantt_df.empty:
            st.info("Selecciona al menos un proyecto para ver el Gantt.")
            return
        color_map = {
            "Retrasado": "#dc3545",
            "Próximo a vencer": "#fd7e14",
            "Inicio": "#6c757d",
            "En ejecución": "#0d6efd",
            "Finalizado": "#198754",
        }
        gantt = px.timeline(
            gantt_df,
            x_start="Fecha Inicio",
            x_end="Fecha fin Planificada",
            y="Title",
            color="Estado Gantt",
            hover_data=[
                "Cliente",
                "Prioridad",
                "Avance Real",
                "Avance Planificado",
                "Duración (días)",
                "Días restantes",
                "Retraso (días)",
            ],
            title="Cronograma de proyectos",
            text="Avance Real",
            color_discrete_map=color_map,
        )
        gantt.update_yaxes(autorange="reversed", categoryorder="array", categoryarray=ordered_titles)
        gantt.update_traces(
            texttemplate="%{text}%",
            textposition="inside",
            hovertemplate=(
                "<b>%{y}</b><br>Inicio: %{x_start|%Y-%m-%d}<br>Fin planificado: %{x_end|%Y-%m-%d}"
                "<br>Estado: %{customdata[0]}<br>Prioridad: %{customdata[1]}"
                "<br>Avance real: %{customdata[2]}%<br>Avance planificado: %{customdata[3]}%"
                "<br>Duración: %{customdata[4]} días<br>Días restantes: %{customdata[5]}"
                "<br>Retraso: %{customdata[6]} días"
            ),
        )
        vline_x = TODAY.to_pydatetime()
        gantt.add_shape(
            type="line",
            x0=vline_x,
            x1=vline_x,
            y0=0,
            y1=1,
            xref="x",
            yref="paper",
            line=dict(color="black", dash="dash"),
        )
        gantt.add_annotation(
            x=vline_x,
            y=1,
            xref="x",
            yref="paper",
            text="Hoy",
            showarrow=False,
            yanchor="bottom",
        )
        gantt.update_xaxes(rangeslider_visible=True)
        gantt.update_layout(xaxis_title="Fecha", yaxis_title="Proyecto", bargap=0.2)
        st.plotly_chart(gantt, width="stretch")
        st.caption("Cronograma: cada barra es un proyecto desde inicio hasta fin planificado, coloreado por estado.")
    else:
        st.info("No hay datos para el diagrama de Gantt")


def export_buttons(df: pd.DataFrame) -> None:
    alerts = build_alerts(df)
    metrics = compute_metrics(df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Proyectos")
    st.download_button(
        label="Descargar datos filtrados (Excel)",
        data=buffer.getvalue(),
        file_name="proyectos_filtrados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        label="Descargar CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="proyectos_filtrados.csv",
        mime="text/csv",
    )

    dash_buffer = io.BytesIO()
    with pd.ExcelWriter(dash_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Proyectos")
        pd.DataFrame([metrics]).to_excel(writer, index=False, sheet_name="KPIs")
        alerts.to_excel(writer, index=False, sheet_name="Alertas")
    st.download_button(
        label="Exportar dashboard (Excel)",
        data=dash_buffer.getvalue(),
        file_name="dashboard_resumen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def admin_panel(df: pd.DataFrame) -> pd.DataFrame:
    st.subheader("Gestión de proyectos (Admin)")
    tabs = st.tabs(["Crear", "Editar", "Eliminar"])

    with tabs[0]:
        st.markdown("### Nuevo proyecto")
        with st.form("crear"):
            title = st.text_input("Título")
            prioridad = st.selectbox("Prioridad", ["Crítica", "Alta", "Media", "Baja"])
            op = st.text_input("Número de OP")
            cliente = st.text_input("Cliente")
            torre = st.text_input("Torre")
            especial = st.text_input("Especialitas")
            estado = st.selectbox("Estado", ["Inicio", "En ejecución", "Finalizado"])
            f_inicio = st.date_input("Fecha Inicio", value=date.today())
            f_fin = st.date_input("Fecha fin Planificada", value=date.today())
            f_cierre = st.date_input("Fecha de Cierre", value=date.today())
            av_plan = st.number_input("Avance Planificado", min_value=0, max_value=100, value=0)
            av_real = st.number_input("Avance Real", min_value=0, max_value=100, value=0)
            submit = st.form_submit_button("Crear")
        if submit:
            new_row = {
                "Title": title,
                "Prioridad": prioridad,
                "Numero de OP": op,
                "Cliente": cliente,
                "Torre": torre,
                "Especialitas": especial,
                "Fecha de Cierre": pd.to_datetime(f_cierre),
                "Fecha fin Planificada": pd.to_datetime(f_fin),
                "Estado": estado,
                "Fecha Inicio": pd.to_datetime(f_inicio),
                "Avance Planificado": av_plan,
                "Avance Real": av_real,
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df, DATA_FILE)
            st.success("Proyecto creado")
            st.rerun()

    with tabs[1]:
        st.markdown("### Editar proyecto")
        if df.empty:
            st.info("No hay proyectos para editar")
        else:
            options = df["Numero de OP"].astype(str) + " - " + df["Title"].astype(str)
            selection = st.selectbox("Selecciona", options)
            idx = options[options == selection].index[0]
            row = df.loc[idx]
            with st.form("editar"):
                title = st.text_input("Título", value=row["Title"])
                prioridad_options = ["Crítica", "Alta", "Media", "Baja"]
                prioridad = st.selectbox(
                    "Prioridad",
                    prioridad_options,
                    index=prioridad_options.index(row["Prioridad"]) if row["Prioridad"] in prioridad_options else 1,
                )
                op = st.text_input("Número de OP", value=str(row["Numero de OP"]))
                cliente = st.text_input("Cliente", value=row["Cliente"])
                torre = st.text_input("Torre", value=row["Torre"])
                especial = st.text_input("Especialitas", value=row["Especialitas"])
                estado = st.selectbox(
                    "Estado",
                    ["Inicio", "En ejecución", "Finalizado"],
                    index=["Inicio", "En ejecución", "Finalizado"].index(row["Estado"]) if row["Estado"] in ["Inicio", "En ejecución", "Finalizado"] else 0,
                )
                f_inicio = st.date_input("Fecha Inicio", value=row["Fecha Inicio"].date() if pd.notnull(row["Fecha Inicio"]) else date.today())
                f_fin = st.date_input("Fecha fin Planificada", value=row["Fecha fin Planificada"].date() if pd.notnull(row["Fecha fin Planificada"]) else date.today())
                f_cierre = st.date_input("Fecha de Cierre", value=row["Fecha de Cierre"].date() if pd.notnull(row["Fecha de Cierre"]) else date.today())
                av_plan = st.number_input(
                    "Avance Planificado", min_value=0, max_value=100, value=int(row["Avance Planificado"] if pd.notnull(row["Avance Planificado"]) else 0)
                )
                av_real = st.number_input(
                    "Avance Real", min_value=0, max_value=100, value=int(row["Avance Real"] if pd.notnull(row["Avance Real"]) else 0)
                )
                submit = st.form_submit_button("Guardar cambios")
            if submit:
                df.loc[idx, :] = [
                    title,
                    prioridad,
                    op,
                    cliente,
                    torre,
                    especial,
                    pd.to_datetime(f_cierre),
                    pd.to_datetime(f_fin),
                    estado,
                    pd.to_datetime(f_inicio),
                    av_plan,
                    av_real,
                ]
                save_data(df, DATA_FILE)
                st.success("Cambios guardados")
                st.rerun()

    with tabs[2]:
        st.markdown("### Eliminar proyecto")
        if df.empty:
            st.info("No hay proyectos para eliminar")
        else:
            options = df["Numero de OP"].astype(str) + " - " + df["Title"].astype(str)
            selection = st.selectbox("Selecciona proyecto", options, key="del")
            idx = options[options == selection].index[0]
            if st.button("Eliminar", type="primary"):
                df = df.drop(index=idx).reset_index(drop=True)
                save_data(df, DATA_FILE)
                st.success("Proyecto eliminado")
                st.rerun()
    return df


# ---------- Main app ----------
st.title("Dashboard de Gestión de Proyectos")
st.caption("Estilo Power BI, ejecutado en Streamlit")
st.markdown(
    """
    **Guía rápida**: usa los filtros a la izquierda para acotar. Las tarjetas resumen volumen y avance.
    Las gráficas muestran distribución (prioridades, especialidades, clientes) y desempeño (avance promedio),
    y el Gantt enseña el calendario planificado por proyecto.
    """
)

st.sidebar.markdown("### Cargar/Actualizar Excel")
upload = st.sidebar.file_uploader("Sube proyectos.xlsx", type=["xlsx"])
if upload:
    DATA_FILE.write_bytes(upload.getbuffer())
    st.sidebar.success("Archivo cargado y datos actualizados")
if st.sidebar.button("Recargar datos", type="secondary"):
    st.rerun()

projects_df = load_data(DATA_FILE)
filtered_df = apply_filters(projects_df)

kpi_cards(filtered_df)
plot_section(filtered_df)

st.subheader("Tabla de proyectos")
st.dataframe(
    filtered_df.sort_values(by="Fecha Inicio", ascending=False),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Exportar")
export_buttons(filtered_df)

if username == "admin":
    projects_df = admin_panel(projects_df)
else:
    st.info("Modo lectura: autenticado como usuario")

# Persist filtered data after any admin change
if username == "admin":
    pass
