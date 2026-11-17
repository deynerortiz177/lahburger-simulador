"""
La H Burger · ¿Cuánto podemos pedir prestado?
Simulador Monte Carlo de flujo de caja y capacidad de endeudamiento.
Ejecutar:  streamlit run app.py
"""
from pathlib import Path
from dataclasses import asdict
import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from modelo import (cargar_excel, resumir, Supuestos, Destino, Credito, DESTINOS_POR_DEFECTO, simular,
                    monto_maximo, comparar_plazos, comparar_destinos, sensibilidad, tabla_amortizacion,
                    requisitos_automaticos, REQUISITOS_DOCUMENTALES, modalidad_credito, tasa_ea_a_mensual, caja_historica)

BASE = Path(__file__).parent
PLANTILLA = BASE / "datos" / "Plantilla_Registro_LaHBurger.xlsx"
SIMULADOS = BASE / "datos" / "Datos_Simulados_LaHBurger.xlsx"
LOGO = BASE / "datos" / "logo.png"

# ----------------------------------------------------------------- identidad de marca
C = {"negro": "#0D1014", "panel": "#171A1F", "dorado": "#D8AC55", "dorado_osc": "#B07E36", "crema": "#FCF0CF",
     "vino": "#894B33", "rojo": "#E05A47", "verde": "#7FB77E", "gris": "#8A8F98"}
st.set_page_config(page_title="La H Burger · Simulador de préstamo", page_icon=str(LOGO), layout="wide")
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Anton&family=Inter:wght@400;600&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  h1, h2, h3 {{ font-family: 'Anton', sans-serif !important; letter-spacing: .5px; color: {C['dorado']} !important; text-transform: uppercase; }}
  [data-testid="stMetricValue"] {{ color: {C['crema']}; font-family: 'Anton', sans-serif; letter-spacing: .5px; font-size: 1.55rem; }}
  [data-testid="stMetricLabel"] {{ color: {C['dorado']}; }}
  [data-testid="stSidebar"] {{ border-right: 1px solid {C['dorado_osc']}55; }}
  .stTabs [data-baseweb="tab"] {{ font-family: 'Anton', sans-serif; letter-spacing: .5px; font-size: 1.05rem; }}
  .marca {{ font-family: 'Anton', sans-serif; color: {C['dorado']}; font-size: 1.6rem; letter-spacing: 1px; margin: 0; }}
  .lema {{ color: {C['crema']}; opacity: .8; margin-top: -4px; }}
</style>""", unsafe_allow_html=True)

PLOT = dict(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter", color=C["crema"]), colorway=[C["dorado"], C["crema"], C["vino"], C["verde"], C["gris"]])


def estilo(fig, **kw):
    fig.update_layout(**PLOT, margin=dict(t=50, l=10, r=10, b=10), **kw)
    return fig


def pesos(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    s = f"{abs(x):,.0f}".replace(",", ".")
    return f"-${s}" if x < 0 else f"${s}"


def pct(x, d=1) -> str:
    return f"{x * 100:.{d}f}%".replace(".", ",")


# ----------------------------------------------------------------- barra lateral: datos
with st.sidebar:
    st.image(str(LOGO), width=110)
    st.markdown('<p class="marca">LA H BURGER</p><p class="lema">¿Vamos para La H o qué? · Campo Valdés</p>', unsafe_allow_html=True)
    st.download_button("⬇️ Plantilla de Excel", PLANTILLA.read_bytes(), file_name=PLANTILLA.name)
    archivo = st.file_uploader("Subir Excel con los registros", type=["xlsx"])
    usar_sim = st.toggle("Usar datos simulados", value=archivo is None, help="Datos simulados con las reglas del negocio. Reemplácelos por los reales.")


@st.cache_data(show_spinner=False)
def _cargar(b: bytes):
    return cargar_excel(io.BytesIO(b))


if archivo is not None and not usar_sim:
    fuente, contenido = archivo.name, archivo.getvalue()
elif usar_sim:
    fuente, contenido = "Datos simulados", SIMULADOS.read_bytes()
else:
    st.title("¿Cuánto podemos pedir prestado?")
    st.info("Suba su Excel o active los datos simulados en la barra lateral.")
    st.stop()

try:
    datos, avisos = _cargar(contenido)
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()
for a in avisos:
    st.sidebar.warning(a)
if datos["Ventas"].empty:
    st.error("El archivo no tiene ventas registradas.")
    st.stop()
PAR = datos["Parametros"]
gp = lambda k, d: (PAR.get(k) if PAR.get(k) is not None and not pd.isna(PAR.get(k)) else d)

# ----------------------------------------------------------------- barra lateral: supuestos
with st.sidebar:
    st.header("Supuestos")
    with st.expander("Periodo base", expanded=True):
        fmin = datos["Ventas"]["Fecha"].min().date(); fmax = datos["Ventas"]["Fecha"].max().date()
        fb_def = pd.to_datetime(gp("Fecha_inicio_base", fmin)).date()
        fib = st.date_input("Simular con los datos desde", value=min(max(fb_def, fmin), fmax), min_value=fmin, max_value=fmax,
                            help="Use solo el periodo en que el negocio opera como hoy (mismos productos y días de apertura).")
R = resumir(datos, pd.Timestamp(fib))
with st.sidebar:
    with st.expander("Ventas", expanded=True):
        st.caption(f"{R.stats_diarias['n_semanas_base']} semanas en el periodo base · venta semanal promedio {pesos(R.stats_diarias['venta_semanal_media'])}")
        factor = st.slider("Escenario de demanda (× ventas actuales)", 0.5, 1.5, 1.0, 0.05)
        choque = st.slider("Variación adicional entre meses (σ)", 0.0, 0.30, float(gp("Choque_mensual", 0.05)), 0.01,
                           help="Clima, temporada, competencia. Con pocos meses de historia es un supuesto.")
    with st.expander("Costos", expanded=True):
        cv_dist = st.radio("Distribución del % de materia prima", ["Triangular", "Uniforme"], horizontal=True)
        c1, c2, c3 = st.columns(3)
        cv_min = c1.number_input("Mín %", 0.0, 100.0, round(R.cv_hist["min"] * 100, 1)) / 100
        cv_moda = c2.number_input("Moda %", 0.0, 100.0, round(R.cv_hist["moda"] * 100, 1), disabled=cv_dist == "Uniforme") / 100
        cv_max = c3.number_input("Máx %", 0.0, 100.0, round(R.cv_hist["max"] * 100, 1)) / 100
        cf = st.number_input("Costos fijos mensuales", 0.0, value=float(R.costos_fijos_mes), step=50000.0, format="%.0f")
        r1, r2 = st.columns(2)
        ret_min = r1.number_input("Retiro mín.", 0.0, value=float(gp("Retiro_min", 0)), step=100000.0, format="%.0f")
        ret_max = r2.number_input("Retiro máx.", 0.0, value=float(gp("Retiro_max", 0)), step=100000.0, format="%.0f")
        CAJA_H = caja_historica(datos)
        caja_def = gp("Caja_actual", None)
        caja_def = float(caja_def) if isinstance(caja_def, (int, float, np.floating)) else float(CAJA_H["Valor"].iloc[-1])
        caja0 = st.number_input("Caja disponible hoy", value=caja_def, step=100000.0, format="%.0f",
                                help="Por defecto: flujo acumulado desde la apertura (pestaña Datos). Reemplácelo por el conteo real si lo tiene.")
    with st.expander("Criterios y simulación"):
        umbral = st.slider("Umbral: P(F > 0) ≥", 0.80, 0.99, float(gp("Umbral_aprobacion", 0.96)), 0.01)
        dscr_min = st.number_input("Cobertura mínima (DSCR)", 1.0, 3.0, float(gp("DSCR_min", 1.2)), 0.05)
        alfa = st.slider("Riesgo aceptado (α)", 0.01, 0.20, float(gp("Alfa_riesgo", 0.05)), 0.01)
        horizonte = st.number_input("Horizonte (meses)", 6, 60, 24)
        n_sim = st.select_slider("Simulaciones", [1000, 5000, 10000, 20000], int(gp("N_simulaciones", 10000)))
        semilla = st.number_input("Semilla", 0, 99999, int(gp("Semilla", 2026)))

S = Supuestos(n_sim=n_sim, horizonte=horizonte, factor_demanda=factor, choque_mensual=choque,
              cv_dist="uniforme" if cv_dist == "Uniforme" else "triangular", cv_min=cv_min, cv_moda=cv_moda, cv_max=cv_max,
              costos_fijos=cf, costos_variables_pct=R.costos_variables_pct, retiro_min=ret_min, retiro_max=ret_max,
              caja_inicial=caja0, dscr_min=dscr_min, alfa=alfa, umbral_aprobacion=umbral, semilla=semilla)
res0 = simular(R, S)

# ----------------------------------------------------------------- encabezado
st.title("¿Cuánto podemos pedir prestado?")
st.caption(f"Fuente: **{fuente}** · ventas del {R.ventas_diarias['Fecha'].min():%d/%m/%Y} al {R.ventas_diarias['Fecha'].max():%d/%m/%Y} · "
           f"periodo base desde {R.fecha_inicio_base:%d/%m/%Y}")
if fuente == "Datos simulados":
    st.warning("Está viendo **datos simulados** con las reglas del negocio. Reemplácelos por los registros reales antes de presentar al banco.")
tabs = st.tabs(["📊 Datos", "📈 Flujo de caja proyectado", "🏦 Préstamo", "✅ Requisitos", "📘 Metodología"])

# ================================================================= 1. DATOS
with tabs[0]:
    s = R.stats_diarias
    k = st.columns(5)
    k[0].metric("Venta semanal promedio", pesos(s["venta_semanal_media"]), f"σ {pesos(s['venta_semanal_desv'])}", delta_color="off")
    k[1].metric("Venta promedio por jornada", pesos(s["media"]))
    k[2].metric("% materia prima / ventas", pct(R.cv_hist["media"]))
    k[3].metric("Costos por jornada / semana", pesos(R.semanas_base["Costos_jornada"].mean()))
    k[4].metric("Costos fijos / mes", pesos(R.costos_fijos_mes))

    sem = R.semanas.copy()
    fig = go.Figure()
    fig.add_bar(x=sem["Semana"], y=sem["Ventas"], marker_color=np.where(sem["En_base"], C["dorado"], C["gris"]), name="Ventas")
    fig.add_vline(x=R.fecha_inicio_base, line_dash="dash", line_color=C["crema"])
    fig.add_annotation(x=R.fecha_inicio_base, y=sem["Ventas"].max(), text="Inicio del periodo base", showarrow=False, xanchor="left", yanchor="bottom")
    estilo(fig, title="Ventas por semana (en dorado, las que alimentan la simulación)", yaxis_title="$", showlegend=False)
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    fig = px.bar(R.por_dia_semana, x="Dia_semana", y="Unidades_prom", title="Hamburguesas promedio por día de apertura (periodo base)",
                 text=R.por_dia_semana["Unidades_prom"].round(0))
    fig.update_traces(marker_color=C["dorado"]); estilo(fig, xaxis_title=None, yaxis_title="Unidades")
    c1.plotly_chart(fig, width="stretch")
    fig = px.bar(R.por_producto, x="Producto", y="Ventas", title="Ventas por hamburguesa (todo el historial)",
                 text=R.por_producto["Participacion"].map(pct))
    fig.update_traces(marker_color=C["vino"]); estilo(fig, xaxis_title=None, yaxis_title="$")
    c2.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    fig = px.pie(R.mp_por_categoria, names="Categoria", values="Gasto", hole=0.55, title="Compras de materia prima por categoría",
                 color_discrete_sequence=[C["dorado"], C["vino"], C["crema"], C["dorado_osc"], C["gris"], C["verde"], "#5C4A3A", "#C9B48A", "#3E4450", "#A0522D", "#6B7280", "#E8D9A8"])
    c1.plotly_chart(estilo(fig), width="stretch")
    c2.markdown("**Estadística descriptiva de la venta por jornada (periodo base)**")
    c2.dataframe(pd.DataFrame({"Estadístico": ["Jornadas", "Media", "Mediana", "Desviación", "Mínimo", "P25", "P75", "Máximo", "Coef. de variación", "Asimetría", "Curtosis"],
                               "Valor": [str(s["n_dias"]), pesos(s["media"]), pesos(s["mediana"]), pesos(s["desv"]), pesos(s["min"]), pesos(s["p25"]),
                                         pesos(s["p75"]), pesos(s["max"]), pct(s["cv_coef"]), f"{s['asimetria']:.2f}", f"{s['curtosis']:.2f}"]}),
                 hide_index=True, width="stretch")
    st.markdown("**Resumen mensual**")
    st.dataframe(R.meses.style.format({c_: pesos for c_ in ["Ventas", "Compras_MP", "Fijo mensual", "Por jornada", "Variable", "F_operativo"]} | {"CV_pct": pct}),
                 hide_index=True, width="stretch")
    st.caption("F_operativo del resumen mensual = ventas − compras de materia prima − costos operativos (antes del retiro del dueño y de las deudas).")
    with st.expander("¿De dónde sale la caja disponible hoy?"):
        st.dataframe(CAJA_H.style.format({"Valor": pesos}), hide_index=True, width="stretch")
        st.caption("Es el dinero que el flujo fue dejando desde la apertura. Si el conteo real de caja es distinto, escríbalo en la barra lateral.")
    with st.expander("Activos y deudas registradas"):
        st.write(f"Depreciación mensual: **{pesos(R.depreciacion_mes)}** (no es salida de caja).")
        st.dataframe(datos["Activos"], hide_index=True)
        st.dataframe(R.deudas.style.format({"Cuota": pesos, "Saldo": pesos}), hide_index=True)

# ================================================================= 2. FLUJO DE CAJA PROYECTADO
with tabs[1]:
    st.subheader("Flujo de caja mensual proyectado, sin préstamo nuevo")
    st.latex(r"F = V\,(1-\%MP) - C_{jornada} - C_{fijos} - \text{Retiro}")
    pf = res0.prob_f_op_positivo()
    k = st.columns(4)
    k[0].metric("P(F > 0)", pct(pf, 2), "Cumple el umbral" if pf >= umbral else "No cumple el umbral", delta_color="normal" if pf >= umbral else "inverse")
    k[1].metric("F mensual promedio", pesos(res0.f_op.mean()))
    k[2].metric("F en el peor 5 % (P5)", pesos(res0.percentiles()["P5"]))
    k[3].metric("P(caja < 0 algún mes)", pct(res0.prob_caja_negativa(), 2))
    (st.success if pf >= umbral else st.error)(
        f"El flujo mensual, después del retiro del dueño, es positivo en el {pct(pf, 2)} de los "
        + f"{n_sim * res0.f_op.shape[1]:,}".replace(",", ".") + f" meses simulados (umbral {pct(umbral, 0)}).")

    pr = res0.proyeccion()
    fig = go.Figure()
    fig.add_scatter(x=pr["Mes"], y=pr["F_operativo_P90"], line=dict(width=0), showlegend=False)
    fig.add_scatter(x=pr["Mes"], y=pr["F_operativo_P10"], fill="tonexty", fillcolor="rgba(216,172,85,0.25)", line=dict(width=0), name="P10–P90")
    fig.add_scatter(x=pr["Mes"], y=pr["F_operativo_P50"], line_color=C["dorado"], name="Mediana")
    fig.add_hline(y=0, line_color=C["rojo"], line_dash="dash")
    estilo(fig, title="Proyección del flujo operativo mensual (abanico de escenarios)", xaxis_title="Mes", yaxis_title="$")
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    x = res0.f_op.ravel()
    fig = px.histogram(x=x, nbins=60, title="Distribución del flujo mensual F")
    fig.update_traces(marker_color=C["dorado"]); fig.add_vline(x=0, line_color=C["rojo"], line_dash="dash")
    fig.add_vline(x=x.mean(), line_color=C["crema"], annotation_text="Promedio")
    c1.plotly_chart(estilo(fig, xaxis_title="$ por mes", yaxis_title="Frecuencia"), width="stretch")
    tor = sensibilidad(res0)
    fig = go.Figure()
    fig.add_bar(y=tor["Variable"], x=tor["Desfavorable"] - tor["Base"], orientation="h", marker_color=C["rojo"], name="−10 % favorable")
    fig.add_bar(y=tor["Variable"], x=tor["Favorable"] - tor["Base"], orientation="h", marker_color=C["verde"], name="+10 % favorable")
    c2.plotly_chart(estilo(fig, barmode="overlay", showlegend=False, title="Sensibilidad: cambio en F si cada variable se mueve 10 %",
                           xaxis_title="Cambio en F promedio ($)"), width="stretch")
    c1, c2 = st.columns([1, 2])
    c1.dataframe(pd.DataFrame({"Percentil": list(res0.percentiles()), "F mensual": [pesos(v) for v in res0.percentiles().values()]}),
                 hide_index=True, width="stretch")
    with c2.expander("Tabla de proyección mensual", expanded=True):
        st.dataframe(pr.style.format({c_: pesos for c_ in pr.columns if c_ not in ("Mes", "P_F_neto_negativo")} | {"P_F_neto_negativo": pct}),
                     hide_index=True, height=320)

# ================================================================= 3. PRÉSTAMO
with tabs[2]:
    st.subheader("1 · ¿Para qué es el préstamo?")
    esc = datos["Escenarios"]
    destinos = [Destino.desde_fila(r) for _, r in esc.iterrows()] if len(esc) else list(DESTINOS_POR_DEFECTO)
    nombres = [d.nombre for d in destinos]
    elegido = st.selectbox("Destino del préstamo", nombres, index=0,
                           help="El banco siempre pregunta el destino. Para el cálculo importa porque cambia ventas y costos.")
    d0 = destinos[nombres.index(elegido)]
    with st.expander("Efecto esperado del destino (supuestos editables)", expanded=not elegido.startswith("Solo")):
        c = st.columns(4)
        inv = c[0].number_input("Inversión necesaria", 0.0, value=float(d0.inversion), step=500000.0, format="%.0f")
        vmin = c[1].number_input("Ventas: cambio mínimo %", -100.0, 300.0, d0.ventas_min * 100, 1.0) / 100
        vmod = c[2].number_input("Ventas: cambio más probable %", -100.0, 300.0, d0.ventas_probable * 100, 1.0) / 100
        vmax = c[3].number_input("Ventas: cambio máximo %", -100.0, 300.0, d0.ventas_max * 100, 1.0) / 100
        c = st.columns(4)
        dcv = c[0].number_input("% materia prima: cambio (puntos)", -20.0, 20.0, d0.cambio_cv_pp * 100, 0.5) / 100
        cfa = c[1].number_input("Costo fijo adicional / mes", 0.0, value=float(d0.costo_fijo_adicional), step=100000.0, format="%.0f")
        dcj = c[2].number_input("Costos por jornada: cambio %", -100.0, 300.0, d0.cambio_costos_jornada * 100, 5.0) / 100
        meses_ef = c[3].number_input("Meses hasta que surte efecto", 0, 24, int(d0.meses_hasta_efecto))
        paga = st.checkbox("El préstamo también cancela las deudas vigentes", value=d0.paga_deudas)
        st.caption("El cambio en ventas se modela como triangular (mín., más probable, máx.) y se sortea una vez por escenario: así la simulación "
                   "incluye la incertidumbre de que la inversión funcione mejor o peor de lo esperado.")
    D = Destino(elegido, inv, vmin, vmod, vmax, dcv, cfa, dcj, meses_ef, paga)

    st.subheader("2 · Condiciones del crédito (las que ofrezca la entidad)")
    c = st.columns(6)
    saldo_d = float(R.deudas["Saldo"].fillna(0).sum()) if paga else 0.0
    com = c[4].number_input("Comisión / FNG al desembolso %", 0.0, 20.0, float(gp("Comision_inicial", 0.03)) * 100, 0.5) / 100
    monto_nec = (D.inversion + saldo_d) / (1 - com) if (D.inversion + saldo_d) > 0 else 0.0
    monto = c[0].number_input("Monto solicitado", 0.0, value=float(np.ceil(monto_nec / 10000) * 10000), step=500000.0, format="%.0f",
                              help="Por defecto: lo necesario para cubrir la inversión (y deudas, si aplica) después de la comisión.")
    tasa = c[1].number_input("Tasa EA %", 0.0, 120.0, float(gp("Tasa_EA", 0.42)) * 100, 0.5) / 100
    plazo = c[2].number_input("Plazo (meses)", 1, 120, int(gp("Plazo_meses", 24)))
    gracia = c[3].number_input("Gracia a capital (meses)", 0, 12, int(gp("Gracia_meses", 0)))
    seg = c[5].number_input("Seguro mensual % del saldo", 0.0, 2.0, float(gp("Seguro_mensual", 0.001)) * 100, 0.05) / 100
    CR = Credito(monto, tasa, plazo, min(gracia, plazo - 1), com, seg)
    st.caption(f"Tasa mensual equivalente: {pct(CR.tasa_mensual, 3)} · Desembolso neto: {pesos(CR.desembolso_neto)} · "
               f"Modalidad: {modalidad_credito(monto, float(gp('SMMLV', 1750905)))}")

    res1 = simular(R, S, D, CR)
    MM = monto_maximo(res1, CR)
    viable = monto <= MM["monto_max"] and res1.prob_dscr_bajo() <= alfa

    st.subheader("3 · Resultado")
    k = st.columns(4)
    k[0].metric("Monto máximo recomendado", pesos(MM["monto_max"]), help=f"Mes más exigente: {MM['mes_critico']}")
    k[1].metric("Pago mensual (después de la gracia)", pesos(float(CR.pagos()[CR.gracia])) if monto > 0 else "–")
    k[2].metric(f"P(DSCR < {dscr_min:.2f})", pct(res1.prob_dscr_bajo(), 2))
    k[3].metric("P(caja < 0 algún mes)", pct(res1.prob_caja_negativa(), 2))
    k = st.columns(4)
    k[0].metric("Costo total del crédito", pesos(CR.costo_total()) if monto > 0 else "–", help="Intereses + seguros + comisión.")
    k[1].metric("Aporte propio necesario", pesos(res1.aporte_propio))
    k[2].metric("F operativo P50 (mes 12)", pesos(float(np.percentile(res1.f_op[:, min(11, res1.f_op.shape[1] - 1)], 50))))
    k[3].metric("Efecto en ventas (promedio simulado)", pct(float(res1.efecto_ventas.mean())))
    if monto == 0:
        st.info(f"Sin monto solicitado: el negocio podría pagar hasta **{pesos(MM['monto_max'])}** a {plazo} meses con {pct(tasa)} EA, "
                f"manteniendo la cobertura ≥ {dscr_min:.2f} en el {pct(1 - alfa, 0)} de los escenarios.")
    elif viable:
        st.success(f"**Viable.** {pesos(monto)} a {plazo} meses está dentro de la capacidad estimada ({pesos(MM['monto_max'])}).")
    else:
        st.error(f"**No recomendable con estos supuestos.** Monto máximo {pesos(MM['monto_max'])} frente a {pesos(monto)} solicitado. "
                 "Pruebe un plazo mayor, un monto menor, otra tasa o revise los efectos del destino.")
    st.caption(f"Regla: en cada mes del plazo, pago del préstamo ≤ P{int(alfa * 100)}(F operativo del mes) / DSCR mínimo − cuotas vigentes.")

    c1, c2 = st.columns(2)
    pr1 = res1.proyeccion()
    fig = go.Figure()
    fig.add_scatter(x=pr1["Mes"], y=pr1["Caja_P90"], line=dict(width=0), showlegend=False)
    fig.add_scatter(x=pr1["Mes"], y=pr1["Caja_P10"], fill="tonexty", fillcolor="rgba(216,172,85,0.25)", line=dict(width=0), name="P10–P90")
    fig.add_scatter(x=pr1["Mes"], y=pr1["Caja_P50"], line_color=C["dorado"], name="Mediana")
    fig.add_hline(y=0, line_color=C["rojo"], line_dash="dash")
    c1.plotly_chart(estilo(fig, title="Caja acumulada con el préstamo", xaxis_title="Mes", yaxis_title="$"), width="stretch")
    dd = res1.dscr(); dd = dd[~np.isnan(dd)]
    if dd.size:
        fig = px.histogram(x=np.clip(dd, np.percentile(dd, .5), np.percentile(dd, 99.5)), nbins=60, title="Cobertura de la deuda (DSCR) por mes simulado")
        fig.update_traces(marker_color=C["vino"]); fig.add_vline(x=dscr_min, line_color=C["rojo"], line_dash="dash", annotation_text=f"Mín. {dscr_min:.2f}")
        c2.plotly_chart(estilo(fig, xaxis_title="DSCR = F operativo / cuotas", yaxis_title="Frecuencia"), width="stretch")

    st.subheader("4 · Comparar plazos")
    Sc = Supuestos(**{**asdict(S), "n_sim": min(n_sim, 5000)})
    cp = comparar_plazos(R, Sc, D, CR)
    st.dataframe(cp.style.format({"Pago mensual del monto solicitado": pesos, "Costo total del crédito": pesos, "Monto máximo": pesos,
                                  "P(DSCR < mín.)": pct, "P(caja < 0 algún mes)": pct}), hide_index=True, width="stretch")
    fig = go.Figure()
    fig.add_bar(x=cp["Plazo (meses)"].astype(str), y=cp["Monto máximo"], marker_color=C["dorado"], name="Monto máximo")
    if monto > 0:
        fig.add_hline(y=monto, line_color=C["crema"], line_dash="dash", annotation_text="Monto solicitado")
    st.plotly_chart(estilo(fig, title="Monto máximo según el plazo", xaxis_title="Plazo (meses)", yaxis_title="$", showlegend=False), width="stretch")

    st.subheader("5 · Comparar destinos")
    cd = comparar_destinos(R, Supuestos(**{**asdict(S), "n_sim": 3000}), [x for x in destinos if x.nombre != "Personalizado"], CR)
    st.dataframe(cd.style.format({"Monto necesario": pesos, "Monto máximo": pesos, "F operativo P50 (mes 12)": pesos,
                                  "P(DSCR < mín.)": pct, "P(caja < 0 algún mes)": pct}), hide_index=True, width="stretch")
    st.caption("Cada destino usa sus efectos de la hoja Escenarios_Prestamo y el monto que necesita, con las mismas condiciones del crédito.")

    with st.expander("Tabla de amortización"):
        if monto > 0:
            am = tabla_amortizacion(monto, CR.tasa_mensual, CR.plazo, CR.gracia, CR.seguro_mensual)
            st.dataframe(am.style.format({c_: pesos for c_ in am.columns if c_ != "Mes"}), hide_index=True, height=300)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        pr1.to_excel(w, "Proyeccion_con_prestamo", index=False)
        res0.proyeccion().to_excel(w, "Proyeccion_sin_prestamo", index=False)
        if monto > 0:
            tabla_amortizacion(monto, CR.tasa_mensual, CR.plazo, CR.gracia, CR.seguro_mensual).to_excel(w, "Amortizacion", index=False)
        cp.to_excel(w, "Comparar_plazos", index=False); cd.to_excel(w, "Comparar_destinos", index=False)
        pd.DataFrame([{"Parametro": k_, "Valor": v_} for k_, v_ in {**asdict(S), **{f"destino_{a}": b for a, b in asdict(D).items()},
                                                                   **{f"credito_{a}": b for a, b in asdict(CR).items()},
                                                                   "monto_maximo": MM["monto_max"]}.items()]).to_excel(w, "Supuestos", index=False)
    st.download_button("⬇️ Descargar resultados (Excel)", out.getvalue(), "resultados_prestamo_LaHBurger.xlsx")

# ================================================================= 4. REQUISITOS
with tabs[3]:
    st.subheader("Requisitos habituales de un microcrédito")
    st.caption("Cada entidad define los suyos: confírmelos con el asesor. Los primeros se verifican con los datos y la simulación.")
    rq = requisitos_automaticos(R, res1, CR, MM, PAR)
    rq["Estado"] = np.where(rq["Cumple"], "✅ Cumple", "⚠️ Revisar")
    st.dataframe(rq[["Estado", "Requisito", "Detalle"]], hide_index=True, width="stretch")
    st.markdown("**Documentos y condiciones que debe reunir el dueño**")
    hechos = 0
    cols = st.columns(2)
    for i, r_ in enumerate(REQUISITOS_DOCUMENTALES):
        hechos += cols[i % 2].checkbox(r_, key=f"rq{i}")
    st.progress(hechos / len(REQUISITOS_DOCUMENTALES), text=f"{hechos} de {len(REQUISITOS_DOCUMENTALES)} listos")
    st.caption("Tasas de usura de noviembre de 2026 (Superfinanciera, vía La República) y SMMLV 2026 en la hoja Parametros; actualícelas cada mes.")

# ================================================================= 5. METODOLOGÍA
with tabs[4]:
    st.subheader("Cómo funciona")
    st.markdown(f"""
**Problema.** La H Burger no puede demostrar su capacidad de pago porque su flujo de caja es incierto y no está medido.
**Solución.** Simular miles de meses posibles con los datos reales y calcular cuánto puede pagar con un riesgo aceptado.

**1. Periodo base.** Solo se usan las semanas desde **{R.fecha_inicio_base:%d/%m/%Y}**, cuando el negocio ya operaba como hoy
(mismos productos y días de apertura). Usar todo el historial subestimaría las ventas actuales.

**2. Ventas y costos por jornada (bootstrap por semanas).** Cada mes simulado se arma con 4 o 5 semanas reales sorteadas con reemplazo
(5 semanas en el ≈ 33 % de los meses, porque un mes tiene 4,33 semanas en promedio). Sortear semanas completas conserva juntos los
jueves, fines de semana y domingos de festivo, y las ventas con sus costos de ayudantes y gas. Luego se aplica el escenario de demanda
y un choque mensual N(1; σ = {choque:.2f}).

**3. Materia prima.** % de compras sobre ventas, medido en ventanas de 4 semanas del periodo base; distribución {cv_dist.lower()}
entre {pct(cv_min)} y {pct(cv_max)}. Se trata en conjunto porque las recetas no están estandarizadas.

**4. Flujo mensual.** F = V(1 − %MP) − costos por jornada − costos fijos − retiro del dueño (uniforme entre {pesos(ret_min)} y {pesos(ret_max)}).
Con préstamo: F neto = F − cuotas vigentes − pago del préstamo (con gracia, intereses y seguro).

**5. Destino del préstamo.** Cambia ventas (triangular, sorteado una vez por escenario), % de materia prima, costos fijos y por jornada,
desde el mes indicado.

**6. Monto máximo.** En cada mes del plazo: pago ≤ P{int(alfa * 100)}(F del mes) / DSCR mín. ({dscr_min:.2f}) − cuotas vigentes. El monto máximo
es el menor que cumple todos los meses.

**¿Es una proyección?** Sí. La simulación Monte Carlo *es* la proyección del flujo de caja: en lugar de un solo pronóstico,
da la distribución de resultados mes a mes (pestaña Flujo de caja proyectado). No es un complemento: el préstamo se paga con flujos futuros.

**Limitaciones.** Pocas semanas en el periodo base (hoy {R.stats_diarias['n_semanas_base']}); no se modela estacionalidad anual;
los efectos del destino son supuestos del dueño; la caja negativa simulada no equivale a la probabilidad de impago que estima un banco.
""")
