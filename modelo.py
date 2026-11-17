"""
Modelo financiero y simulación Monte Carlo de La H Burger.

Pregunta que responde: ¿cuánto puede pedir prestado La H Burger, a qué plazo y para qué destino,
sin que el riesgo de no cubrir la cuota supere el límite aceptado?

No depende de Streamlit: en Google Colab se usa igual (`from modelo import *`).

Flujo de caja mensual simulado (proyección estocástica):
    F_operativo = V·(1 − %CV) − Costos_por_jornada − Costos_fijos − Retiro_dueño
    F_neto      = F_operativo − Cuotas de deudas vigentes − Pago del préstamo nuevo
    DSCR        = F_operativo / (Cuotas vigentes + Pago del préstamo nuevo)
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
import math
import numpy as np
import pandas as pd

HOJAS = {
    "Productos": ["Producto", "Precio_venta"],
    "Ventas": ["Fecha", "Producto", "Cantidad", "Precio_unitario"],
    "Compras_MP": ["Fecha_compra", "Categoria", "Insumo", "Cantidad", "Unidad", "Precio_unitario"],
    "Costos_Operativos": ["Fecha", "Categoria", "Tipo_costo", "Valor"],
    "Activos": ["Fecha_compra", "Activo", "Cantidad", "Valor_unitario", "Vida_util_meses"],
    "Financiamiento": ["Fecha", "Tipo_fuente", "Monto", "Tasa_mensual", "Plazo_meses", "Estado"],
}
OPCIONALES = {"Activos", "Financiamiento"}
SEMANAS_POR_MES = 52 / 12  # 4,333…


# ======================================================================================
# Utilidades financieras
# ======================================================================================
def tasa_ea_a_mensual(ea: float) -> float:
    return (1 + ea) ** (1 / 12) - 1


def cuota_francesa(monto: float, i: float, n: int) -> float:
    if not monto or not n or n <= 0:
        return 0.0
    return monto * i / (1 - (1 + i) ** (-n)) if i > 0 else monto / n


def tabla_amortizacion(monto: float, i: float, plazo: int, gracia: int = 0, seguro: float = 0.0) -> pd.DataFrame:
    """Gracia a capital: en los primeros `gracia` meses se pagan solo intereses y seguro.
    Luego cuota fija (sistema francés) por (plazo − gracia) meses. El seguro se cobra sobre el saldo."""
    filas, saldo = [], float(monto)
    n_amort = max(plazo - gracia, 1)
    c = cuota_francesa(monto, i, n_amort)
    for m in range(1, plazo + 1):
        interes = saldo * i
        seg = saldo * seguro
        abono = 0.0 if m <= gracia else c - interes
        abono = min(abono, saldo)
        filas.append({"Mes": m, "Saldo_inicial": saldo, "Interes": interes, "Abono_capital": abono,
                      "Seguro": seg, "Pago_total": interes + abono + seg, "Saldo_final": saldo - abono})
        saldo -= abono
    return pd.DataFrame(filas)


def pagos_por_peso(i: float, plazo: int, gracia: int, seguro: float) -> np.ndarray:
    """Pago de cada mes por cada $1 prestado (sirve para despejar el monto máximo)."""
    return tabla_amortizacion(1.0, i, plazo, gracia, seguro)["Pago_total"].to_numpy()


def modalidad_credito(monto: float, smmlv: float) -> str:
    if monto <= 6 * smmlv:
        return "Popular productivo (≤ 6 SMMLV)"
    if monto <= 25 * smmlv:
        return "Productivo (> 6 y ≤ 25 SMMLV)"
    if monto <= 120 * smmlv:
        return "Productivo de mayor monto (> 25 y ≤ 120 SMMLV)"
    return "Supera el tope de microcrédito (> 120 SMMLV)"


# ======================================================================================
# 1. Carga y limpieza
# ======================================================================================
def _num(s):
    return pd.to_numeric(s, errors="coerce")


def _tipo_costo(x: str) -> str:
    s = str(x).lower()
    if "jornada" in s:
        return "Por jornada"
    if "variable" in s:
        return "Variable"
    return "Fijo mensual"


def cargar_excel(archivo) -> tuple[dict, list[str]]:
    avisos: list[str] = []
    xls = pd.read_excel(archivo, sheet_name=None)
    datos: dict[str, pd.DataFrame] = {}
    for hoja, cols in HOJAS.items():
        if hoja not in xls:
            if hoja in OPCIONALES:
                datos[hoja] = pd.DataFrame(columns=cols)
                continue
            raise ValueError(f"Falta la hoja '{hoja}'. Use la plantilla sin cambiar nombres de hojas.")
        df = xls[hoja].copy()
        df.columns = [str(c).strip() for c in df.columns]
        faltan = [c for c in cols if c not in df.columns]
        if faltan:
            raise ValueError(f"En la hoja '{hoja}' faltan las columnas: {', '.join(faltan)}.")
        datos[hoja] = df

    p = datos["Productos"].dropna(subset=["Producto"]).copy()
    p["Precio_venta"] = _num(p["Precio_venta"])
    datos["Productos"] = p
    precios = dict(zip(p["Producto"], p["Precio_venta"]))

    v = datos["Ventas"].copy()
    v["Fecha"] = pd.to_datetime(v["Fecha"], errors="coerce", dayfirst=True)
    v["Cantidad"] = _num(v["Cantidad"])
    v["Precio_unitario"] = _num(v["Precio_unitario"])
    v = v.dropna(subset=["Fecha", "Producto", "Cantidad"])
    sin = v["Precio_unitario"].isna()
    v.loc[sin, "Precio_unitario"] = v.loc[sin, "Producto"].map(precios)
    if v["Precio_unitario"].isna().any():
        avisos.append(f"{int(v['Precio_unitario'].isna().sum())} fila(s) de Ventas sin precio: se omitieron.")
        v = v.dropna(subset=["Precio_unitario"])
    v["Total"] = v["Cantidad"] * v["Precio_unitario"]
    datos["Ventas"] = v

    c = datos["Compras_MP"].copy()
    c["Fecha_compra"] = pd.to_datetime(c["Fecha_compra"], errors="coerce", dayfirst=True)
    for col in ("Cantidad", "Precio_unitario"):
        c[col] = _num(c[col])
    if "Fecha_vencimiento" in c.columns:
        c["Fecha_vencimiento"] = pd.to_datetime(c["Fecha_vencimiento"], errors="coerce", dayfirst=True)
    c = c.dropna(subset=["Fecha_compra", "Cantidad", "Precio_unitario"])
    c["Total"] = c["Cantidad"] * c["Precio_unitario"]
    datos["Compras_MP"] = c

    o = datos["Costos_Operativos"].copy()
    o["Fecha"] = pd.to_datetime(o["Fecha"], errors="coerce", dayfirst=True)
    o["Valor"] = _num(o["Valor"])
    o = o.dropna(subset=["Fecha", "Valor"])
    o["Tipo"] = o["Tipo_costo"].map(_tipo_costo)
    datos["Costos_Operativos"] = o

    a = datos["Activos"].copy()
    for col in ("Cantidad", "Valor_unitario", "Vida_util_meses"):
        a[col] = _num(a[col])
    a["Valor_residual"] = _num(a["Valor_residual"]).fillna(0) if "Valor_residual" in a.columns else 0.0
    a = a.dropna(subset=["Activo", "Valor_unitario"])
    a["Cantidad"] = a["Cantidad"].fillna(1)
    a["Valor_total"] = a["Cantidad"] * a["Valor_unitario"]
    a["Depreciacion_mensual"] = np.where(a["Vida_util_meses"] > 0, (a["Valor_total"] - a["Valor_residual"]) / a["Vida_util_meses"], 0.0)
    datos["Activos"] = a

    f = datos["Financiamiento"].copy()
    for col in ("Monto", "Tasa_mensual", "Plazo_meses"):
        f[col] = _num(f[col])
    f["Saldo_actual"] = _num(f["Saldo_actual"]) if "Saldo_actual" in f.columns else np.nan
    f = f.dropna(subset=["Monto"])
    if (f["Tasa_mensual"] > 1).any():
        avisos.append("Alguna Tasa_mensual parece escrita como entero (2,5 en vez de 2,5 %). Se dividió entre 100.")
        f.loc[f["Tasa_mensual"] > 1, "Tasa_mensual"] /= 100
    f["Tasa_mensual"] = f["Tasa_mensual"].fillna(0)
    f["Cuota_mensual"] = [cuota_francesa(m, t, int(n) if pd.notna(n) else 0) for m, t, n in zip(f["Monto"], f["Tasa_mensual"], f["Plazo_meses"])]
    f.loc[f["Tipo_fuente"].eq("Aporte propio"), "Cuota_mensual"] = 0.0
    datos["Financiamiento"] = f

    rt = xls.get("Retiros_Dueno", pd.DataFrame(columns=["Fecha", "Valor"])).copy()
    if len(rt):
        rt["Fecha"] = pd.to_datetime(rt["Fecha"], errors="coerce", dayfirst=True)
        rt["Valor"] = _num(rt["Valor"])
        rt = rt.dropna(subset=["Fecha", "Valor"])
    datos["Retiros_Dueno"] = rt

    # Hojas de configuración (opcionales)
    datos["Parametros"] = {}
    if "Parametros" in xls:
        par = xls["Parametros"].dropna(subset=["Parametro"])
        datos["Parametros"] = dict(zip(par["Parametro"].astype(str).str.strip(), par["Valor"]))
    esc = xls["Escenarios_Prestamo"].copy() if "Escenarios_Prestamo" in xls else pd.DataFrame()
    if len(esc):
        esc = esc.dropna(subset=["Destino"])
        esc = esc[pd.to_numeric(esc.get("Inversion"), errors="coerce").notna()]  # descarta filas de notas
    datos["Escenarios"] = esc.reset_index(drop=True)

    if v.empty:
        avisos.append("La hoja Ventas no tiene registros válidos.")
    return datos, avisos


def caja_historica(datos: dict) -> pd.DataFrame:
    """Caja acumulada a hoy = ventas − compras MP − costos operativos − retiros + préstamos y aportes
    recibidos − activos comprados − abonos a capital de deudas (monto − saldo actual)."""
    f = datos["Financiamiento"]
    deuda = f[~f["Tipo_fuente"].eq("Aporte propio")] if len(f) else f
    abonos = float(deuda["Monto"].sum() - deuda["Saldo_actual"].fillna(deuda["Monto"]).sum()) if len(deuda) else 0.0
    partidas = [("Ventas", datos["Ventas"]["Total"].sum()),
                ("Compras de materia prima", -datos["Compras_MP"]["Total"].sum()),
                ("Costos operativos", -datos["Costos_Operativos"]["Valor"].sum()),
                ("Retiros del dueño", -datos["Retiros_Dueno"]["Valor"].sum() if len(datos.get("Retiros_Dueno", [])) else 0.0),
                ("Préstamos y aportes recibidos", f["Monto"].sum() if len(f) else 0.0),
                ("Activos comprados", -datos["Activos"]["Valor_total"].sum() if len(datos["Activos"]) else 0.0),
                ("Abonos a capital de deudas", -abonos)]
    df = pd.DataFrame(partidas, columns=["Partida", "Valor"])
    return pd.concat([df, pd.DataFrame([{"Partida": "Caja estimada hoy", "Valor": df["Valor"].sum()}])], ignore_index=True)


# ======================================================================================
# 2. Estadística descriptiva y supuestos estimados
# ======================================================================================
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


@dataclass
class Resumen:
    ventas_diarias: pd.DataFrame
    semanas: pd.DataFrame            # todas las semanas: Semana, Ventas, Unidades, Dias, Costos_jornada, Compras_MP, En_base
    semanas_base: pd.DataFrame
    meses: pd.DataFrame
    por_producto: pd.DataFrame
    por_dia_semana: pd.DataFrame
    mp_por_categoria: pd.DataFrame
    fecha_inicio_base: pd.Timestamp
    cv_muestras: np.ndarray          # razones MP/ventas (ventanas de 4 semanas del periodo base)
    cv_hist: dict
    costos_fijos_mes: float
    costos_variables_pct: float
    depreciacion_mes: float
    deudas: pd.DataFrame             # deudas vigentes: Fuente, Cuota, Meses_restantes, Saldo
    stats_diarias: dict = field(default_factory=dict)
    primera_venta: pd.Timestamp | None = None


def resumir(datos: dict, fecha_inicio_base=None) -> Resumen:
    v, c, o = datos["Ventas"], datos["Compras_MP"], datos["Costos_Operativos"]
    par = datos.get("Parametros", {})
    if fecha_inicio_base is None:
        fecha_inicio_base = par.get("Fecha_inicio_base")
    fib = pd.to_datetime(fecha_inicio_base) if fecha_inicio_base is not None and not pd.isna(fecha_inicio_base) else v["Fecha"].min()

    diarias = v.groupby("Fecha").agg(Total=("Total", "sum"), Unidades=("Cantidad", "sum")).reset_index()
    piv = v.pivot_table(index="Fecha", columns="Producto", values="Cantidad", aggfunc="sum", fill_value=0)
    diarias = diarias.merge(piv.reset_index(), on="Fecha", how="left")
    diarias["Dia_semana"] = diarias["Fecha"].dt.dayofweek.map(dict(enumerate(DIAS)))

    # semanas (lunes a domingo)
    wk = lambda s: (s - pd.to_timedelta(s.dt.dayofweek, unit="D")).dt.normalize()
    sem = v.assign(Semana=wk(v["Fecha"])).groupby("Semana").agg(Ventas=("Total", "sum"), Unidades=("Cantidad", "sum"), Dias=("Fecha", "nunique"))
    oj = o[o["Tipo"].eq("Por jornada")]
    sem["Costos_jornada"] = oj.assign(Semana=wk(oj["Fecha"])).groupby("Semana")["Valor"].sum()
    sem["Compras_MP"] = c.assign(Semana=wk(c["Fecha_compra"])).groupby("Semana")["Total"].sum()
    sem = sem.fillna(0).reset_index()
    sem["En_base"] = sem["Semana"] >= (fib - pd.Timedelta(days=int(fib.dayofweek))).normalize()
    base = sem[sem["En_base"] & (sem["Ventas"] > 0)].reset_index(drop=True)

    # meses
    vm = v.assign(Mes=v["Fecha"].dt.to_period("M"))
    m = vm.groupby("Mes").agg(Ventas=("Total", "sum"), Dias_operados=("Fecha", "nunique"))
    m["Compras_MP"] = c.assign(Mes=c["Fecha_compra"].dt.to_period("M")).groupby("Mes")["Total"].sum()
    for t in ("Fijo mensual", "Por jornada", "Variable"):
        ot = o[o["Tipo"].eq(t)]
        m[t] = ot.assign(Mes=ot["Fecha"].dt.to_period("M")).groupby("Mes")["Valor"].sum()
    m = m.fillna(0)
    m["CV_pct"] = np.where(m["Ventas"] > 0, m["Compras_MP"] / m["Ventas"], np.nan)
    m["F_operativo"] = m["Ventas"] - m["Compras_MP"] - m["Fijo mensual"] - m["Por jornada"] - m["Variable"]
    m = m.reset_index(); m["Mes"] = m["Mes"].astype(str)

    # %CV: razón compras/ventas en ventanas móviles de 4 semanas del periodo base (suaviza el desfase de las compras)
    if len(base) >= 4:
        r = (base["Compras_MP"].rolling(4).sum() / base["Ventas"].rolling(4).sum()).dropna().to_numpy()
    else:
        r = m["CV_pct"].dropna().to_numpy()
    r = r[np.isfinite(r)]
    if r.size >= 2:
        cv = {"min": float(r.min()), "moda": float(np.median(r)), "max": float(r.max()), "media": float(r.mean())}
    elif r.size == 1:
        cv = {"min": float(r[0]) - 0.03, "moda": float(r[0]), "max": float(r[0]) + 0.03, "media": float(r[0])}
    else:
        cv = {"min": 0.40, "moda": 0.45, "max": 0.50, "media": 0.45}
    if cv["max"] - cv["min"] < 0.02:  # amplitud mínima de ±1 punto
        cv["min"] -= 0.01; cv["max"] += 0.01

    # costos fijos mensuales: promedio de los meses del periodo base (o de todos)
    ofm = o[o["Tipo"].eq("Fijo mensual")]
    fm = ofm.assign(Mes=ofm["Fecha"].dt.to_period("M")).groupby("Mes")["Valor"].sum()
    fm_base = fm[fm.index >= fib.to_period("M")]
    cf = float((fm_base if len(fm_base) else fm).mean()) if len(fm) else 0.0
    ventas_base = float(base["Ventas"].sum()) if len(base) else float(v["Total"].sum())
    ov = o[o["Tipo"].eq("Variable") & (o["Fecha"] >= fib)]
    var_pct = float(ov["Valor"].sum() / ventas_base) if ventas_base > 0 else 0.0

    prod = v.groupby("Producto").agg(Unidades=("Cantidad", "sum"), Ventas=("Total", "sum")).reset_index()
    prod["Participacion"] = prod["Ventas"] / prod["Ventas"].sum()
    prod = prod.sort_values("Ventas", ascending=False)
    dbase = diarias[diarias["Fecha"] >= fib]
    dsem = dbase.groupby("Dia_semana").agg(Ventas_prom=("Total", "mean"), Ventas_desv=("Total", "std"),
                                           Unidades_prom=("Unidades", "mean"), Dias=("Total", "size")).reindex(DIAS).dropna(subset=["Dias"]).reset_index()
    mpcat = c.groupby("Categoria").agg(Gasto=("Total", "sum"), Compras=("Total", "size")).reset_index()
    mpcat["Participacion"] = mpcat["Gasto"] / mpcat["Gasto"].sum() if len(mpcat) else 0
    mpcat = mpcat.sort_values("Gasto", ascending=False)

    a, f = datos["Activos"], datos["Financiamiento"]
    dep = float(a["Depreciacion_mensual"].sum()) if len(a) else 0.0
    vig = f[f["Estado"].fillna("Vigente").eq("Vigente") & (f["Cuota_mensual"] > 0)] if len(f) else f
    deudas = []
    for _, row in vig.iterrows():
        saldo = row.get("Saldo_actual")
        if pd.notna(saldo) and row["Cuota_mensual"] > 0:
            if row["Tasa_mensual"] > 0:
                i = row["Tasa_mensual"]; c_ = row["Cuota_mensual"]
                n = math.ceil(-math.log(1 - saldo * i / c_) / math.log(1 + i)) if saldo * i < c_ else int(row["Plazo_meses"])
            else:
                n = math.ceil(saldo / row["Cuota_mensual"])
        else:
            n = int(row["Plazo_meses"]) if pd.notna(row["Plazo_meses"]) else 0
            saldo = np.nan
        deudas.append({"Fuente": row.get("Fuente", row["Tipo_fuente"]), "Cuota": float(row["Cuota_mensual"]),
                       "Meses_restantes": int(n), "Saldo": saldo})
    deudas = pd.DataFrame(deudas, columns=["Fuente", "Cuota", "Meses_restantes", "Saldo"])

    t = dbase["Total"]
    stats = {"n_dias": int(t.size), "media": float(t.mean()) if t.size else 0, "mediana": float(t.median()) if t.size else 0,
             "desv": float(t.std()) if t.size > 1 else 0, "min": float(t.min()) if t.size else 0, "max": float(t.max()) if t.size else 0,
             "cv_coef": float(t.std() / t.mean()) if t.size > 1 and t.mean() else 0,
             "asimetria": float(t.skew()) if t.size > 2 else 0, "curtosis": float(t.kurt()) if t.size > 3 else 0,
             "p25": float(t.quantile(.25)) if t.size else 0, "p75": float(t.quantile(.75)) if t.size else 0,
             "venta_semanal_media": float(base["Ventas"].mean()) if len(base) else 0,
             "venta_semanal_desv": float(base["Ventas"].std()) if len(base) > 1 else 0,
             "n_semanas_base": int(len(base))}
    return Resumen(diarias, sem, base, m, prod, dsem, mpcat, fib, r, cv, cf, var_pct, dep, deudas, stats,
                   v["Fecha"].min() if len(v) else None)


# ======================================================================================
# 3. Destino del préstamo (efectos) y condiciones del crédito
# ======================================================================================
@dataclass
class Destino:
    nombre: str = "Solo capacidad de pago (sin destino)"
    inversion: float = 0.0
    ventas_min: float = 0.0          # cambio % de ventas (triangular)
    ventas_probable: float = 0.0
    ventas_max: float = 0.0
    cambio_cv_pp: float = 0.0        # puntos porcentuales sobre %CV (−0,01 = baja 1 punto)
    costo_fijo_adicional: float = 0.0
    cambio_costos_jornada: float = 0.0
    meses_hasta_efecto: int = 0
    paga_deudas: bool = False

    @staticmethod
    def desde_fila(r) -> "Destino":
        g = lambda k, d=0.0: (float(r[k]) if k in r and pd.notna(r[k]) else d)
        return Destino(str(r["Destino"]), g("Inversion"), g("Ventas_min"), g("Ventas_probable"), g("Ventas_max"),
                       g("Cambio_CV_pp"), g("Costo_fijo_adicional_mes"), g("Cambio_costos_jornada"),
                       int(g("Meses_hasta_efecto")), str(r.get("Paga_deudas_vigentes", "No")).strip().lower().startswith("s"))


DESTINOS_POR_DEFECTO = [
    Destino(),
    Destino("Activos para la operación", 4_500_000, 0.03, 0.08, 0.15, -0.01, 50_000, 0.05, 1),
    Destino("Expandir a un local", 25_000_000, -0.10, 0.45, 0.80, 0.0, 3_000_000, 0.40, 3),
    Destino("Capital de trabajo (compras al por mayor)", 3_000_000, 0, 0, 0, -0.015, 0, 0, 0),
    Destino("Pagar deuda familiar + capital de trabajo", 2_500_000, 0, 0, 0, -0.01, 0, 0, 0, True),
    Destino("Personalizado"),
]


@dataclass
class Credito:
    monto: float = 0.0
    tasa_ea: float = 0.42
    plazo: int = 24
    gracia: int = 0
    comision_inicial: float = 0.03   # % del monto, descontado al desembolsar
    seguro_mensual: float = 0.001    # % mensual sobre el saldo

    @property
    def tasa_mensual(self) -> float:
        return tasa_ea_a_mensual(self.tasa_ea)

    @property
    def desembolso_neto(self) -> float:
        return self.monto * (1 - self.comision_inicial)

    def pagos(self) -> np.ndarray:
        return tabla_amortizacion(self.monto, self.tasa_mensual, self.plazo, self.gracia, self.seguro_mensual)["Pago_total"].to_numpy() if self.monto > 0 else np.zeros(self.plazo)

    def costo_total(self) -> float:
        return float(self.pagos().sum() - self.monto + self.monto * self.comision_inicial)


# ======================================================================================
# 4. Simulación Monte Carlo
# ======================================================================================
@dataclass
class Supuestos:
    n_sim: int = 10_000
    horizonte: int = 24
    factor_demanda: float = 1.0
    choque_mensual: float = 0.05
    cv_dist: str = "triangular"
    cv_min: float = 0.40
    cv_moda: float = 0.45
    cv_max: float = 0.50
    costos_fijos: float = 0.0
    costos_variables_pct: float = 0.0
    retiro_min: float = 0.0
    retiro_max: float = 0.0
    caja_inicial: float = 0.0
    dscr_min: float = 1.2
    alfa: float = 0.05
    umbral_aprobacion: float = 0.96
    semilla: int | None = 2026


@dataclass
class Resultado:
    ventas: np.ndarray; cv: np.ndarray; costos_jornada: np.ndarray; costos_fijos: np.ndarray
    retiro: np.ndarray; f_op: np.ndarray; deuda_vigente: np.ndarray; pago_nuevo: np.ndarray
    f_neto: np.ndarray; caja: np.ndarray; efecto_ventas: np.ndarray
    supuestos: Supuestos; destino: Destino; credito: Credito; aporte_propio: float

    @property
    def servicio_deuda(self) -> np.ndarray:
        return self.deuda_vigente + self.pago_nuevo

    def prob_f_op_positivo(self) -> float:
        return float((self.f_op > 0).mean())

    def prob_f_neto_positivo(self) -> float:
        return float((self.f_neto > 0).mean())

    def prob_mes_con_deficit(self) -> float:
        return float((self.f_neto < 0).any(axis=1).mean())

    def prob_caja_negativa(self) -> float:
        return float((self.caja < 0).any(axis=1).mean())

    def dscr(self) -> np.ndarray:
        s = self.servicio_deuda
        return np.where(s[None, :] > 0, self.f_op / np.where(s > 0, s, 1)[None, :], np.nan)

    def prob_dscr_bajo(self) -> float:
        d = self.dscr()
        mask = ~np.isnan(d)
        return float((d[mask] < self.supuestos.dscr_min).mean()) if mask.any() else 0.0

    def percentiles(self, arr=None, qs=(5, 10, 25, 50, 75, 90, 95)) -> dict:
        a = (self.f_op if arr is None else arr).ravel()
        return {f"P{q}": float(np.percentile(a, q)) for q in qs}

    def proyeccion(self) -> pd.DataFrame:
        """Flujo de caja proyectado por mes (percentiles de los escenarios simulados)."""
        H = self.f_op.shape[1]
        p = lambda a, q: np.percentile(a, q, axis=0)
        return pd.DataFrame({
            "Mes": np.arange(1, H + 1),
            "Ventas_P50": p(self.ventas, 50),
            "Costo_MP_P50": p(self.ventas * self.cv, 50),
            "Costos_jornada_P50": p(self.costos_jornada, 50),
            "Costos_fijos": self.costos_fijos.mean(axis=0),
            "Retiro_prom": self.retiro.mean(axis=0),
            "F_operativo_P10": p(self.f_op, 10), "F_operativo_P50": p(self.f_op, 50), "F_operativo_P90": p(self.f_op, 90),
            "Cuotas_vigentes": self.deuda_vigente, "Pago_prestamo": self.pago_nuevo,
            "F_neto_P50": p(self.f_neto, 50),
            "Caja_P10": p(self.caja, 10), "Caja_P50": p(self.caja, 50), "Caja_P90": p(self.caja, 90),
            "P_F_neto_negativo": (self.f_neto < 0).mean(axis=0),
        })


def _calendario_deudas(deudas: pd.DataFrame, H: int) -> np.ndarray:
    out = np.zeros(H)
    for _, d in deudas.iterrows():
        n = min(int(d["Meses_restantes"]), H)
        out[:n] += d["Cuota"]
    return out


def simular(R: Resumen, s: Supuestos, destino: Destino | None = None, credito: Credito | None = None) -> Resultado:
    destino = destino or Destino()
    credito = credito or Credito()
    rng = np.random.default_rng(s.semilla)
    base = R.semanas_base
    if base.empty:
        raise ValueError("No hay semanas con ventas en el periodo base.")
    H = int(max(s.horizonte, credito.plazo if credito.monto > 0 else 0))
    n = int(s.n_sim)

    # Ventas y costos por jornada: remuestreo de SEMANAS reales del periodo base (bootstrap por bloques)
    vs, cj = base["Ventas"].to_numpy(float), base["Costos_jornada"].to_numpy(float)
    W = 5  # máximo de semanas que puede tener un mes
    idx = rng.integers(0, vs.size, size=(n, H, W))
    extra = rng.random((n, H)) < (SEMANAS_POR_MES - 4)       # meses con 5 semanas (≈ 33 %)
    peso = np.ones((n, H, W)); peso[:, :, 4] = extra
    ventas = (vs[idx] * peso).sum(axis=2)
    costos_j = (cj[idx] * peso).sum(axis=2)
    ventas *= s.factor_demanda
    if s.choque_mensual > 0:
        ventas *= np.clip(rng.normal(1.0, s.choque_mensual, (n, H)), 0, None)

    # %CV
    if s.cv_dist == "uniforme":
        cv = rng.uniform(s.cv_min, s.cv_max, (n, H))
    else:
        mo = min(max(s.cv_moda, s.cv_min), s.cv_max)
        cv = rng.triangular(s.cv_min, mo, s.cv_max, (n, H)) if s.cv_max > s.cv_min else np.full((n, H), s.cv_min)
    cv = cv + s.costos_variables_pct

    # Efecto del destino (un efecto por escenario simulado, desde el mes indicado)
    activo = (np.arange(1, H + 1) > destino.meses_hasta_efecto) if credito.monto > 0 or destino.inversion > 0 else np.zeros(H, bool)
    lo, mo_, hi = destino.ventas_min, destino.ventas_probable, destino.ventas_max
    efecto = rng.triangular(lo, min(max(mo_, lo), hi), hi, n) if hi > lo else np.full(n, lo)
    ventas = ventas * (1 + np.outer(efecto, activo))
    cv = cv + destino.cambio_cv_pp * activo
    costos_j = costos_j * (1 + destino.cambio_costos_jornada * activo)
    fijos = np.full((n, H), s.costos_fijos) + destino.costo_fijo_adicional * activo

    retiro = rng.uniform(s.retiro_min, max(s.retiro_max, s.retiro_min), (n, H))
    f_op = ventas * (1 - cv) - costos_j - fijos - retiro

    deuda_vig = np.zeros(H) if destino.paga_deudas else _calendario_deudas(R.deudas, H)
    pagos = credito.pagos()
    pago_nuevo = np.zeros(H); pago_nuevo[:min(H, pagos.size)] = pagos[:H]
    f_neto = f_op - deuda_vig - pago_nuevo

    saldo_deudas = float(R.deudas["Saldo"].fillna(0).sum()) if destino.paga_deudas else 0.0
    necesidad = destino.inversion + saldo_deudas
    aporte = max(0.0, necesidad - credito.desembolso_neto) if (credito.monto > 0 or necesidad > 0) else 0.0
    sobrante = max(0.0, credito.desembolso_neto - necesidad)
    caja = s.caja_inicial - aporte + sobrante + np.cumsum(f_neto, axis=1)
    return Resultado(ventas, cv, costos_j, fijos, retiro, f_op, deuda_vig, pago_nuevo, f_neto, caja, efecto,
                     s, destino, credito, aporte)


def monto_maximo(res: Resultado, credito: Credito, alfa: float | None = None, dscr_min: float | None = None) -> dict:
    """Monto máximo tal que en cada mes del plazo: Pago(M)_m ≤ P_alfa(F_op_m)/DSCR_min − cuotas vigentes_m.
    Pago(M)_m = M × pago_por_peso_m (incluye gracia, interés y seguro)."""
    alfa = res.supuestos.alfa if alfa is None else alfa
    dscr_min = res.supuestos.dscr_min if dscr_min is None else dscr_min
    a = pagos_por_peso(credito.tasa_mensual, credito.plazo, credito.gracia, credito.seguro_mensual)
    H = res.f_op.shape[1]
    k = min(credito.plazo, H)
    q = np.percentile(res.f_op[:, :k], alfa * 100, axis=0)
    cap = q / dscr_min - res.deuda_vigente[:k]
    limites = np.where(a[:k] > 0, cap / a[:k], np.inf)
    m_max = max(0.0, float(np.min(limites)))
    mes_critico = int(np.argmin(limites)) + 1
    return {"monto_max": m_max, "mes_critico": mes_critico, "cuota_max_mes_critico": float(max(cap[mes_critico - 1], 0)),
            "percentil_f_op_mes_critico": float(q[mes_critico - 1])}


def comparar_plazos(R: Resumen, s: Supuestos, destino: Destino, credito: Credito, plazos=(6, 12, 18, 24, 36, 48, 60)) -> pd.DataFrame:
    filas = []
    for pl in plazos:
        cr = Credito(credito.monto, credito.tasa_ea, pl, min(credito.gracia, pl - 1), credito.comision_inicial, credito.seguro_mensual)
        sp = Supuestos(**{**asdict(s), "horizonte": max(s.horizonte, pl)})
        r = simular(R, sp, destino, cr)
        mm = monto_maximo(r, cr)
        pag = cr.pagos()
        filas.append({"Plazo (meses)": pl, "Pago mensual del monto solicitado": float(pag[cr.gracia]) if credito.monto > 0 else 0.0,
                      "Costo total del crédito": cr.costo_total(), "Monto máximo": mm["monto_max"],
                      "P(DSCR < mín.)": r.prob_dscr_bajo(), "P(caja < 0 algún mes)": r.prob_caja_negativa()})
    return pd.DataFrame(filas)


def comparar_destinos(R: Resumen, s: Supuestos, destinos: list[Destino], credito: Credito) -> pd.DataFrame:
    filas = []
    for d in destinos:
        cr = Credito(max(credito.monto, d.inversion / (1 - credito.comision_inicial)) if d.inversion > 0 else credito.monto,
                     credito.tasa_ea, credito.plazo, credito.gracia, credito.comision_inicial, credito.seguro_mensual)
        r = simular(R, s, d, cr)
        mm = monto_maximo(r, cr)
        filas.append({"Destino": d.nombre, "Monto necesario": cr.monto, "Monto máximo": mm["monto_max"],
                      "F operativo P50 (mes 12)": float(np.percentile(r.f_op[:, min(11, r.f_op.shape[1] - 1)], 50)),
                      "P(DSCR < mín.)": r.prob_dscr_bajo(), "P(caja < 0 algún mes)": r.prob_caja_negativa(),
                      "Viable": cr.monto <= mm["monto_max"]})
    return pd.DataFrame(filas)


def sensibilidad(res: Resultado, delta: float = 0.10) -> pd.DataFrame:
    """Tornado (uno a la vez, ±delta) sobre el F operativo promedio."""
    V = res.ventas.mean(); CV = res.cv.mean(); CJ = res.costos_jornada.mean(); CF = res.costos_fijos.mean(); RT = res.retiro.mean()
    F = lambda v=V, c=CV, j=CJ, f=CF, r=RT: v * (1 - c) - j - f - r
    base = F()
    filas = [("Ventas", F(v=V * (1 - delta)), F(v=V * (1 + delta))),
             ("% costo de materia prima", F(c=CV * (1 + delta)), F(c=CV * (1 - delta))),
             ("Costos por jornada (ayudantes, gas)", F(j=CJ * (1 + delta)), F(j=CJ * (1 - delta))),
             ("Costos fijos (espacio, servicios)", F(f=CF * (1 + delta)), F(f=CF * (1 - delta))),
             ("Retiro del dueño", F(r=RT * (1 + delta)), F(r=RT * (1 - delta)))]
    df = pd.DataFrame(filas, columns=["Variable", "Desfavorable", "Favorable"])
    df["Base"] = base; df["Rango"] = df["Favorable"] - df["Desfavorable"]
    return df.sort_values("Rango")


# ======================================================================================
# 5. Requisitos habituales de microcrédito
# ======================================================================================
REQUISITOS_DOCUMENTALES = [
    "Cédula de ciudadanía vigente (mayor de edad; muchas entidades piden entre 18 y 70 años)",
    "Historial en centrales de riesgo (Datacrédito / TransUnion) sin reportes negativos vigentes",
    "RUT (algunas entidades lo exigen; se obtiene gratis en la DIAN)",
    "Registro en Cámara de Comercio (suele pedirse en montos mayores)",
    "Soportes de ingresos: registros de ventas, extractos bancarios o de Nequi / Daviplata",
    "Permiso o contrato del espacio donde opera el negocio",
    "Recibo de servicio público con la dirección de residencia",
    "Referencias personales y comerciales (proveedores)",
    "Codeudor o garantía del Fondo Nacional de Garantías (FNG)",
    "Visita del asesor de crédito al negocio",
    "Cotizaciones de lo que se va a comprar (destino del crédito)",
    "Flujo de caja proyectado y plan de inversión (lo genera esta app)",
]


def requisitos_automaticos(R: Resumen, res: Resultado, credito: Credito, mm: dict, par: dict, hoy=None) -> pd.DataFrame:
    hoy = pd.Timestamp(hoy or pd.Timestamp.today().normalize())
    filas = []
    meses_op = (hoy - R.primera_venta).days / 30.44 if R.primera_venta is not None else 0
    filas.append(("Antigüedad del negocio ≥ 6 meses", meses_op >= 6, f"{meses_op:.1f} meses desde la primera venta"))
    meses_reg = R.ventas_diarias["Fecha"].dt.to_period("M").nunique()
    filas.append(("Registros de ventas de al menos 3 meses", meses_reg >= 3, f"{meses_reg} meses con ventas registradas"))
    p = res.prob_dscr_bajo()
    filas.append((f"Capacidad de pago: P(DSCR < {res.supuestos.dscr_min:.2f}) ≤ {res.supuestos.alfa:.0%}", p <= res.supuestos.alfa,
                  f"P = {p:.1%}"))
    filas.append(("Monto solicitado ≤ monto máximo estimado", credito.monto <= mm["monto_max"],
                  f"{credito.monto:,.0f} frente a {mm['monto_max']:,.0f}".replace(",", ".")))
    smmlv = float(par.get("SMMLV", 1_750_905))
    mod = modalidad_credito(credito.monto, smmlv)
    usura = {"Popular": par.get("Usura_popular_productivo", 0.8844), "Productivo (": par.get("Usura_productivo", 0.5987),
             "Productivo de": par.get("Usura_mayor_monto", 0.4235)}
    tope = next((float(v) for k, v in usura.items() if mod.startswith(k)), None)
    filas.append(("Monto dentro del rango de microcrédito (≤ 120 SMMLV)", credito.monto <= 120 * smmlv, mod))
    if tope is not None:
        filas.append(("Tasa ofrecida ≤ tasa de usura de la modalidad", credito.tasa_ea <= tope,
                      f"{credito.tasa_ea:.2%} EA frente a usura {tope:.2%} EA"))
    filas.append(("Destino del crédito definido (productivo)", not res.destino.nombre.startswith("Solo"), res.destino.nombre))
    filas.append(("Préstamo cubre la inversión (sin aporte propio adicional)", res.aporte_propio <= 0,
                  "Aporte propio necesario: " + f"{res.aporte_propio:,.0f}".replace(",", ".")))
    if len(R.deudas):
        filas.append(("Deudas vigentes declaradas", True, f"{len(R.deudas)} deuda(s); cuota total {R.deudas['Cuota'].sum():,.0f}".replace(",", ".")))
    return pd.DataFrame(filas, columns=["Requisito", "Cumple", "Detalle"])
