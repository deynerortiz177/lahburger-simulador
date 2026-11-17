"""
Genera los archivos de Excel de La H Burger:
  1) Plantilla_Registro_LaHBurger.xlsx  (vacía, para llenar)
  2) Datos_Simulados_LaHBurger.xlsx      (datos SIMULADOS con las reglas del negocio dadas por el dueño)

Uso:  python generar_datos.py <carpeta_salida> [semilla]

Reglas de simulación (dadas por el dueño, 16/11/2026):
- Inicio de operación: viernes 15/05/2026. Datos hasta el domingo 15/11/2026.
- 15/05 a 04/06: solo "La Paisa"; abre viernes y sábado; 80–100 hamburguesas por fin de semana (vie+sáb).
- 05/06 a 17/09: entra "La Polluda"; 100–150 hamburguesas por fin de semana entre las dos.
- Desde 18/09: entran "La Nuevona" y "Paisa Ardiente"; 150–200 por fin de semana entre las cuatro.
- Desde la semana siguiente al 18/09 (jueves 24/09): abre los jueves (35–50 hamburguesas) y los domingos
  anteriores a un lunes festivo (35–50 hamburguesas): 11/10, 01/11 y 15/11 de 2026.
- Ayudantes: viernes y sábado 3 × $100.000; jueves y domingo 2 × $100.000.
- Espacio y servicios: $800.000 al mes. Gas: una pipeta de $100.000 al mes.
- La Paisa y La Polluda son las más vendidas; La Nuevona y Paisa Ardiente se venden menos.
- Retiros del dueño: entre $2.000.000 y $2.500.000 al final de cada mes (medio mes en mayo).
- Caja actual: se calcula con el flujo acumulado (fórmula en la hoja Parametros).
- Activos (≈ $3.000.000) comprados antes de abrir, financiados por un familiar; saldo actual $1.000.000.
- Retiro del dueño: entre $2.000.000 y $2.500.000 al mes.
- Carnes: se compran listas (filete de pechuga porcionado que se apana en el puesto; chorizo artesanal hecho).
"""
import sys
from datetime import date, timedelta
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

FONT = "Arial"
MAXR = 3000
HDR_FILL = PatternFill("solid", fgColor="0D1014")   # negro de la marca
HDR_FONT = "D8AC55"                                  # dorado de la marca
IN_FILL = PatternFill("solid", fgColor="FCF0CF")    # crema de la marca: celdas a llenar
F_FILL = PatternFill("solid", fgColor="E7E6E6")     # fórmulas
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

PRODUCTOS = [("La Polluda", 26000), ("La Paisa", 23000), ("La Nuevona", 27000), ("Paisa Ardiente", 24000)]

LISTAS = {
    "Productos": [p for p, _ in PRODUCTOS],
    "Medio_pago": ["Efectivo", "Nequi", "Daviplata", "Transferencia", "Tarjeta"],
    "Categoria_MP": ["Pan", "Proteína", "Lácteos y quesos", "Frutas y verduras", "Salsas y aderezos",
                     "Tocineta y embutidos", "Papas y acompañamientos", "Insumos de apanado", "Bebidas",
                     "Empaques y desechables", "Aceite y grasas", "Otros insumos"],
    "Unidad": ["und", "kg", "g", "L", "ml", "paquete", "caja", "cubeta", "docena"],
    "Categoria_costo": ["Gas / combustible", "Arriendo / uso de espacio", "Permisos y licencias", "Servicios (agua, luz, celular)",
                        "Transporte", "Nómina / ayudantes", "Mantenimiento", "Publicidad", "Comisiones plataformas / datáfono",
                        "Aseo y bioseguridad", "Impuestos", "Otros"],
    "Tipo_costo": ["Fijo mensual", "Por jornada", "Variable (% ventas)"],
    "Categoria_activo": ["Equipo de cocina", "Carpa / estructura", "Mobiliario", "Refrigeración", "Tecnología", "Vehículo", "Utensilios", "Otros"],
    "Estado_activo": ["Nuevo", "Bueno", "Regular", "Malo"],
    "Tipo_fuente": ["Aporte propio", "Banco", "Cooperativa", "Microfinanciera", "Fintech", "Familiar / amigos",
                    "Prestamista (gota a gota)", "Proveedor"],
    "Estado_deuda": ["Vigente", "Pagada"],
    "Si_No": ["Sí", "No"],
}

SHEETS = {
    "Ventas": [
        ("Fecha", 12, "date", None, "Fecha de la venta (dd/mm/aaaa)."),
        ("Producto", 16, "in", "Productos", "Elija de la lista (hoja Productos)."),
        ("Cantidad", 10, "in", None, "Unidades vendidas."),
        ("Precio_unitario", 15, "money_f", None, "Se llena solo con el precio de la hoja Productos. Puede sobrescribirlo si hubo descuento."),
        ("Medio_pago", 14, "in", "Medio_pago", None),
        ("Observacion", 28, "in", None, "Opcional: clima, evento, festivo, quincena."),
        ("Total", 14, "f", None, "Fórmula: Cantidad × Precio. No escribir."),
    ],
    "Compras_MP": [
        ("Fecha_compra", 13, "date", None, None),
        ("Categoria", 22, "in", "Categoria_MP", None),
        ("Insumo", 30, "in", None, "Ej.: Pan brioche, Filete de pechuga porcionado, Chorizo artesanal."),
        ("Marca", 16, "in", None, None),
        ("Proveedor", 24, "in", None, None),
        ("Cantidad", 10, "in", None, "Cantidad comprada en la unidad indicada."),
        ("Unidad", 10, "in", "Unidad", None),
        ("Precio_unitario", 15, "money", None, "Precio por unidad (no el total)."),
        ("Fecha_vencimiento", 15, "date", None, "Recomendado: sirve para medir desperdicio."),
        ("Factura_lote", 14, "in", None, "Opcional."),
        ("Total", 14, "f", None, "Fórmula: Cantidad × Precio_unitario. No escribir."),
    ],
    "Costos_Operativos": [
        ("Fecha", 12, "date", None, None),
        ("Categoria", 30, "in", "Categoria_costo", None),
        ("Descripcion", 34, "in", None, None),
        ("Tipo_costo", 18, "in", "Tipo_costo", "Fijo mensual: se paga una vez al mes (espacio, servicios). Por jornada: depende de los días que se abre (ayudantes, gas). Variable (% ventas): sube con las ventas (comisión datáfono)."),
        ("Valor", 14, "money", None, None),
        ("Medio_pago", 14, "in", "Medio_pago", None),
    ],
    "Activos": [
        ("Fecha_compra", 13, "date", None, None),
        ("Activo", 34, "in", None, None),
        ("Categoria", 18, "in", "Categoria_activo", None),
        ("Cantidad", 10, "in", None, None),
        ("Valor_unitario", 15, "money", None, None),
        ("Vida_util_meses", 14, "in", None, "Referencia: equipos 60, carpa 36, utensilios 12."),
        ("Valor_residual", 14, "money", None, "Valor de venta al final de la vida útil (puede ser 0)."),
        ("Estado", 10, "in", "Estado_activo", None),
        ("Valor_total", 14, "f", None, "Fórmula. No escribir."),
        ("Depreciacion_mensual", 18, "f", None, "Fórmula (línea recta). No escribir."),
    ],
    "Financiamiento": [
        ("Fecha", 12, "date", None, "Fecha de desembolso o del aporte."),
        ("Fuente", 22, "in", None, None),
        ("Tipo_fuente", 24, "in", "Tipo_fuente", None),
        ("Monto", 14, "money", None, None),
        ("Tasa_mensual", 12, "pct", None, "Tasa MENSUAL (ej. 2,5%). 0% si no cobra interés."),
        ("Plazo_meses", 11, "in", None, None),
        ("Cuota_mensual", 14, "f", None, "Fórmula (cuota fija). No escribir."),
        ("Saldo_actual", 14, "money", None, "Lo que aún se debe hoy."),
        ("Destino", 28, "in", None, None),
        ("Estado", 10, "in", "Estado_deuda", None),
    ],
    "Retiros_Dueno": [
        ("Fecha", 12, "date", None, "Fecha en que el dueño saca dinero del negocio para sus gastos personales."),
        ("Valor", 14, "money", None, None),
        ("Observacion", 34, "in", None, None),
    ],
}

FORMULAS = {
    ("Ventas", "Precio_unitario"): '=IF(B{r}="","",IFERROR(INDEX(Productos!$B$2:$B$50,MATCH(B{r},Productos!$A$2:$A$50,0)),""))',
    ("Ventas", "Total"): '=IF(OR(C{r}="",D{r}=""),"",C{r}*D{r})',
    ("Compras_MP", "Total"): '=IF(OR(F{r}="",H{r}=""),"",F{r}*H{r})',
    ("Activos", "Valor_total"): '=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})',
    ("Activos", "Depreciacion_mensual"): '=IF(OR(I{r}="",F{r}="",F{r}=0),"",(I{r}-N(G{r}))/F{r})',
    ("Financiamiento", "Cuota_mensual"): '=IF(OR(D{r}="",F{r}="",F{r}=0),"",IF(N(E{r})>0,-PMT(E{r},F{r},D{r}),D{r}/F{r}))',
}
FMT = {"date": "dd/mm/yyyy", "money": '$#,##0;($#,##0);"-"', "money_f": '$#,##0;($#,##0);"-"', "pct": "0.00%"}

# Parámetros del análisis (clave, valor, formato, descripción)
PARAMETROS = [
    ("Fecha_inicio_base", date(2026, 9, 24), "date", "Desde esta fecha se toman los datos para simular (desde aquí el negocio opera como hoy: 4 hamburguesas, jueves a sábado + domingos antes de festivo)."),
    ("Caja_actual", '=SUM(Ventas!G2:G3000)-SUM(Compras_MP!K2:K3000)-SUM(Costos_Operativos!E2:E3000)-SUM(Retiros_Dueno!B2:B3000)+SUM(Financiamiento!D2:D3000)-SUM(Activos!I2:I3000)-(SUMIFS(Financiamiento!D2:D3000,Financiamiento!C2:C3000,"<>Aporte propio")-SUM(Financiamiento!H2:H3000))', "money", "CALCULADA con el flujo acumulado: ventas − compras de MP − costos operativos − retiros del dueño + préstamos y aportes recibidos − activos comprados − abonos a capital de las deudas. Si cuenta la caja real, escriba ese valor encima de la fórmula."),
    ("Retiro_min", 2000000, "money", "Retiro mensual mínimo del dueño (dato del dueño)."),
    ("Retiro_max", 2500000, "money", "Retiro mensual máximo del dueño (dato del dueño)."),
    ("Choque_mensual", 0.05, "pct", "Variación adicional entre meses (σ). SUPUESTO: 5 % mientras haya menos de 3 meses completos en el periodo base."),
    ("Tasa_EA", 0.42, "pct", "Tasa efectiva anual del crédito. EJEMPLO: reemplazar con la oferta de la entidad."),
    ("Plazo_meses", 24, "int", "Plazo del crédito."),
    ("Gracia_meses", 0, "int", "Meses de gracia a capital (solo se pagan intereses)."),
    ("Comision_inicial", 0.03, "pct", "Comisiones y garantía (FNG / MiPyme) cobradas al desembolsar, % del monto. EJEMPLO."),
    ("Seguro_mensual", 0.001, "pct", "Seguro de vida deudor, % mensual sobre el saldo. EJEMPLO."),
    ("DSCR_min", 1.2, "dec", "Cobertura mínima: flujo disponible / cuotas."),
    ("Alfa_riesgo", 0.05, "pct", "Riesgo aceptado de que un mes no se cubra la cuota con la cobertura mínima."),
    ("Umbral_aprobacion", 0.96, "pct", "Probabilidad mínima de flujo positivo (criterio del curso)."),
    ("SMMLV", 1750905, "money", "Salario mínimo 2026 (Portafolio, 2026). Define la modalidad del crédito."),
    ("Usura_popular_productivo", 0.8844, "pct", "Tasa de usura nov. 2026, popular productivo urbano (≤ 6 SMMLV). La República, 11/11/2026."),
    ("Usura_productivo", 0.5987, "pct", "Tasa de usura nov. 2026, productivo urbano (> 6 y ≤ 25 SMMLV)."),
    ("Usura_mayor_monto", 0.4235, "pct", "Tasa de usura nov. 2026, productivo de mayor monto (> 25 y ≤ 120 SMMLV)."),
    ("N_simulaciones", 10000, "int", "Número de escenarios Monte Carlo."),
    ("Semilla", 2026, "int", "Semilla para reproducir los resultados."),
]

# Escenarios de destino del préstamo (todos los efectos son SUPUESTOS editables)
ESC_COLS = [("Destino", 34), ("Inversion", 14), ("Ventas_min", 11), ("Ventas_probable", 14), ("Ventas_max", 11),
            ("Cambio_CV_pp", 13), ("Costo_fijo_adicional_mes", 22), ("Cambio_costos_jornada", 20),
            ("Meses_hasta_efecto", 17), ("Paga_deudas_vigentes", 19), ("Descripcion", 70)]
ESCENARIOS = [
    ["Solo capacidad de pago (sin destino)", 0, 0, 0, 0, 0, 0, 0, 0, "No",
     "Calcula cuánto podría pagar el negocio tal como opera hoy."],
    ["Activos para la operación", 4500000, 0.03, 0.08, 0.15, -0.01, 50000, 0.05, 1, "No",
     "Segunda plancha, nevera vertical y toldo: atender más rápido en horas pico y reducir merma. Más gas y mantenimiento."],
    ["Expandir a un local", 25000000, -0.10, 0.45, 0.80, 0.0, 3000000, 0.40, 3, "No",
     "Adecuación, mobiliario y depósito del arriendo. Arriendo y servicios del local; más días y personal. Riesgo de perder clientes del sector."],
    ["Capital de trabajo (compras al por mayor)", 3000000, 0, 0, 0, -0.015, 0, 0, 0, "No",
     "Comprar insumos por volumen para bajar el costo unitario."],
    ["Pagar deuda familiar + capital de trabajo", 2500000, 0, 0, 0, -0.01, 0, 0, 0, "Sí",
     "Cancelar el saldo con el familiar y reforzar compras."],
    ["Personalizado", 0, 0, 0, 0, 0, 0, 0, 0, "No", "Defina usted los efectos."],
]


def hdr(ws, cols):
    for j, c_ in enumerate(cols, start=1):
        name, width = c_[0], c_[1]
        com = c_[4] if len(c_) > 4 else None
        c = ws.cell(row=1, column=j, value=name)
        c.font = Font(name=FONT, bold=True, color=HDR_FONT)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        if com:
            c.comment = Comment(com, "La H Burger")
        ws.column_dimensions[c.column_letter].width = width
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30


def build(path, data=None, simulado=False):
    wb = Workbook()
    ins = wb.active
    ins.title = "Instrucciones"

    wp = wb.create_sheet("Productos")
    hdr(wp, [("Producto", 22), ("Precio_venta", 15), ("Activo", 10)])
    for i, (p, pr) in enumerate(PRODUCTOS, 2):
        for j, v in enumerate([p, pr, "Sí"], 1):
            c = wp.cell(row=i, column=j, value=v)
            c.font = Font(name=FONT, color="0000FF"); c.fill = IN_FILL; c.border = BORDER
        wp.cell(row=i, column=2).number_format = FMT["money"]
    wp.cell(row=8, column=1, value="Fuente de precios: documento 'Contexto de la empresa' (dato del dueño).").font = Font(name=FONT, italic=True, size=9)

    wl = wb.create_sheet("Listas")
    ranges = {}
    for j, (k, vals) in enumerate(LISTAS.items(), 1):
        c = wl.cell(row=1, column=j, value=k); c.font = Font(name=FONT, bold=True)
        for i, v in enumerate(vals, 2):
            wl.cell(row=i, column=j, value=v).font = Font(name=FONT)
        col = c.column_letter
        ranges[k] = f"Listas!${col}$2:${col}${len(vals) + 1 + 10}"
        wl.column_dimensions[col].width = max(14, max(len(x) for x in vals) + 2)

    for sname, cols in SHEETS.items():
        ws = wb.create_sheet(sname)
        hdr(ws, cols)
        for j, (name, width, kind, lst, com) in enumerate(cols, start=1):
            L = ws.cell(row=1, column=j).column_letter
            if lst:
                dv = DataValidation(type="list", formula1="=" + ranges[lst], allow_blank=True, showErrorMessage=True,
                                    errorTitle="Valor no válido", error="Elija un valor de la lista (hoja Listas).")
                ws.add_data_validation(dv); dv.add(f"{L}2:{L}{MAXR}")
            if kind == "date":
                dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2015,1,1)", allow_blank=True,
                                    showErrorMessage=True, error="Escriba una fecha válida (dd/mm/aaaa).")
                ws.add_data_validation(dv); dv.add(f"{L}2:{L}{MAXR}")
            if kind == "money" or name in ("Cantidad", "Vida_util_meses", "Plazo_meses"):
                dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                                    showErrorMessage=True, error="Debe ser un número mayor o igual a 0.")
                ws.add_data_validation(dv); dv.add(f"{L}2:{L}{MAXR}")
        rows = (data or {}).get(sname, [])
        last = max(len(rows) + 1, 300)
        for r in range(2, last + 1):
            rec = rows[r - 2] if r - 2 < len(rows) else None
            for j, (name, width, kind, lst, com) in enumerate(cols, start=1):
                c = ws.cell(row=r, column=j)
                c.font = Font(name=FONT, color="0000FF" if kind not in ("f", "money_f") else "000000")
                c.border = BORDER
                if kind in FMT:
                    c.number_format = FMT[kind]
                if (sname, name) in FORMULAS:
                    if rec is not None and kind == "money_f" and rec[j - 1] is not None:
                        c.value = rec[j - 1]; c.fill = IN_FILL; c.font = Font(name=FONT, color="0000FF")
                    else:
                        c.value = FORMULAS[(sname, name)].format(r=r)
                        c.fill = F_FILL if kind == "f" else IN_FILL
                    if kind == "f":
                        c.number_format = FMT["money"]
                else:
                    c.fill = IN_FILL
                    if rec is not None:
                        c.value = rec[j - 1]
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(cols)).column_letter}{last}"

    # Escenarios del préstamo
    we = wb.create_sheet("Escenarios_Prestamo")
    hdr(we, ESC_COLS)
    fm = {"Inversion": FMT["money"], "Costo_fijo_adicional_mes": FMT["money"], "Ventas_min": "0.0%", "Ventas_probable": "0.0%",
          "Ventas_max": "0.0%", "Cambio_costos_jornada": "0.0%", "Cambio_CV_pp": "0.0%"}
    dv = DataValidation(type="list", formula1="=" + ranges["Si_No"], allow_blank=True); we.add_data_validation(dv); dv.add("J2:J20")
    for i, row in enumerate(ESCENARIOS, 2):
        for j, v in enumerate(row, 1):
            c = we.cell(row=i, column=j, value=v)
            c.font = Font(name=FONT, color="0000FF"); c.fill = IN_FILL; c.border = BORDER
            if ESC_COLS[j - 1][0] in fm:
                c.number_format = fm[ESC_COLS[j - 1][0]]
            if j == len(row):
                c.alignment = Alignment(wrap_text=True, vertical="top")
    n = len(ESCENARIOS) + 3
    notas = ["Todos los efectos son SUPUESTOS editables. Ventas_min / probable / max: cambio % de las ventas (distribución triangular, se sortea una vez por escenario simulado).",
             "Cambio_CV_pp: puntos porcentuales que sube (+) o baja (−) el costo de materia prima sobre ventas. Ej.: −0,01 = baja 1 punto.",
             "Costo_fijo_adicional_mes: costo fijo nuevo por mes (ej. arriendo del local). Cambio_costos_jornada: % de cambio en ayudantes y gas.",
             "Meses_hasta_efecto: meses desde el desembolso hasta que el cambio empieza. Paga_deudas_vigentes = Sí: el préstamo cancela las deudas vigentes."]
    for k, t in enumerate(notas):
        we.cell(row=n + k, column=1, value=t).font = Font(name=FONT, italic=True, size=9)

    # Parámetros
    wpar = wb.create_sheet("Parametros")
    hdr(wpar, [("Parametro", 26), ("Valor", 16), ("Descripcion", 100)])
    for i, (k, v, f, d) in enumerate(PARAMETROS, 2):
        wpar.cell(row=i, column=1, value=k).font = Font(name=FONT, bold=True)
        c = wpar.cell(row=i, column=2, value=v); c.font = Font(name=FONT, color="000000" if str(v).startswith("=") else "0000FF")
        c.fill = F_FILL if str(v).startswith("=") else IN_FILL; c.border = BORDER
        c.number_format = {"date": "dd/mm/yyyy", "money": FMT["money"], "pct": "0.00%", "dec": "0.00", "int": "0"}[f]
        wpar.cell(row=i, column=3, value=d).font = Font(name=FONT, size=9)

    # Instrucciones
    ins.column_dimensions["A"].width = 24
    ins.column_dimensions["B"].width = 100
    title = "La H Burger · Registro financiero para el simulador Monte Carlo"
    if simulado:
        title += "  (DATOS SIMULADOS)"
    ins["A1"] = title; ins["A1"].font = Font(name=FONT, bold=True, size=14)
    lines = [
        ("Cómo usar", "Llene las celdas de color crema (texto azul). Las grises tienen fórmulas. No cambie nombres de hojas ni de columnas: las apps y Colab los usan."),
        ("Productos", "Catálogo y precio de venta de cada hamburguesa."),
        ("Ventas", "Una fila por producto, día y medio de pago. Mínimo: Fecha, Producto y Cantidad."),
        ("Compras_MP", "Cada compra de materia prima con su detalle (insumo, marca, proveedor, cantidad, unidad, precio, vencimiento). El costo variable se calcula en general: compras del mes / ventas del mes."),
        ("Costos_Operativos", "Tipo_costo: Fijo mensual (espacio, servicios), Por jornada (ayudantes, gas) o Variable (% ventas)."),
        ("Activos", "Equipos y su vida útil. La depreciación no es salida de caja."),
        ("Financiamiento", "Aportes y préstamos (incluidos familiares). Las deudas Vigentes se restan del flujo hasta que se termine el Saldo_actual."),
        ("Retiros_Dueno", "Lo que el dueño saca cada mes para sus gastos. Sirve para calcular la caja actual."),
        ("Escenarios_Prestamo", "Destinos posibles del préstamo y su efecto esperado en ventas y costos (supuestos editables)."),
        ("Parametros", "Supuestos del análisis y condiciones del crédito. Reemplace los valores marcados como EJEMPLO con la oferta real del banco."),
    ]
    r = 3
    for a, b in lines:
        ins.cell(row=r, column=1, value=a).font = Font(name=FONT, bold=True)
        cb = ins.cell(row=r, column=2, value=b); cb.font = Font(name=FONT); cb.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    if simulado:
        r += 1
        ins.cell(row=r, column=1, value="AVISO").font = Font(name=FONT, bold=True, color="C00000")
        cb = ins.cell(row=r, column=2, value=("Datos SIMULADOS con las reglas del negocio indicadas por el dueño (fases de productos, días de apertura, "
                                             "rangos de hamburguesas por fin de semana, ayudantes, espacio, gas, activos y deuda familiar). "
                                             "Reemplácelos por los registros reales cuando estén disponibles. Generados con generar_datos.py."))
        cb.font = Font(name=FONT, color="C00000"); cb.alignment = Alignment(wrap_text=True)
    wb.move_sheet("Listas", offset=20)
    wb.save(path)


# ---------------------------------------------------------------------------- simulación de datos
FESTIVO_DOMINGOS = {date(2026, 10, 11), date(2026, 11, 1), date(2026, 11, 15)}  # domingos antes de lunes festivo
INICIO, FIN = date(2026, 5, 15), date(2026, 11, 15)
F2, F3, JUEVES = date(2026, 6, 5), date(2026, 9, 18), date(2026, 9, 24)
PRECIOS = dict(PRODUCTOS)

# Receta aproximada SOLO para simular compras (el negocio no tiene cantidades estandarizadas)
# insumo: (categoría, marca, proveedor, unidad, precio base, vida útil días, {producto: cantidad por hamburguesa}, unidad mínima de compra)
TODAS = {p: 1 for p, _ in PRODUCTOS}
INSUMOS = {
    "Pan artesano": ("Pan", "Panadería artesanal", "Panadería El Trigal", "und", 1500, 4, {"La Polluda": 1}, 1),
    "Pan brioche": ("Pan", "Panadería artesanal", "Panadería El Trigal", "und", 1300, 4, {"La Paisa": 1, "Paisa Ardiente": 1}, 1),
    "Pan pretzel": ("Pan", "Panadería artesanal", "Panadería El Trigal", "und", 1900, 5, {"La Nuevona": 1}, 1),
    "Filete de pechuga porcionado": ("Proteína", "Pollos El Bucanero", "Carnicería Campo Valdés", "kg", 21000, 4, {"La Polluda": 0.15}, 0.5),
    "Chorizo artesanal para hamburguesa": ("Tocineta y embutidos", "Artesanal", "Salsamentaría La Paisita", "kg", 25000, 6, {"La Paisa": 0.13}, 0.5),
    "Chorizo picante artesanal": ("Tocineta y embutidos", "Artesanal", "Salsamentaría La Paisita", "kg", 27000, 6, {"Paisa Ardiente": 0.13}, 0.5),
    "Carne de res porcionada 150 g": ("Proteína", "Res fresca", "Carnicería Campo Valdés", "kg", 30000, 3, {"La Nuevona": 0.15}, 0.5),
    "Tocineta": ("Tocineta y embutidos", "Zenú", "Mayorista Central", "kg", 39000, 20, {"La Paisa": 0.03, "Paisa Ardiente": 0.03, "La Nuevona": 0.035, "La Polluda": 0.025}, 0.5),
    "Queso mozzarella": ("Lácteos y quesos", "Colanta", "Mayorista Central", "kg", 29000, 20, {"La Paisa": 0.035, "Paisa Ardiente": 0.035}, 0.5),
    "Queso para asar": ("Lácteos y quesos", "Colanta", "Mayorista Central", "kg", 24000, 15, {"La Nuevona": 0.045}, 0.5),
    "Queso crema": ("Lácteos y quesos", "Alpina", "Mayorista Central", "kg", 23000, 25, {"La Polluda": 0.03}, 0.5),
    "Aguacate Hass": ("Frutas y verduras", "Hass", "Plaza Minorista", "kg", 9500, 4, {"La Paisa": 0.045, "La Polluda": 0.045, "Paisa Ardiente": 0.055}, 0.5),
    "Cebolla cabezona": ("Frutas y verduras", "Sin marca", "Plaza Minorista", "kg", 3500, 12, {"La Paisa": 0.04, "Paisa Ardiente": 0.04}, 0.5),
    "Repollo": ("Frutas y verduras", "Sin marca", "Plaza Minorista", "kg", 2500, 8, {"La Nuevona": 0.04}, 0.5),
    "Piña": ("Frutas y verduras", "Oro miel", "Plaza Minorista", "kg", 4500, 5, {"La Nuevona": 0.03}, 0.5),
    "Mayonesa (base salsa de la casa)": ("Salsas y aderezos", "Fruco", "Mayorista Central", "L", 15500, 60, {**TODAS, **{k: 0.025 for k in TODAS}}, 1),
    "Salsa BBQ y miel (reducción)": ("Salsas y aderezos", "San Jorge", "Mayorista Central", "L", 16000, 60, {"La Nuevona": 0.015}, 1),
    "Harina de trigo": ("Insumos de apanado", "Haz de Oros", "Mayorista Central", "kg", 4200, 90, {"La Polluda": 0.02}, 1),
    "Huevos": ("Insumos de apanado", "Kikes", "Mayorista Central", "cubeta", 17000, 20, {"La Polluda": 0.3 / 30}, 1),
    "Apanadura panko": ("Insumos de apanado", "Panko", "Mayorista Central", "kg", 14000, 90, {"La Polluda": 0.03}, 1),
    "Papa a la francesa precocida": ("Papas y acompañamientos", "McCain", "Mayorista Central", "kg", 9000, 60, {**{k: 0.12 for k in TODAS}}, 2.5),
    "Aceite vegetal": ("Aceite y grasas", "Premier", "Mayorista Central", "L", 9800, 180, {**{k: 0.025 for k in TODAS}}, 3),
    "Caja kraft para hamburguesa": ("Empaques y desechables", "Darnel", "Desechables Medellín", "und", 650, 365, {**TODAS}, 50),
    "Servilletas (paquete x 300)": ("Empaques y desechables", "Familia", "Desechables Medellín", "paquete", 6500, 365, {**{k: 3 / 300 for k in TODAS}}, 1),
    "Bolsa kraft para llevar": ("Empaques y desechables", "Darnel", "Desechables Medellín", "und", 250, 365, {**{k: 0.5 for k in TODAS}}, 50),
}
PAGOS = (["Efectivo", "Nequi", "Daviplata", "Tarjeta"], [0.42, 0.38, 0.08, 0.12])


def simular_datos(seed=2026):
    rng = np.random.default_rng(seed)
    ventas = []
    unidades_por_dia = {}  # fecha -> {producto: unidades}
    d = INICIO
    while d <= FIN:
        wd = d.weekday()  # 0 lun ... 4 vie, 5 sáb, 6 dom, 3 jue
        if wd == 4:  # viernes: decide el total del fin de semana
            if d < F2:
                total, mix = rng.integers(80, 101), {"La Paisa": 1.0}
            elif d < F3:
                total = rng.integers(100, 151)
                s = np.clip(rng.normal(0.55, 0.05), 0.4, 0.7)
                mix = {"La Paisa": s, "La Polluda": 1 - s}
            else:
                total = rng.integers(150, 201)
                w = rng.dirichlet(np.array([0.38, 0.34, 0.14, 0.14]) * 60)
                mix = dict(zip(["La Paisa", "La Polluda", "La Nuevona", "Paisa Ardiente"], w))
            share_vie = rng.uniform(0.42, 0.50)
            n_vie = int(round(total * share_vie))
            for dia, n in ((d, n_vie), (d + timedelta(days=1), int(total) - n_vie)):
                if dia <= FIN:
                    unidades_por_dia[dia] = dict(zip(mix.keys(), rng.multinomial(n, list(mix.values()))))
        if (wd == 3 and d >= JUEVES) or (wd == 6 and d in FESTIVO_DOMINGOS):
            n = int(rng.integers(35, 51))
            w = rng.dirichlet(np.array([0.38, 0.34, 0.14, 0.14]) * 60)
            unidades_por_dia[d] = dict(zip(["La Paisa", "La Polluda", "La Nuevona", "Paisa Ardiente"], rng.multinomial(n, w)))
        d += timedelta(days=1)

    for dia in sorted(unidades_por_dia):
        obs = "Domingo antes de festivo" if dia in FESTIVO_DOMINGOS else None
        for prod, q in unidades_por_dia[dia].items():
            if q <= 0:
                continue
            for medio, k in zip(PAGOS[0], rng.multinomial(int(q), PAGOS[1])):
                if k > 0:
                    ventas.append([dia, prod, int(k), None, medio, obs, None])

    # ---- compras de materia prima: compra grande el día antes de abrir cada semana + reposición el sábado
    dias = sorted(unidades_por_dia)
    semanas = {}
    for dia in dias:
        lunes = dia - timedelta(days=dia.weekday())
        semanas.setdefault(lunes, []).append(dia)
    compras = []
    inventario = {k: 0.0 for k in INSUMOS}  # sobrantes de no perecederos
    n_fact = 1000
    for lunes, ds in sorted(semanas.items()):
        primer = ds[0]
        necesidad = {k: 0.0 for k in INSUMOS}
        for dia in ds:
            for prod, q in unidades_por_dia[dia].items():
                for ins_, spec in INSUMOS.items():
                    necesidad[ins_] += spec[6].get(prod, 0) * q
        merma = rng.uniform(1.03, 1.10)
        semanas_desde_inicio = (lunes - INICIO).days / 7
        inflacion = (1 + 0.0012) ** semanas_desde_inicio  # ≈ 0,5 % mensual
        # perecederos (vida ≤ 6 días): mitad el día antes de abrir, resto el sábado; resto: una compra semanal
        for ins_, (cat, marca, prov, uni, precio, vida, rec, minimo) in INSUMOS.items():
            req = necesidad[ins_] * merma - inventario[ins_]
            if req <= 0:
                inventario[ins_] -= necesidad[ins_] * merma
                continue
            perecedero = vida <= 6
            partes = [(primer - timedelta(days=1), 0.6), (max(ds), 0.4)] if (perecedero and len(ds) >= 2) else [(primer - timedelta(days=1), 1.0)]
            if perecedero and len(ds) >= 2:
                sab = [x for x in ds if x.weekday() == 5]
                partes = [(primer - timedelta(days=1), 0.6), ((sab[0] if sab else ds[-1]), 0.4)]
            comprado = 0.0
            for fecha, frac in partes:
                q = req * frac
                q = max(minimo, np.ceil(q / minimo) * minimo) if uni not in ("und",) else max(minimo, int(np.ceil(q / minimo) * minimo))
                p = round(precio * inflacion * rng.uniform(0.96, 1.05), -1)
                n_fact += 1
                compras.append([fecha, cat, ins_, marca, prov, float(round(q, 2)), uni, float(p), fecha + timedelta(days=vida), f"F-{n_fact}", None])
                comprado += q
            inventario[ins_] = 0.0 if perecedero else max(0.0, inventario[ins_] + comprado - necesidad[ins_] * merma)

    # ---- costos operativos
    costos = []
    jornadas = 0
    for dia in dias:
        wd = dia.weekday()
        n_ay = 3 if wd in (4, 5) else 2
        costos.append([dia, "Nómina / ayudantes", f"{n_ay} ayudantes × $100.000", "Por jornada", n_ay * 100000, "Efectivo"])
    m = date(2026, 5, 15)
    while m <= FIN:
        costos.append([m, "Arriendo / uso de espacio", "Espacio y servicios (mes)", "Fijo mensual", 800000, "Transferencia"])
        costos.append([m, "Gas / combustible", "Pipeta de gas 40 lb (mes)", "Fijo mensual", 100000, "Efectivo"])
        m = date(m.year + (m.month // 12), m.month % 12 + 1, 15)
    costos.sort(key=lambda r: r[0])

    activos = [
        [date(2026, 5, 9), "Plancha industrial a gas (2 puestos)", "Equipo de cocina", 1, 900000, 60, 100000, "Bueno", None, None],
        [date(2026, 5, 9), "Freidora de papas a gas", "Equipo de cocina", 1, 550000, 60, 50000, "Bueno", None, None],
        [date(2026, 5, 10), "Carpa 3 × 3 m", "Carpa / estructura", 1, 450000, 36, 0, "Bueno", None, None],
        [date(2026, 5, 10), "Mesa de trabajo en acero", "Mobiliario", 1, 300000, 48, 30000, "Bueno", None, None],
        [date(2026, 5, 11), "Nevera portátil (cooler)", "Refrigeración", 2, 90000, 24, 0, "Bueno", None, None],
        [date(2026, 5, 11), "Utensilios (espátulas, pinzas, cuchillos, recipientes)", "Utensilios", 1, 250000, 12, 0, "Regular", None, None],
        [date(2026, 5, 12), "Pipeta de gas 40 lb (envase) y regulador", "Equipo de cocina", 1, 220000, 60, 50000, "Bueno", None, None],
        [date(2026, 5, 12), "Dispensadores de salsa y termos", "Utensilios", 1, 150000, 24, 0, "Bueno", None, None],
    ]
    fin = [[date(2026, 5, 8), "Familiar", "Familiar / amigos", 3000000, 0, 9, None, 1000000, "Compra de activos para iniciar (plancha, freidora, carpa, otros)", "Vigente"]]
    retiros = [[date(2026, 5, 31), 1100000, "Medio mes (inició el 15/05)"]]
    for mes_fin in [date(2026, 6, 30), date(2026, 7, 31), date(2026, 8, 31), date(2026, 9, 30), date(2026, 10, 31)]:
        retiros.append([mes_fin, float(round(rng.uniform(2000000, 2500000), -4)), None])
    return {"Ventas": ventas, "Compras_MP": compras, "Costos_Operativos": costos, "Activos": activos, "Financiamiento": fin,
            "Retiros_Dueno": retiros}


if __name__ == "__main__":
    out = sys.argv[1]
    seed = int(sys.argv[2]) if len(sys.argv) > 2 else 2026
    build(f"{out}/Plantilla_Registro_LaHBurger.xlsx")
    build(f"{out}/Datos_Simulados_LaHBurger.xlsx", simular_datos(seed), simulado=True)
    print("ok")
