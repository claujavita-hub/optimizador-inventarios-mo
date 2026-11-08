"""Tablero de solvencia y liquidez + impacto de las acciones de inventario.

Lee el tablero (hoja 'Tablero' con los datos de entrada por período y hoja
'Umbrales' con los límites del semáforo), recalcula los indicadores con las
mismas fórmulas del Excel y simula cómo cambian si se libera capital de
trabajo reduciendo inventario. Cifras en las unidades del tablero (US$ mm).

Supuesto de la simulación: el inventario que se reduce se convierte en caja
al valor por libra indicado. Esa caja puede quedarse en caja o usarse para
pagar deuda; la deuda neta baja lo mismo en ambos casos, pero cambian la
liquidez corriente, la razón ácida y el pasivo.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl import load_workbook

from .loader import to_float
from .schema import norm

INPUTS = {
    "ebitda": ("ebitda_ultimos_12_meses", "ebitda"),
    "gastos_fin": ("gastos_financieros_netos",),
    "deuda": ("deuda_financiera_total",),
    "deuda_cp": ("deuda_financiera_corto_plazo",),
    "caja": ("caja_y_depositos", "caja"),
    "activo_corriente": ("activo_corriente",),
    "inventario": ("inventario_corriente", "inventario"),
    "pasivo_corriente": ("pasivo_corriente",),
    "pasivo_total": ("pasivo_total",),
    "patrimonio": ("patrimonio",),
    "costo_ventas": ("costo_de_ventas",),
    "flujo_operacional": ("flujo_operacional",),
    "capex": ("capex",),
    "dividendos": ("dividendos",),
}

# nombre visible, clave, "mayor"/"menor" es mejor, verde, amarillo (defaults = hoja Umbrales entregada)
INDICATORS = [
    ("Deuda neta / EBITDA", "nd_ebitda", "menor", 2.5, 3.0),
    ("EBITDA / gastos financieros netos", "cobertura", "mayor", 4.0, 3.0),
    ("Flujo libre tras capex y dividendos", "flujo_libre", "mayor", 0.0, -50.0),
    ("Liquidez corriente", "liquidez", "mayor", 1.5, 1.0),
    ("Razón ácida", "acida", "mayor", 1.0, 0.7),
    ("Caja / deuda de corto plazo", "caja_cp", "mayor", 1.5, 1.0),
    ("Pasivo total / patrimonio", "pasivo_patrimonio", "menor", 1.0, 1.5),
    ("Días de inventario", "dias_inventario", "menor", 150.0, 180.0),
    ("Holgura de EBITDA hasta covenant", "holgura", "mayor", 0.30, 0.15),
]
FMT = {"flujo_libre": "{:,.1f}", "holgura": "{:.1%}", "dias_inventario": "{:.0f}"}


@dataclass
class Threshold:
    name: str
    key: str
    better: str
    green: float
    yellow: float

    def color(self, v: float | None) -> str:
        if v is None:
            return "s/d"
        if self.better == "mayor":
            return "Verde" if v >= self.green else "Amarillo" if v >= self.yellow else "Rojo"
        return "Verde" if v <= self.green else "Amarillo" if v <= self.yellow else "Rojo"


@dataclass
class Tablero:
    periods: dict[str, dict[str, float | None]]     # período -> datos de entrada
    thresholds: list[Threshold]
    covenant: float = 3.0
    warnings: list[str] = field(default_factory=list)

    @property
    def latest(self) -> str | None:
        ok = [p for p, v in self.periods.items() if v.get("ebitda") is not None and v.get("deuda") is not None]
        return ok[-1] if ok else None


def _label_key(text) -> str | None:
    t = norm(text)
    for key, aliases in INPUTS.items():
        if any(t.startswith(a) for a in aliases):
            return key
    return None


def is_tablero_workbook(source) -> bool:
    try:
        return load_tablero(source) is not None
    except Exception:
        return False


def load_tablero(source) -> Tablero | None:
    wb = load_workbook(source, data_only=True, read_only=True)
    try:
        sheet = None
        for name in wb.sheetnames:
            rows = list(wb[name].iter_rows(values_only=True))
            labels = {_label_key(r[0]) for r in rows if r and r[0]}
            if {"ebitda", "deuda", "caja", "inventario"} <= labels:
                sheet = rows
                break
        if sheet is None:
            return None
        # encabezado de períodos: fila que empieza con "Datos de entrada" (o la fila previa a EBITDA)
        hdr_i = next((i for i, r in enumerate(sheet) if r and r[0] and norm(r[0]).startswith("datos_de_entrada")), None)
        if hdr_i is None:
            hdr_i = next(i for i, r in enumerate(sheet) if r and _label_key(r[0]) == "ebitda") - 1
        hdr = sheet[hdr_i]
        cols = [(j, str(h)) for j, h in enumerate(hdr) if j > 0 and h not in (None, "")]
        periods: dict[str, dict[str, float | None]] = {p: {} for _, p in cols}
        for r in sheet[hdr_i + 1:]:
            if not r or not r[0]:
                continue
            key = _label_key(r[0])
            if key is None or key in next(iter(periods.values()), {}):
                if key is None and norm(r[0]).startswith("indicadores"):
                    break
                continue
            for j, p in cols:
                periods[p][key] = to_float(r[j]) if j < len(r) else None
        tab = Tablero(periods, [Threshold(*i) for i in INDICATORS])
        # umbrales y covenant
        for name in wb.sheetnames:
            if "umbral" not in norm(name):
                continue
            for r in wb[name].iter_rows(values_only=True):
                if not r or not r[0]:
                    continue
                t = norm(r[0])
                if t.startswith("covenant") and len(r) > 2:
                    cov = to_float(r[2]) if to_float(r[2]) is not None else to_float(r[1])
                    if cov:
                        tab.covenant = cov
                for th in tab.thresholds:
                    if t == norm(th.name) and len(r) > 3:
                        g, y = to_float(r[2]), to_float(r[3])
                        if g is not None and y is not None:
                            th.green, th.yellow = g, y
                        if r[1]:
                            th.better = "mayor" if "mayor" in norm(r[1]) else "menor"
        return tab
    finally:
        wb.close()
        if hasattr(source, "seek"):
            source.seek(0)


def indicators(d: dict[str, float | None], covenant: float, release: float = 0.0,
               use: str = "caja") -> dict[str, float | None]:
    """Indicadores del tablero, opcionalmente liberando `release` (US$ mm) de inventario.

    use = "caja": la caja liberada queda en caja. use = "deuda": se usa para pagar deuda
    (primero la de corto plazo)."""
    g = lambda k: d.get(k)  # noqa: E731
    if None in (g("ebitda"), g("deuda"), g("caja")):
        return {k: None for _, k, *_ in INDICATORS}
    deuda, caja, deuda_cp = g("deuda"), g("caja"), g("deuda_cp") or 0.0
    ac, pc, inv = g("activo_corriente"), g("pasivo_corriente"), g("inventario")
    pasivo = g("pasivo_total")
    inv_x = inv - release if inv is not None else None
    if use == "deuda":
        pago_cp = min(release, deuda_cp)
        deuda, deuda_cp = deuda - release, deuda_cp - pago_cp
        ac_x = ac - release if ac is not None else None          # inventario sale y la caja se usa en pagar
        pc_x = pc - pago_cp if pc is not None else None
        pasivo_x = pasivo - release if pasivo is not None else None
    else:
        caja = caja + release
        ac_x, pc_x, pasivo_x = ac, pc, pasivo                     # inventario se transforma en caja
    nd = deuda - caja
    ebitda = g("ebitda")
    out = {
        "nd_ebitda": nd / ebitda if ebitda else None,
        "cobertura": ebitda / g("gastos_fin") if g("gastos_fin") else None,
        "flujo_libre": (g("flujo_operacional") - (g("capex") or 0) - (g("dividendos") or 0) + release
                        if g("flujo_operacional") is not None else None),
        "liquidez": ac_x / pc_x if ac_x is not None and pc_x else None,
        "acida": (ac_x - inv_x) / pc_x if None not in (ac_x, inv_x) and pc_x else None,
        "caja_cp": caja / deuda_cp if deuda_cp else None,
        "pasivo_patrimonio": pasivo_x / g("patrimonio") if pasivo_x is not None and g("patrimonio") else None,
        "dias_inventario": inv_x / g("costo_ventas") * 365 if inv_x is not None and g("costo_ventas") else None,
        "holgura": 1 - nd / (covenant * ebitda) if ebitda else None,
        "deuda_neta": nd,
    }
    return out


def fmt(key: str, v: float | None) -> str:
    if v is None:
        return "s/d"
    return FMT.get(key, "{:.2f}").format(v)


# ---------------------------------------------------------------- acciones de inventario

@dataclass
class Action:
    code: str
    name: str
    mmlb: float
    status: str              # "Validado" / "Por validar" / "Decisión comercial"
    basis: str
    default: bool

    def usd(self, price: float) -> float:
        return self.mmlb * price


def inventory_actions(r, target_transit: float = 45, target_bodega: float = 20,
                      target_assigned: float = 10) -> list[Action]:
    """Acciones cuantificadas a partir del análisis agregado (AggResult)."""
    from statistics import mean
    data, p = r.data, r.data.params
    acts: list[Action] = []
    # Pre-PT: últimos 3 meses vs días de recorrido x producción propia de esos meses
    pre = r.groups[0] if r.groups else None
    last = [m for m in data.months if m.prept is not None and m.production is not None][-3:]
    if pre is not None and pre.pipeline is not None and len(last) == 3:
        days = pre.pipeline / pre.flow_day if pre.flow_day else 0
        stock = mean(m.prept for m in last)
        prod = mean(m.production - (m.tolling or 0) for m in last) / 30
        exc = stock - prod * days
        if exc > 0:
            acts.append(Action("A2", "Ajustar compras de concentrado a la producción (Pre-PT)", exc, "Validado",
                               f"Pre-PT {last[0].label}–{last[-1].label} {stock:.2f} MMlb vs {days:.0f} días x "
                               f"producción propia {prod * 30:.2f} MMlb/mes = {prod * days:.2f}", True))
    # Por mercado: contra las metas de días de Celonis si el export las trae; si no, días objetivo x venta
    if data.celonis is not None and data.celonis.has_targets:
        names = {"Transit": "tránsito", "Ext.WH": "bodegas"}
        for (m, st), (q, d, t, ex) in sorted(data.celonis.excess_vs_target().items(), key=lambda x: -x[1][3]):
            if st not in names or ex <= 0.05 or norm(m) == "not_assigned":
                continue
            acts.append(Action(f"{'T' if st == 'Transit' else 'B'}-{m}",
                               f"Llevar {names[st]} de {m} a la meta Celonis ({t:.0f} días)", ex,
                               "Validado" if st == "Ext.WH" else "Por validar",
                               f"Celonis: {q:.2f} MMlb a {d:.0f} días vs meta {t:.0f} días (exceso por material)",
                               st == "Ext.WH"))
    elif data.celonis is not None and data.sales_market:
        mk = data.celonis.by("market", "stage")
        for m, v in sorted(data.sales_market.items(), key=lambda x: -x[1]):
            if v <= 0:
                continue
            tr, wh = mk.get((m, "Transit"), 0.0), mk.get((m, "Ext.WH"), 0.0)
            exc_tr = tr - v / 30 * target_transit
            if exc_tr > 0.05:
                acts.append(Action(f"T-{m}", f"Normalizar tránsito a {m}", exc_tr, "Por validar",
                                   f"Tránsito {tr:.2f} MMlb vs {target_transit:.0f} días de venta "
                                   f"({v:.2f} MMlb/mes) = {v / 30 * target_transit:.2f}", False))
            exc_wh = wh - v / 30 * target_bodega
            if exc_wh > 0.05:
                acts.append(Action(f"B-{m}", f"Bajar bodegas de {m} a {target_bodega:.0f} días", exc_wh, "Validado",
                                   f"Bodegas {wh:.2f} MMlb vs {target_bodega:.0f} días de venta "
                                   f"({v:.2f} MMlb/mes) = {v / 30 * target_bodega:.2f}", True))
    # Stock de seguridad por cliente/SKU sobre el teórico
    for row in r.sku_rows:
        if row[10] is not None and row[10] > 0:
            acts.append(Action(f"SS-{row[0][:12]}", f"Ajustar safety stock de {row[0]}", row[10] / 1e6, "Validado",
                               f"SS real {row[8] / 1e6:.2f} vs teórico {row[9] / 1e6:.2f} MMlb", True))
    # Asignado esperando despacho
    asg = [m.assigned for m in data.months if m.assigned is not None]
    if asg and r.d_day:
        exc = mean(asg) - r.d_day * target_assigned
        if exc > 0.05:
            acts.append(Action("B1", f"Despachar en fecha cliente: asignado a {target_assigned:.0f} días", exc,
                               "Por validar", f"Asignado {mean(asg):.2f} MMlb ({mean(asg) / r.d_day:.0f} días) vs "
                               f"{target_assigned:.0f} días = {r.d_day * target_assigned:.2f}", False))
    if p.stock_spot:
        acts.append(Action("SPOT", "Reducir stock para ventas spot", p.stock_spot, "Decisión comercial",
                           f"Stock spot informado {p.stock_spot:.2f} MMlb", False))
    return acts


def scenario_table(tab: Tablero, period: str, release: float, use: str = "caja") -> dict:
    d = tab.periods[period]
    base, sim = indicators(d, tab.covenant), indicators(d, tab.covenant, release, use)
    rows = []
    for th in tab.thresholds:
        b, s = base.get(th.key), sim.get(th.key)
        rows.append([th.name, fmt(th.key, b), th.color(b), fmt(th.key, s), th.color(s)])
    rows.append(["Deuda neta (US$ mm)", fmt("x", base["deuda_neta"]), "", fmt("x", sim["deuda_neta"]), ""])
    return dict(headers=["Indicador", f"{period} actual", "Semáforo", f"Con acciones (US${release:,.0f} mm)", "Semáforo "],
                rows=rows, fmts=[None] * 5)


def thresholds_needed(tab: Tablero, period: str, use: str = "caja") -> list[list]:
    """Cuánto hay que liberar para que cada indicador amarillo/rojo pase a verde (búsqueda simple)."""
    d = tab.periods[period]
    out = []
    for th in tab.thresholds:
        cur = indicators(d, tab.covenant).get(th.key)
        if th.color(cur) in ("Verde", "s/d"):
            continue
        def search(mode):
            for x in range(0, 2001):
                if th.color(indicators(d, tab.covenant, float(x), mode).get(th.key)) == "Verde":
                    return float(x)
            return None
        need = search(use)
        alt = search("deuda") if need is None and use == "caja" else None
        out.append([th.name, fmt(th.key, cur), th.color(cur), need, alt])
    return out


ICON = {"Verde": "🟢", "Amarillo": "🟡", "Rojo": "🔴", "s/d": "⚪", "": ""}


def actions_release(actions: list[Action], value_lb: float) -> float:
    """US$ mm liberados por las acciones (MMlb x US$/lb = US$ mm)."""
    return sum(a.usd(value_lb) for a in actions)


def _no_need(row) -> str:
    if row[4] is not None:
        return f"solo pagando deuda: US$ {row[4]:,.0f} mm"
    return "no se resuelve con inventario"


def write_solvency_report(tab: Tablero, period: str, actions: list[Action], value_lb: float, rate: float,
                          use: str, path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    fills = {"Verde": "C6EFCE", "Amarillo": "FFEB9C", "Rojo": "FFC7CE"}
    release = actions_release(actions, value_lb)
    wb = Workbook()
    ws = wb.active
    ws.title = "Escenario"
    ws.append([f"Impacto de las acciones de inventario en solvencia y liquidez · {period}"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Liberación total US$ {release:,.1f} mm · valor {value_lb:.2f} US$/lb · caja "
               + ("usada para pagar deuda" if use == "deuda" else "mantenida en caja")
               + f" · ahorro de intereses US$ {release * rate:,.1f} mm/año"])
    ws.append([])
    tb = scenario_table(tab, period, release, use)
    ws.append(tb["headers"])
    for row in tb["rows"]:
        ws.append(row)
        for col in (3, 5):
            c = ws.cell(ws.max_row, col)
            if c.value in fills:
                c.fill = PatternFill("solid", fgColor=fills[c.value])
    ws.append([])
    ws.append(["Liberación necesaria para llegar a verde (US$ mm)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append(["Indicador", "Actual", "Semáforo", "US$ mm a liberar"])
    for row in thresholds_needed(tab, period, use):
        ws.append([*row[:3], row[3] if row[3] is not None else _no_need(row)])
    a = wb.create_sheet("Acciones")
    a.append(["Código", "Acción", "MMlb", "US$ mm", "Estado", "Base del cálculo"])
    for x in actions:
        a.append([x.code, x.name, round(x.mmlb, 3), round(x.usd(value_lb), 1), x.status, x.basis])
    for s in (ws, a):
        for col, w in zip("ABCDEF", (44, 16, 12, 22, 14, 80)):
            s.column_dimensions[col].width = w
    wb.save(path)
