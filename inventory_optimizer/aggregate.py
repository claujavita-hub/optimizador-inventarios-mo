"""Modo agregado: diagnóstico de rotación y cobertura óptima por etapa.

Se usa cuando el Excel trae la foto física de toda la empresa (stock total,
producto terminado, pre-PT y ventas mensuales, en MMlb de Mo) en vez del
detalle planta / cliente / SKU. Reconoce el formato del "Modelo de Caja y
Capital de Trabajo":

  08_Inventario_Fisico   Mes, Stock total, Producto terminado, Pre-PT, Own Sales, Safety stock
  09_PT_Cobertura        Mes, Not Assigned, Assigned, In-Transit, Warehouse, Consignment, venta mes siguiente
  01_Supuestos           Valor promedio por libra, tasa de deuda
  03_Ciclo_Transito      Días por etapa (actuales / objetivo)            -> si está completa
  02_Cliente_SKU         Ventas, desviación, LT, servicio, SS real        -> si está completa

Dos niveles de referencia por grupo de stock:

  Nivel demostrado = percentil 25 de los meses observados (lo que la
  operación ya logró en 1 de cada 4 meses). No requiere supuestos.

  Nivel técnico (solo si se informan los días por etapa; d = venta diaria media, MMlb/día):

  Pre-PT          = d · (compra + tránsito MP + puerto + espera pre-proceso + proceso) + SS materia prima
  PT sin asignar  = d · (liberación + ciclo de campaña) + SS producto terminado
  Asignado + In-Transit = d · (preparación + tránsito a cliente)
  Bodegas destino + consignación = d · días de bodega destino + consignación contractual

  SS = z · sqrt(LT · σd² + d² · σLT²), σd diaria = σ mensual de ventas / sqrt(30)
"""
from __future__ import annotations

import math
import statistics
from dataclasses import dataclass, field
from statistics import NormalDist, mean, stdev

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from .loader import to_float, to_fraction, to_month
from .report import D0, PRIO_FILL, _table, _widths
from .schema import norm

X1, X2, X3 = "0.0", "0.00", "0.000"


# ---------------------------------------------------------------- parámetros

@dataclass
class Stage:
    key: str
    name: str            # tal como aparece en 03_Ciclo_Transito
    group: str           # prept | pt | despacho
    days: float          # días necesarios (objetivo)
    help: str
    actual_days: float | None = None
    source: str = "sin dato"

    @property
    def known(self) -> bool:
        return self.source != "sin dato"


def default_stages() -> list[Stage]:
    return [
        Stage("compra", "Compra/espera embarque", "prept", 10, "Concentrado comprado (propio) esperando embarque"),
        Stage("transito_mp", "Tránsito", "prept", 25, "Tránsito del concentrado hasta la planta"),
        Stage("puerto", "Puerto/aduana/recepción", "prept", 5, "Desaduanaje, muestreo y recepción"),
        Stage("espera", "Espera pre-proceso", "prept", 15, "Stock de ciclo de MP: intervalo entre recepciones / 2"),
        Stage("proceso", "Proceso productivo", "prept", 7, "Tostación / conversión (WIP)"),
        Stage("pt", "Producto terminado/espera", "pt", 12, "Liberación QA + ciclo de campaña / 2"),
        Stage("despacho", "Despacho/entrega", "despacho", 35, "Preparación + tránsito hasta el cliente (si la carga es propia)"),
        # Detalle opcional de Despacho/entrega: si se informan los tres, reemplaza a "Despacho/entrega".
        Stage("prep_despacho", "Asignado/preparación despacho", "despacho_det", 15,
              "Días desde que el PT se asigna hasta que sale (promedio sobre la venta total)"),
        Stage("transito_pt", "Tránsito a destino", "despacho_det", 30,
              "Días de tránsito hasta la bodega destino o el cliente (promedio sobre la venta total)"),
        Stage("bodega", "Bodega destino", "despacho_det", 10,
              "Días que el producto regular permanece en bodega destino (sin stock spot ni buffer)"),
    ]


@dataclass
class AggParams:
    precio_usd_lb: float = 33.0
    tasa_costo_capital: float = 0.0
    nivel_servicio: float = 0.95
    lt_reposicion_pt_dias: float = 30.0
    desv_lt_pt_dias: float = 5.0
    desv_lt_mp_dias: float = 7.0
    dias_bodega_destino: float = 10.0
    despacho_incluye_bodega: bool = True  # los días de Despacho/entrega incluyen la estadía en bodega destino
    stock_spot: float = 0.0               # MMlb en bodegas reservadas para ventas spot (decisión comercial)
    buffer_stock: float = 0.0             # MMlb de buffer / safety stock formal en bodegas
    pct_venta_bodega: float = 1.0         # fracción de la venta que pasa por bodega destino
    segmentos_independientes: int = 1     # >1 aproxima el efecto mix cliente/SKU/planta sobre el SS
    stock_inicial: float | None = None    # stock total al cierre del mes previo al primero (balance de masa)
    stages: list[Stage] = field(default_factory=default_stages)
    sources: dict[str, str] = field(default_factory=dict)

    def stage_days(self, group: str) -> float:
        return sum(s.days for s in self.stages if s.group == group)

    def lt_mp(self) -> float:
        return sum(s.days for s in self.stages if s.key in ("compra", "transito_mp", "puerto"))


PARAM_HELP = {
    "stock_inicial": "Stock total al cierre del mes previo al primero (para el balance de masa)",
    "precio_usd_lb": "Valor promedio por libra de Mo (US$/lb)",
    "tasa_costo_capital": "Costo anual del capital inmovilizado (tasa de deuda incremental)",
    "nivel_servicio": "Nivel de servicio objetivo para el stock de seguridad",
    "lt_reposicion_pt_dias": "Lead time de reposición del producto terminado (días)",
    "desv_lt_pt_dias": "Desviación del lead time de reposición PT (días)",
    "desv_lt_mp_dias": "Desviación del lead time de la materia prima (días)",
    "dias_bodega_destino": "Días de venta a mantener en bodegas de destino",
    "segmentos_independientes": "Nº de segmentos independientes (mix); 1 = demanda agregada",
    "despacho_incluye_bodega": "Los días de Despacho/entrega incluyen la estadía en bodega destino",
    "stock_spot": "Stock en bodegas reservado para ventas spot (MMlb)",
    "buffer_stock": "Buffer / safety stock formal en bodegas (MMlb)",
    "pct_venta_bodega": "Fracción de la venta que pasa por bodega destino",
}


# ---------------------------------------------------------------- datos

@dataclass
class MonthRow:
    month: str
    label: str
    total: float | None = None
    pt: float | None = None
    prept: float | None = None
    sales: float | None = None
    production: float | None = None
    tolling: float | None = None          # maquila (material de terceros procesado)
    ss: float | None = None
    not_assigned: float | None = None
    assigned: float | None = None
    in_transit: float | None = None
    warehouse: float | None = None
    consignment: float | None = None
    next_sales: float | None = None


@dataclass
class SkuRow:
    cliente: str
    sku: str
    ventas_lb_mes: float
    desv_lb_mes: float
    lt_dias: float
    desv_lt: float
    servicio: float
    ss_real_lb: float | None
    precio: float | None
    contrato: str = ""


@dataclass
class AggData:
    months: list[MonthRow]
    params: AggParams
    skus: list[SkuRow]
    warnings: list[str]
    title: str = ""


def is_aggregate_workbook(source) -> bool:
    wb = load_workbook(source, read_only=True)
    names = [norm(n) for n in wb.sheetnames]
    wb.close()
    if hasattr(source, "seek"):
        source.seek(0)
    return any("inventario_fisico" in n for n in names)


def _sheet(wb, fragment: str):
    for n in wb.sheetnames:
        if fragment in norm(n):
            return wb[n]
    return None


def _header_rows(ws, must: tuple[str, ...]):
    """Fila de encabezado que contiene todas las palabras de `must` -> (índice columna por nombre, filas)."""
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows):
        hdr = [norm(c) if c is not None else "" for c in row]
        if all(any(m in h for h in hdr) for m in must):
            return hdr, rows[i + 1:]
    return None, []


def _col(hdr: list[str], *fragments: str) -> int | None:
    for frag in fragments:
        for j, h in enumerate(hdr):
            if frag in h:
                return j
    return None


def _cell(row, j):
    return to_float(row[j]) if j is not None and j < len(row) else None


def load_aggregate(source) -> AggData:
    wb = load_workbook(source, data_only=True)
    warnings: list[str] = []
    p = AggParams()
    months: dict[str, MonthRow] = {}

    ws = _sheet(wb, "inventario_fisico")
    hdr, body = _header_rows(ws, ("mes", "stock_total")) if ws else (None, [])
    if hdr:
        c = dict(total=_col(hdr, "stock_total"), pt=_col(hdr, "producto_terminado"), prept=_col(hdr, "pre_pt"),
                 sales=_col(hdr, "own_sales", "ventas"), ss=_col(hdr, "safety"),
                 production=_col(hdr, "produccion_mmlb", "produccion"), tolling=_col(hdr, "maquila", "tolling"))
        for row in body:
            m = to_month(row[0])
            if not m:
                if months:
                    break
                continue
            months[m] = MonthRow(m, str(row[0]), **{k: _cell(row, j) for k, j in c.items()})
    else:
        warnings.append("No se encontró la tabla mensual de inventario físico (Mes / Stock total).")

    ws = _sheet(wb, "pt_cobertura")
    hdr, body = _header_rows(ws, ("mes", "not_assigned")) if ws else (None, [])
    if hdr:
        c = dict(not_assigned=_col(hdr, "not_assigned"), assigned=_col(hdr, "assigned"),
                 in_transit=_col(hdr, "in_transit"), warehouse=_col(hdr, "warehouse"),
                 consignment=_col(hdr, "consignment"), next_sales=_col(hdr, "mes_siguiente"))
        c["assigned"] = next((j for j, h in enumerate(hdr) if h == "assigned"), c["assigned"])
        for row in body:
            m = to_month(row[0])
            if not m:
                if any(r.not_assigned is not None for r in months.values()):
                    break
                continue
            r = months.setdefault(m, MonthRow(m, str(row[0])))
            for k, j in c.items():
                setattr(r, k, _cell(row, j))
    else:
        warnings.append("Sin apertura de producto terminado (Not Assigned / Assigned / In-Transit / Warehouse).")

    ws = _sheet(wb, "supuestos")
    if ws:
        for row in ws.iter_rows(values_only=True):
            label = norm(row[0]) if row and row[0] else ""
            val = to_float(row[1]) if len(row) > 1 else None
            if val is None:
                continue
            if "stock_spot" in label:
                p.stock_spot, p.sources["stock_spot"] = val, f"archivo ({ws.title})"
            elif "buffer_stock" in label:
                p.buffer_stock, p.sources["buffer_stock"] = val, f"archivo ({ws.title})"
            elif "venta" in label and "bodega" in label:
                p.pct_venta_bodega, p.sources["pct_venta_bodega"] = to_fraction(val), f"archivo ({ws.title})"
            elif "stock_total_inicial" in label or "saldo_inicial" in label:   # no confundir con "Stock inicial habitual"
                p.stock_inicial, p.sources["stock_inicial"] = val, f"archivo ({ws.title})"
            elif "valor_promedio_por_libra" in label or "precio" in label:
                p.precio_usd_lb, p.sources["precio_usd_lb"] = val, f"archivo ({ws.title})"
            elif "tasa" in label:
                if val > 0:
                    p.tasa_costo_capital = to_fraction(val)
                    p.sources["tasa_costo_capital"] = f"archivo ({ws.title})"
                else:
                    warnings.append("Tasa de deuda incremental = 0 en el archivo: el ahorro anual en intereses "
                                    "no se calcula hasta completarla (el capital liberable sí).")

    ws = _sheet(wb, "ciclo_transito")
    if ws:
        hdr, body = _header_rows(ws, ("etapa", "dias"))
        if hdr:
            ca, co = _col(hdr, "dias_actuales"), _col(hdr, "dias_objetivo")
            by_name = {norm(s.name): s for s in p.stages}
            filled = 0
            for row in body:
                st = by_name.get(norm(row[0])) if row and row[0] else None
                if not st:
                    continue
                act, obj = _cell(row, ca), _cell(row, co)
                st.actual_days = act
                if obj is not None or act is not None:
                    st.days = obj if obj is not None else act
                    st.source = f"archivo ({ws.title}, {'objetivo' if obj is not None else 'actual'})"
                    filled += 1
            if not filled:
                warnings.append(f"La hoja {ws.title} está vacía: sin días por etapa no se calcula la cobertura "
                                "técnica, solo el nivel demostrado por los propios datos.")

    skus: list[SkuRow] = []
    ws = _sheet(wb, "cliente_sku")
    if ws:
        hdr, body = _header_rows(ws, ("cliente", "sku"))
        if hdr:
            ix = dict(v=_col(hdr, "ventas"), sd=_col(hdr, "desv_demanda"), lt=_col(hdr, "lt_dias"),
                      sdlt=_col(hdr, "desv_lt"), sl=_col(hdr, "servicio"), ss=_col(hdr, "ss_real"),
                      pr=_col(hdr, "us_lb"), ct=_col(hdr, "contrato"))
            for row in body:
                v = _cell(row, ix["v"])
                if not row or not row[0] or not v:
                    continue
                skus.append(SkuRow(
                    str(row[0]), str(row[1] or ""), v, _cell(row, ix["sd"]) or 0.0, _cell(row, ix["lt"]) or 0.0,
                    _cell(row, ix["sdlt"]) or 0.0, to_fraction(row[ix["sl"]]) if ix["sl"] is not None and row[ix["sl"]] else 0.95,
                    _cell(row, ix["ss"]), _cell(row, ix["pr"]),
                    str(row[ix["ct"]] or "") if ix["ct"] is not None and ix["ct"] < len(row) else ""))
            if not skus:
                warnings.append(f"La hoja {ws.title} está vacía: no hay validación de safety stock por cliente/SKU.")
    title = ""
    first = wb.worksheets[0]
    if first["A1"].value:
        title = str(first["A1"].value)
    wb.close()

    rows = sorted(months.values(), key=lambda r: r.month)
    # Si el archivo se guardó sin recalcular, las celdas con fórmula vienen vacías: se reconstruyen.
    for i, r in enumerate(rows):
        if r.pt is None and r.not_assigned is not None:
            r.pt = sum(v or 0 for v in (r.not_assigned, r.assigned, r.in_transit, r.warehouse, r.consignment))
        if r.prept is None and r.total is not None and r.pt is not None:
            r.prept = r.total - r.pt
        if r.next_sales is None and i + 1 < len(rows) and r.not_assigned is not None:
            r.next_sales = rows[i + 1].sales
    if rows and not any(r.production is not None for r in rows):
        warnings.append("Sin producción mensual (columna 'Producción MMlb' de la hoja de inventario físico): "
                        "no se calcula la cobertura técnica del Pre-PT y sus días se expresan sobre la venta.")
    if not rows:
        warnings.append("No hay meses con datos.")
    return AggData(rows, p, skus, warnings, title)


# ---------------------------------------------------------------- cálculo

@dataclass
class GroupResult:
    name: str
    stages: tuple[str, ...]    # etapas que definen el pipeline técnico
    series: list[float]        # MMlb por mes
    pipeline: float | None = None
    buffer: float = 0.0
    detail: str = ""
    flow_day: float = 0.0          # flujo diario que pasa por el grupo (venta o producción)
    flow_label: str = "venta"

    @property
    def actual(self) -> float:
        return mean(self.series) if self.series else 0.0

    @property
    def benchmark(self) -> float:
        """Nivel demostrado: percentil 25 de los meses observados."""
        if len(self.series) >= 4:
            return statistics.quantiles(self.series, n=4, method="inclusive")[0]
        return min(self.series) if self.series else 0.0

    @property
    def technical(self) -> float | None:
        return None if self.pipeline is None else self.pipeline + self.buffer


@dataclass
class AggResult:
    data: AggData
    d_month: float
    sd_month: float
    groups: list[GroupResult]
    ss_pt: float
    ss_mp: float
    stats: dict[str, float | None]
    recs: list[list] = field(default_factory=list)
    sku_rows: list[list] = field(default_factory=list)

    @property
    def d_day(self) -> float:
        return self.d_month / 30

    @property
    def actual_total(self) -> float:
        return sum(g.actual for g in self.groups)

    @property
    def benchmark_total(self) -> float:
        return sum(g.benchmark for g in self.groups)

    @property
    def technical_total(self) -> float | None:
        vals = [g.technical for g in self.groups]
        return None if any(v is None for v in vals) else sum(vals)


def _avg(values) -> float | None:
    vals = [v for v in values if v is not None]
    return mean(vals) if vals else None


def safety(z, lt, sd_day, d_day, sd_lt, segments=1) -> float:
    """SS sumado sobre `segments` partes independientes e iguales de la demanda agregada:
    N · z · sqrt(LT · σd²/N + (d/N)² · σLT²) = z · sqrt(N · LT · σd² + d² · σLT²)."""
    return z * math.sqrt(max(segments, 1) * lt * sd_day ** 2 + d_day ** 2 * sd_lt ** 2)


def analyze_aggregate(data: AggData) -> AggResult:
    p = data.params
    rows = [r for r in data.months if r.sales is not None]
    sales = [r.sales for r in rows]
    d_month = mean(sales) if sales else 0.0
    sd_month = stdev(sales) if len(sales) >= 3 else 0.3 * d_month
    d = d_month / 30
    sd_d = sd_month / math.sqrt(30)
    z = NormalDist().inv_cdf(min(max(p.nivel_servicio, 0.5), 0.9999))
    seg = max(int(p.segmentos_independientes), 1)
    ss_pt = safety(z, p.lt_reposicion_pt_dias, sd_d, d, p.desv_lt_pt_dias, seg)
    ss_mp = safety(z, p.lt_mp(), sd_d, d, p.desv_lt_mp_dias, seg)
    known = {s.key for s in p.stages if s.known}

    def col(*attrs):
        out = []
        for m in data.months:
            vals = [getattr(m, a) for a in attrs]
            if all(v is not None for v in vals):
                out.append(sum(vals))
        return out

    def tech(keys, extra=0.0):
        return d * sum(s.days for s in p.stages if s.key in keys) + extra if set(keys) <= known else None

    def grp(*args, **kw):
        return GroupResult(*args, flow_day=d, **kw)

    # El Pre-PT se vacía con la producción (no con la venta): su flujo es la producción mensual.
    prept_keys = ("compra", "transito_mp", "puerto", "espera", "proceso")
    # Producción propia = producción - maquila (el material de maquila es de terceros, no es stock propio).
    prod = [m.production - (m.tolling or 0) for m in data.months if m.production is not None]
    if prod:
        dp = mean(prod) / 30
        sd_p = (stdev(prod) if len(prod) >= 3 else sd_month) / math.sqrt(30)
        ss_mp = safety(z, p.lt_mp(), sd_p, dp, p.desv_lt_mp_dias, seg)
        pipe = dp * sum(s.days for s in p.stages if s.key in prept_keys) if set(prept_keys) <= known else None
        prept = GroupResult("Pre-PT (compra + MP + WIP)", prept_keys, col("prept"), pipe, ss_mp,
                            "producción diaria x días de compra, tránsito MP, puerto, espera y proceso + SS MP",
                            flow_day=dp, flow_label="producción")
    else:
        prept = GroupResult("Pre-PT (compra + MP + WIP)", prept_keys, col("prept"), None, ss_mp,
                            "Técnico: falta la producción mensual (flujo del Pre-PT)",
                            flow_day=d, flow_label="venta (aprox.)")
    groups = [prept]
    if any(m.not_assigned is not None for m in data.months):
        consig = _avg(m.consignment for m in data.months) or 0.0
        groups += [
            grp("PT sin asignar (Not Assigned)", ("pt",), col("not_assigned"), tech(("pt",)), ss_pt,
                        "d x días de liberación + ciclo de campaña + SS de producto terminado"),
        ]
        policy = p.stock_spot + p.buffer_stock      # stock de política comercial en bodegas, no de flujo
        split = {"prep_despacho", "transito_pt", "bodega"} <= known
        if split:
            bod_days = next(s.days for s in p.stages if s.key == "bodega")
            groups += [
                grp("Asignado (preparación despacho)", ("prep_despacho",), col("assigned"),
                    tech(("prep_despacho",)), 0.0, "d x días desde la asignación hasta la salida"),
                grp("In-Transit", ("transito_pt",), col("in_transit"), tech(("transito_pt",)), 0.0,
                    "d x días de tránsito a bodega destino o cliente"),
                grp("Bodegas destino + consignación", ("bodega",), col("warehouse", "consignment"),
                    d * p.pct_venta_bodega * bod_days + policy + consig, 0.0,
                    f"d x {p.pct_venta_bodega:.0%} de la venta x días en bodega + spot {p.stock_spot:.1f} + "
                    f"buffer {p.buffer_stock:.1f} + consignación {consig:.1f} MMlb"),
            ]
        elif p.despacho_incluye_bodega:
            # Despacho/entrega cubre desde la asignación hasta la entrega, incluida la estadía en bodega destino.
            t = tech(("despacho",))
            groups.append(grp("Despacho: asignado + tránsito + bodegas", ("despacho",),
                              col("assigned", "in_transit", "warehouse", "consignment"),
                              t + policy if t is not None else None, 0.0,
                              "d x días de despacho/entrega (preparación + tránsito + bodega destino)"
                              + (f" + spot y buffer {policy:.1f} MMlb" if policy else "")))
        else:
            groups += [
                grp("Asignado + In-Transit", ("despacho",), col("assigned", "in_transit"), tech(("despacho",)), 0.0,
                    "d x días de preparación + tránsito a cliente"),
                grp("Bodegas destino + consignación", ("bodega",), col("warehouse", "consignment"),
                    d * p.dias_bodega_destino + consig + policy if "dias_bodega_destino" in p.sources else None, 0.0,
                    "d x días de bodega destino + consignación contractual + spot y buffer"),
            ]
    else:
        groups.append(grp("Producto terminado", ("pt", "despacho"), col("pt"),
                                  tech(("pt", "despacho"), d * p.dias_bodega_destino), ss_pt,
                                  "d x días de PT + despacho + bodega destino + SS"))

    stats: dict[str, float | None] = {"cv": sd_month / d_month if d_month else None}
    pairs = [(r.total, r.sales) for r in rows if r.total is not None]
    if len(pairs) >= 3:
        try:
            stats["corr_stock_ventas"] = statistics.correlation([a for a, _ in pairs], [b for _, b in pairs])
        except statistics.StatisticsError:      # stock o venta constantes: no hay correlación que medir
            stats["corr_stock_ventas"] = None
        stats["cv_stock"] = stdev([a for a, _ in pairs]) / mean([a for a, _ in pairs])
        stats["stock_medio"] = mean([a for a, _ in pairs])
    if len(rows) >= 6:
        for k in ("prept", "pt", "total", "sales"):
            stats[f"{k}_ult3"] = _avg(getattr(r, k) for r in rows[-3:])
            stats[f"{k}_prev"] = _avg(getattr(r, k) for r in rows[:-3])
    cover = [min(1.0, ((r.assigned or 0) + (r.warehouse or 0)) / r.next_sales) for r in data.months
             if r.next_sales and r.assigned is not None]
    stats["cobertura_bodega_venta_sig"] = mean(cover) if cover else None
    stats["meses_cobertura_completa"] = sum(1 for c in cover if c >= 0.98) if cover else None
    stats["meses_con_split"] = len(cover)

    res = AggResult(data, d_month, sd_month, groups, ss_pt, ss_mp, stats)
    res.sku_rows = _sku_table(data)
    res.recs = _recommendations(res)
    return res


def _sku_table(data: AggData) -> list[list]:
    out = []
    for s in data.skus:
        z = NormalDist().inv_cdf(min(max(s.servicio, 0.5), 0.9999))
        d, sd = s.ventas_lb_mes / 30, s.desv_lb_mes / math.sqrt(30)
        ss_t = z * math.sqrt(s.lt_dias * sd ** 2 + d ** 2 * s.desv_lt ** 2)
        exceso = (s.ss_real_lb - ss_t) if s.ss_real_lb is not None else None
        price = s.precio or data.params.precio_usd_lb
        out.append([s.cliente, s.sku, s.ventas_lb_mes, s.desv_lb_mes, s.lt_dias, s.desv_lt, s.servicio, z,
                    s.ss_real_lb, ss_t, exceso, s.ss_real_lb / d if s.ss_real_lb is not None and d else None,
                    ss_t / d if d else None, exceso * price / 1e6 if exceso is not None else None, s.contrato])
    return out


def _recommendations(r: AggResult) -> list[list]:
    p, st, d = r.data.params, r.stats, r.d_day
    usd = p.precio_usd_lb
    recs: list[list] = []

    recs.append(["Alta", "Valor de un día de cobertura",
                 f"Con ventas medias de {r.d_month:.2f} MMlb/mes, 1 día de inventario = {d:.3f} MMlb = "
                 f"US${d * usd:,.1f} mm a {usd:.0f} US$/lb.",
                 "Usar este factor para convertir cada día que se reduzca en una etapa (una vez validado) en "
                 "caja y en intereses evitados.", None, None])

    for g in r.groups:
        fd = g.flow_day or d
        gap = g.actual - g.benchmark
        if gap >= fd:
            recs.append(["Alta" if gap * usd >= 30 else "Media", f"{g.name}: volver al nivel ya demostrado",
                         f"Promedio {g.actual:.2f} MMlb ({g.actual / fd:.0f} días de {g.flow_label}); en el 25% de "
                         f"los mejores meses estuvo en {g.benchmark:.2f} MMlb ({g.benchmark / fd:.0f} días) o menos.",
                         _action_for(g.name), gap, gap * usd])
        if g.technical is not None:
            tgap = g.actual - g.technical
            if g.buffer and g.pipeline is not None and g.pipeline - fd <= g.actual < g.technical:
                recs.append(["Media", f"{g.name}: cubre el pipeline, sin colchón de seguridad",
                             f"Promedio {g.actual:.2f} MMlb ({g.actual / fd:.0f} días de {g.flow_label}) = días por etapa "
                             f"({g.pipeline / fd:.0f}); el modelo sugiere {g.buffer / fd:.0f} días adicionales de "
                             f"seguridad ({g.buffer:.2f} MMlb), que dependen de la desviación supuesta del lead time.",
                             "No hay exceso estructural en este grupo. Validar la desviación real del lead time (barra "
                             "lateral): si las esperas ya actúan como colchón, el nivel actual es el adecuado.",
                             None, None])
            elif abs(tgap) >= fd:
                recs.append(["Alta" if tgap > 0 else "Media",
                             f"{g.name}: {'sobre' if tgap > 0 else 'bajo'} la cobertura técnica",
                             f"Promedio {g.actual:.2f} MMlb ({g.actual / fd:.0f} días de {g.flow_label}) vs técnico "
                             f"{g.technical:.2f} MMlb ({g.technical / fd:.0f} días = pipeline {g.pipeline / fd:.0f} + "
                             f"buffer {g.buffer / fd:.0f}).",
                             _action_for(g.name) if tgap > 0 else
                             "Con estos días por etapa el stock actual no alcanza: revisar los días informados o "
                             "el riesgo de servicio de la etapa.", tgap, tgap * usd])

    if st.get("prept_ult3") is not None and st.get("prept_prev") is not None:
        dp = st["prept_ult3"] - st["prept_prev"]
        ds = (st.get("sales_ult3") or 0) - (st.get("sales_prev") or 0)
        if dp > 0.5 and ds <= 0.1:
            recs.append(["Alta", "Pre-PT creciendo sin aumento de ventas",
                         f"Pre-PT de los últimos 3 meses {st['prept_ult3']:.1f} MMlb vs {st['prept_prev']:.1f} MMlb "
                         f"en los meses anteriores (+{dp:.1f} MMlb = US${dp * usd:,.0f} mm), mientras la venta "
                         f"varió {ds:+.1f} MMlb/mes.",
                         "Abrir Pre-PT en compra en ruta / MP en planta / WIP y contrastar el programa de compras de "
                         "concentrado con el plan de ventas: es inventario que se acumula antes de producir.",
                         dp, dp * usd])

    corr = st.get("corr_stock_ventas")
    if corr is not None and corr < 0.3:
        recs.append(["Alta", "El stock no sigue a la venta",
                     f"Correlación stock total vs venta del mes = {corr:+.2f}; el stock varía {st['cv_stock']:.0%} "
                     f"(CV) y la venta {st['cv']:.0%}. El stock se comporta como un nivel fijo "
                     f"(~{st['stock_medio']:.0f} MMlb) y no como días de venta futura.",
                     "Fijar la política en días de cobertura por etapa x pronóstico de venta: cuando la venta baja, "
                     "el objetivo de stock baja automáticamente.", None, None])

    na = next((g for g in r.groups if g.name.startswith("PT sin asignar")), None)
    if na:
        ss_formal = next((m.ss for m in reversed(r.data.months) if m.ss is not None), None)
        recs.append(["Media", "PT sin asignar vs stock de seguridad estadístico",
                     f"Not Assigned promedio {na.actual:.2f} MMlb ({na.actual / d:.0f} días). Con la variabilidad "
                     f"observada de la venta, el SS de PT al {p.nivel_servicio:.0%} con reposición "
                     f"{p.lt_reposicion_pt_dias:.0f}±{p.desv_lt_pt_dias:.0f} días es {r.ss_pt:.2f} MMlb "
                     f"({r.ss_pt / d:.0f} días)" + (f"; el safety stock formal informado es {ss_formal:.1f} MMlb."
                                                    if ss_formal is not None else "."),
                     "Separar el Not Assigned en buffer requerido, producción anticipada a contratos y stock sin "
                     "destino. La hoja Sensibilidad_SS muestra cómo cambia el SS con servicio, lead time y mix.",
                     None, None])

    if p.stock_spot:
        cost = p.stock_spot * usd * p.tasa_costo_capital
        recs.append(["Media", "Stock para ventas spot (decisión comercial)",
                     f"{p.stock_spot:.1f} MMlb reservados para spot = US${p.stock_spot * usd:,.0f} mm inmovilizados"
                     + (f", con un costo financiero de US${cost:,.1f} mm/año." if cost else "."),
                     "Mantenerlo solo si el margen adicional de las ventas spot supera ese costo; revisar el tamaño "
                     "del stock spot según el volumen spot efectivamente vendido.", None, None])
    it = _avg(x.in_transit for x in r.data.months)
    asg, wh = _avg(x.assigned for x in r.data.months), _avg(x.warehouse for x in r.data.months)
    if p.despacho_incluye_bodega and None not in (it, asg, wh):
        recs.append(["Media", "Dónde están los días de despacho",
                     f"Del bloque de despacho, Assigned = {asg / d:.0f} días, In-Transit = {it / d:.0f} días y "
                     f"Warehouse = {wh / d:.0f} días de venta (promedio).",
                     "La mayor parte del tiempo está entre la asignación y la llegada (preparación + tránsito), no "
                     "en las bodegas: es la palanca principal para reducir días de despacho.", None, None])
    if it:
        recs.append(["Media", "Tránsito implícito",
                     f"In-Transit promedio {it:.2f} MMlb = {it / d:.0f} días de venta: si todo es propio, el "
                     f"tránsito medio a cliente sería ~{it / d:.0f} días.",
                     "Comparar con el tránsito real por ruta y validar Incoterms (lo vendido FOB/FCA deja de ser "
                     "inventario al embarcar). Si el tránsito real es menor, se despacha antes de lo que el cliente "
                     "necesita.", None, None])

    cov = st.get("cobertura_bodega_venta_sig")
    if cov is not None:
        recs.append(["Media", "Cobertura de la venta del mes siguiente",
                     f"Assigned + Warehouse cubrió en promedio {cov:.0%} de la venta del mes siguiente (tope 100%) "
                     f"({st['meses_cobertura_completa']} de {st['meses_con_split']} meses al 100%), sin usar "
                     "Not Assigned ni In-Transit.",
                     "El In-Transit remanente y el Not Assigned están financiando meses posteriores: definir "
                     "cuántos días hacia adelante se quiere cubrir y convertirlo en política.", None, None])

    bal = balance(r.data)
    if bal:
        if abs(bal["gap"]) > 0.05 * bal["venta"]:
            recs.append(["Alta", "Balance de masa: producción vs venta vs stock",
                         f"{bal['desde']} a {bal['hasta']}: producción {bal['produccion']:.1f} MMlb = venta propia "
                         f"{bal['venta']:.1f} + maquila {bal['maquila']:.1f} + variación de stock "
                         f"{bal['dstock']:+.1f} ({bal['base']}). Quedan {bal['gap']:+.1f} MMlb sin explicar "
                         f"({bal['gap'] / bal['meses']:+.1f} MMlb/mes).",
                         "Aclarar qué incluye la producción y no la venta propia ni el stock: maquila de terceros, "
                         "producción intermedia contada dos veces (óxido que luego se convierte en FeMo), ventas no "
                         "incluidas en Own Sales o mermas. Para el Pre-PT usar solo la producción propia.",
                         gap, None])
        elif bal["base"].startswith("aprox"):
            recs.append(["Baja", "Balance de masa: cuadra (falta saldo inicial)",
                         f"Producción {bal['produccion']:.1f} = venta {bal['venta']:.1f} + maquila {bal['maquila']:.1f} "
                         f"+ stock {bal['dstock']:+.1f}; diferencia {bal['gap']:+.1f} MMlb. La variación de stock se "
                         f"aproximó con el cierre de {bal['desde']} como saldo inicial.",
                         "Agregar en 01_Supuestos el 'Stock total inicial' (cierre del mes anterior al primero) para "
                         "cerrar el balance exacto.", None, None])
    pre = r.groups[0]
    if pre.flow_label != "producción":
        recs.append(["Media", "Agregar producción mensual",
                     "El Pre-PT se vacía con la producción, no con la venta. Sin la producción mensual no se puede "
                     "saber si su stock corresponde a los días de compra, tránsito, puerto, espera y proceso.",
                     "Completar la columna 'Producción MMlb' de la hoja 08_Inventario_Fisico (una cifra por mes).",
                     None, None])
    missing = [s.name for s in p.stages if not s.known]
    if missing:
        recs.append(["Media", "Completar días por etapa (03_Ciclo_Transito)",
                     f"Sin días reales para: {', '.join(missing)}. Sin ellos no se calcula la cobertura técnica "
                     "(pipeline + seguridad), solo el nivel demostrado.",
                     "Completar 'Días actuales' y 'Días objetivo' por etapa; el modelo los toma directamente.",
                     None, None])
    if not r.data.skus:
        recs.append(["Media", "Completar 02_Cliente_SKU",
                     "Sin ventas, desviación y lead time por cliente/SKU el stock de seguridad se calcula agregado, "
                     "lo que subestima el efecto mix (ver columnas de segmentos en Sensibilidad_SS).",
                     "Completar 02_Cliente_SKU para validar SS real vs teórico por cliente; con planta de origen y "
                     "tránsitos se puede usar el modo detallado (reasignación entre plantas).", None, None])
    else:
        exc = [row for row in r.sku_rows if row[10] is not None and row[10] > 0]
        if exc:
            lb = sum(row[10] for row in exc)
            recs.append(["Alta", "Safety stock sobre el teórico por cliente/SKU",
                         f"{len(exc)} líneas cliente/SKU con SS real sobre el teórico: {lb / 1e6:.2f} MMlb.",
                         "Revisar contrato y nivel de servicio de cada línea (hoja Cliente_SKU).", lb / 1e6,
                         sum(row[13] or 0 for row in exc)])
    order = {"Alta": 0, "Media": 1, "Baja": 2}
    return sorted(recs, key=lambda x: (order[x[0]], -(x[5] or 0)))


def balance(data: AggData) -> dict | None:
    """Producción = venta propia + maquila + variación de stock (+ diferencia sin explicar)."""
    bal = [m for m in data.months if m.production is not None and m.sales is not None]
    if len(bal) < 3 or bal[-1].total is None:
        return None
    opening = data.params.stock_inicial
    if opening is not None:
        base = "saldo inicial informado"
    elif bal[0].total is not None:
        opening, base = bal[0].total, f"aprox.: cierre de {bal[0].label} como saldo inicial"
    else:
        return None
    prod = sum(m.production for m in bal)
    sold = sum(m.sales for m in bal)
    toll = sum(m.tolling or 0 for m in bal)
    dstock = bal[-1].total - opening
    return dict(desde=bal[0].label, hasta=bal[-1].label, meses=len(bal), produccion=prod, venta=sold,
                maquila=toll, stock_inicial=opening, stock_final=bal[-1].total, dstock=dstock, base=base,
                gap=prod - sold - toll - dstock, saldo_inicial_implicito=bal[-1].total - (prod - sold - toll))


def balance_table(r: "AggResult") -> dict:
    b = balance(r.data)
    if not b:
        return dict(headers=["Concepto", "MMlb"], rows=[["Sin datos de producción", None]], fmts=[None, X1])
    rows = [["Producción", b["produccion"]], ["(−) Venta propia (Own Sales)", b["venta"]],
            ["(−) Maquila", b["maquila"]],
            [f"Stock inicial ({b['base']})", b["stock_inicial"]], [f"Stock final ({b['hasta']})", b["stock_final"]],
            ["(−) Variación de stock", b["dstock"]], ["= Diferencia sin explicar", b["gap"]],
            ["Saldo inicial que cerraría el balance", b["saldo_inicial_implicito"]]]
    return dict(headers=["Concepto", "MMlb"], rows=rows, fmts=[None, X1])


def _action_for(name: str) -> str:
    if name.startswith("Pre-PT"):
        return ("Sincronizar compras de concentrado con el programa de producción; reducir esperas pre-proceso y "
                "tiempos de puerto/recepción; abrir Pre-PT por etapa para ubicar el exceso.")
    if name.startswith("PT sin asignar"):
        return ("Producir contra pedidos/pronóstico (campañas más cortas), acelerar la liberación QA y asignar el "
                "stock sin destino a contratos o venta spot.")
    if name.startswith("Despacho"):
        return ("Despachar según la fecha requerida por el cliente (no antes), revisar tiempos de preparación de "
                "embarque y ajustar el stock de bodegas destino a días de venta objetivo por mercado.")
    if name.startswith("Asignado (prep"):
        return "Acortar el tiempo entre la asignación y el embarque: programar la producción y el despacho contra la fecha de salida."
    if name.startswith("In-Transit"):
        return "Embarcar según la fecha requerida en destino (no antes) y revisar rutas y frecuencias navieras."
    if name.startswith("Asignado"):
        return ("Despachar según la fecha requerida por el cliente (no antes), revisar Incoterms y tiempos de "
                "preparación de embarque.")
    if name.startswith("Bodegas"):
        return "Ajustar los niveles de bodegas de destino a días de venta objetivo por mercado."
    return "Revisar la política de cobertura de esta etapa."


# ---------------------------------------------------------------- tablas y reporte

def month_table(r: AggResult) -> dict:
    d = r.d_day
    rows = []
    for m in r.data.months:
        s = m.sales
        pr = m.production - (m.tolling or 0) if m.production is not None else None
        rows.append([m.label, m.total, m.pt, m.prept, s, m.production, m.tolling, pr,
                     s / m.total if s and m.total else None, s / m.pt if s and m.pt else None,
                     pr / m.prept if pr and m.prept else None,
                     m.total / s * 30 if s and m.total else None,
                     m.prept / pr * 30 if pr and m.prept is not None else None,
                     m.not_assigned / d if m.not_assigned is not None and d else None,
                     m.assigned / d if m.assigned is not None and d else None,
                     m.in_transit / d if m.in_transit is not None and d else None,
                     m.warehouse / d if m.warehouse is not None and d else None,
                     ((m.assigned or 0) + (m.warehouse or 0)) / m.next_sales
                     if m.next_sales and m.assigned is not None else None])
    return dict(headers=["Mes", "Stock total MMlb", "PT MMlb", "Pre-PT MMlb", "Ventas MMlb", "Producción MMlb",
                         "Maquila MMlb", "Producción propia MMlb",
                         "Ciclos/mes total (venta)", "Ciclos/mes PT (venta)", "Ciclos/mes Pre-PT (prod. propia)",
                         "Días stock total", "Días Pre-PT (prod. propia)",
                         "Días Not Assigned", "Días Assigned", "Días In-Transit", "Días Warehouse",
                         "% venta sig. cubierta por bodega"],
                rows=rows, fmts=[None, X1, X1, X1, X1, X1, X1, X1, X3, X3, X3, D0, D0, D0, D0, D0, D0, "0%"])


def group_table(r: AggResult) -> dict:
    d, usd = r.d_day, r.data.params.precio_usd_lb

    def row(name, fd, flow, actual, bench, tech, detail):
        return [name, flow, actual, actual / fd, bench, bench / fd, actual - bench, (actual - bench) * usd,
                tech, tech / fd if tech is not None else None,
                (actual - tech) * usd if tech is not None else None, detail]

    def note(g):
        if g.technical is not None:
            return g.detail
        return g.detail if g.detail.startswith("Técnico:") else "Técnico: faltan días reales por etapa"

    rows = [row(g.name, g.flow_day or d, g.flow_label, g.actual, g.benchmark, g.technical, note(g)) for g in r.groups]
    rows.append(row("TOTAL", d, "venta", r.actual_total, r.benchmark_total, r.technical_total,
                    "El nivel demostrado suma los mejores meses de cada grupo (no ocurrieron todos a la vez)"))
    return dict(headers=["Grupo", "Días medidos sobre", "Actual MMlb", "Actual días", "Demostrado MMlb (P25)",
                         "Demostrado días", "Exceso vs demostrado MMlb", "Exceso vs demostrado US$mm",
                         "Técnico MMlb", "Técnico días", "Exceso vs técnico US$mm", "Nota"],
                rows=rows, fmts=[None, None, X2, D0, X2, D0, X2, "#,##0.0", X2, D0, "#,##0.0", None])


def stage_table(r: AggResult) -> dict:
    d = r.d_day
    return dict(headers=["Etapa", "Días necesarios", "Días actuales informados", "MMlb equivalentes", "Fuente",
                         "Descripción"],
                rows=[[s.name, s.days if s.known else None, s.actual_days, s.days * d if s.known else None,
                       s.source, s.help] for s in r.data.params.stages],
                fmts=[None, D0, D0, X2, None, None])


def sensitivity_table(r: AggResult) -> dict:
    p = r.data.params
    d, sd_d = r.d_day, r.sd_month / math.sqrt(30)
    rows = []
    for sl in (0.90, 0.95, 0.98, 0.99):
        z = NormalDist().inv_cdf(sl)
        row = [f"{sl:.0%}"]
        for lt, sdlt in ((15, 3), (30, 5), (45, 7), (60, 10)):
            row.append(safety(z, lt, sd_d, d, sdlt) / d)
        for seg in (5, 20):
            row.append(safety(z, p.lt_reposicion_pt_dias, sd_d, d, p.desv_lt_pt_dias, seg) / d)
        rows.append(row)
    return dict(headers=["Nivel de servicio", "SS días LT 15±3", "SS días LT 30±5", "SS días LT 45±7",
                         "SS días LT 60±10", f"SS días LT {p.lt_reposicion_pt_dias:.0f}±{p.desv_lt_pt_dias:.0f}, 5 segmentos",
                         f"SS días LT {p.lt_reposicion_pt_dias:.0f}±{p.desv_lt_pt_dias:.0f}, 20 segmentos"],
                rows=rows, fmts=[None, X1, X1, X1, X1, X1, X1])


def sku_table(r: AggResult) -> dict:
    return dict(headers=["Cliente", "SKU", "Ventas lb/mes", "Desv. lb/mes", "LT días", "Desv. LT", "Servicio", "Z",
                         "SS real lb", "SS teórico lb", "Exceso lb", "Días SS real", "Días SS teórico",
                         "Caja liberable US$mm", "Contrato"],
                rows=r.sku_rows, fmts=[None, None, "#,##0", "#,##0", D0, D0, "0%", X2, "#,##0", "#,##0", "#,##0",
                                       D0, D0, X2, None])


def kpis(r: AggResult) -> list[list]:
    p, d = r.data.params, r.d_day
    rate = p.tasa_costo_capital
    bench_gap = r.actual_total - r.benchmark_total
    tech = r.technical_total

    def turns(stock, k=1):
        return k * r.d_month / stock if stock else None

    return [
        ["Ventas promedio (MMlb/mes)", r.d_month, None, None],
        ["Variabilidad de la venta (CV mensual)", r.stats["cv"], None, None],
        ["Stock promedio (MMlb)", r.actual_total, r.benchmark_total, tech],
        ["Ciclos de rotación por mes", turns(r.actual_total), turns(r.benchmark_total), turns(tech)],
        ["Ciclos de rotación por año", turns(r.actual_total, 12), turns(r.benchmark_total, 12), turns(tech, 12)],
        ["Días de cobertura", r.actual_total / d if d else None, r.benchmark_total / d if d else None,
         tech / d if tech is not None and d else None],
        ["Valor de 1 día de cobertura (US$mm)", d * p.precio_usd_lb, None, None],
        ["Capital liberable vs actual (US$mm)", None, bench_gap * p.precio_usd_lb,
         (r.actual_total - tech) * p.precio_usd_lb if tech is not None else None],
        ["Intereses evitados por año (US$mm)", None, bench_gap * p.precio_usd_lb * rate if rate else None,
         (r.actual_total - tech) * p.precio_usd_lb * rate if rate and tech is not None else None],
    ]


def write_aggregate_report(r: AggResult, path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "Rotación y cobertura óptima de inventario - vista agregada"
    ws["A1"].font = Font(bold=True, size=14)
    months = r.data.months
    ws["A2"] = (f"{r.data.title} · {months[0].label} a {months[-1].label} ({len(months)} meses) · "
                f"unidades: MMlb de Mo, días de venta, US$mm a {r.data.params.precio_usd_lb:.0f} US$/lb")
    nxt = _table(ws, 4, ["Indicador", "Actual (promedio)", "Nivel demostrado (P25)", "Técnico (pipeline + SS)"],
                 kpis(r), [None, "#,##0.000", "#,##0.000", "#,##0.000"])
    g = group_table(r)
    ws.cell(nxt, 1, "Cobertura por grupo de stock").font = Font(bold=True, size=12)
    nxt = _table(ws, nxt + 1, g["headers"], g["rows"], g["fmts"])
    for i, note in enumerate((
            "Nivel demostrado = percentil 25 de los meses observados: un nivel que la operación ya logró en 1 de "
            "cada 4 meses. Es una meta conservadora y verificable.",
            "Técnico = días necesarios por etapa x venta diaria + stock de seguridad estadístico. Requiere los días "
            "reales por etapa (03_Ciclo_Transito); sin ellos queda vacío.",
            "Ver hojas Etapas, Sensibilidad_SS y Supuestos para el origen de cada número.")):
        ws.cell(nxt - 1 + i, 1, note).font = Font(italic=True, size=9)
    _widths(ws, {1: 40, 2: 16, 3: 16, 4: 16, 11: 60})

    ws = wb.create_sheet("Recomendaciones")
    _table(ws, 1, ["Prioridad", "Tema", "Diagnóstico", "Acción propuesta", "Impacto MMlb", "Impacto US$mm"],
           r.recs, [None, None, None, None, X2, "#,##0.0"], status_col=0, fills=PRIO_FILL)
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    _widths(ws, {1: 10, 2: 34, 3: 70, 4: 70, 5: 13, 6: 13})

    for name, tb in (("Rotacion_Mensual", month_table(r)), ("Balance_Masa", balance_table(r)), ("Etapas", stage_table(r)),
                     ("Sensibilidad_SS", sensitivity_table(r))):
        ws = wb.create_sheet(name)
        nxt = _table(ws, 1, tb["headers"], tb["rows"], tb["fmts"])
        _widths(ws, {1: 26, 6: 34, 7: 50} if name == "Etapas" else {1: 18})
        if name == "Sensibilidad_SS":
            ws.cell(nxt, 1, "Stock de seguridad expresado en días de venta. Segmentos = cuántos grupos "
                            "cliente/SKU/planta independientes hay: con mix la demanda se diversifica menos y el SS "
                            "necesario sube.").font = Font(italic=True, size=9)
    if r.sku_rows:
        tb = sku_table(r)
        ws = wb.create_sheet("Cliente_SKU")
        _table(ws, 1, tb["headers"], tb["rows"], tb["fmts"])
        _widths(ws, {1: 24, 2: 18})

    ws = wb.create_sheet("Supuestos")
    p = r.data.params
    rows = [[k, getattr(p, k), PARAM_HELP[k], p.sources.get(k, "supuesto por defecto")] for k in PARAM_HELP]
    nxt = _table(ws, 1, ["Parámetro", "Valor", "Descripción", "Fuente"], rows)
    ws.cell(nxt, 1, "Advertencias").font = Font(bold=True, size=12)
    _table(ws, nxt + 1, ["#", "Advertencia"], [[i, w] for i, w in enumerate(r.data.warnings, 1)] or [["-", "Sin advertencias"]])
    _widths(ws, {1: 28, 2: 90, 3: 60, 4: 30})

    ws = wb.create_sheet("Metodologia")
    for i, line in enumerate((__doc__ or "").strip().splitlines(), start=1):
        ws.cell(i, 1, line)
    ws.column_dimensions["A"].width = 120
    wb.save(path)
