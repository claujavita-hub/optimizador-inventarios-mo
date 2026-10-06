"""Motor de cálculo: rotación actual, stock óptimo y ciclos por mes.

Modelo (por línea cliente-producto, abastecida desde una planta):

  Exposición E = intervalo de campaña R + (proceso + QA) + tránsito T - aviso firme F
  Stock de seguridad SS = z * sqrt(E * σd² + d² * (σproceso² + σtránsito²))
  Stock de ciclo      = d * max(R, intervalo de despacho) / 2
  Stock en proceso    = d * tiempo de proceso
  Stock en tránsito   = d * T  (si la empresa es dueña de la carga en tránsito)

  d, σd = demanda diaria media y su desviación (de la historia mensual,
  σd_diaria = σd_mensual / sqrt(días_mes)); z = cuantil normal del nivel de
  servicio del cliente.

Por nodo planta-producto el SS se agrega con pooling de riesgo
(sqrt de la suma de cuadrados) y se definen niveles mínimo (SS), objetivo
(SS + ciclo) y máximo (SS + 2·ciclo) del stock de producto terminado.
Rotación = demanda mensual / inventario total (PT + proceso + tránsito):
ciclos por mes. En concentrado se aplica el mismo modelo con el lead time
de abastecimiento de cada planta.
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field
from statistics import NormalDist, mean, stdev

from .model import Dataset
from .schema import norm

DEFAULT_MO_CONTENT = 0.57
CONVERSION_YIELD = 0.97


def z_value(service_level: float) -> float:
    service_level = min(max(service_level, 0.5), 0.9999)
    return NormalDist().inv_cdf(service_level)


def safety_stock(z: float, exposure_days: float, sd_day: float, d_day: float, lt_var: float) -> float:
    if exposure_days <= 0:
        return 0.0
    return z * math.sqrt(exposure_days * sd_day ** 2 + d_day ** 2 * lt_var)


# ---------------------------------------------------------------- resultados

@dataclass
class LineResult:
    customer: str
    product: str
    plant: str
    country: str
    region: str
    abc: str
    xyz: str
    service_level: float
    z: float
    months: int
    d_month: float
    sd_month: float
    cv: float
    transit_days: float
    transit_sd: float
    freight_usd_t: float | None
    lane_source: str
    process_days: float
    process_sd: float
    qa_days: float
    campaign_days: float
    dispatch_days: float
    firm_days: float
    exposure_days: float
    ss_t: float
    cycle_t: float
    wip_t: float
    transit_stock_t: float
    share_demand: float
    share_process: float
    share_transit: float
    value_usd_t: float
    conversion_usd_t: float
    raw_material_usd_t: float = 0.0     # logística del concentrado por t de producto

    @property
    def annual_t(self) -> float:
        return self.d_month * 12


@dataclass
class NodeResult:
    plant: str
    product: str
    value_usd_t: float
    d_month: float = 0.0
    sd_month: float = 0.0
    n_customers: int = 0
    ss_t: float = 0.0
    ss_unpooled_t: float = 0.0
    cycle_t: float = 0.0
    wip_t: float = 0.0
    transit_t: float = 0.0
    lead_prod_days: float = 0.0
    campaign_days: float = 0.0
    min_t: float = 0.0
    target_t: float = 0.0
    max_t: float = 0.0
    rop_t: float = 0.0
    has_current: bool = False
    cur_fg_t: float = 0.0
    cur_consig_t: float = 0.0
    cur_wip_t: float = 0.0
    cur_transit_t: float = 0.0
    status: str = "Sin dato"

    @property
    def optimal_total_t(self) -> float:
        return self.target_t + self.wip_t + self.transit_t

    @property
    def cur_fg_total_t(self) -> float:
        return self.cur_fg_t + self.cur_consig_t

    @property
    def current_total_t(self) -> float:
        return self.cur_fg_total_t + self.cur_wip_t + self.cur_transit_t

    @property
    def cycles_month_current(self) -> float | None:
        return self.d_month / self.current_total_t if self.has_current and self.current_total_t > 0 else None

    @property
    def cycles_month_optimal(self) -> float | None:
        return self.d_month / self.optimal_total_t if self.optimal_total_t > 0 else None

    @property
    def doi_current(self) -> float | None:
        d = self.d_month / 30.0
        return self.current_total_t / d if self.has_current and d > 0 else None

    @property
    def doi_optimal(self) -> float | None:
        d = self.d_month / 30.0
        return self.optimal_total_t / d if d > 0 else None

    @property
    def delta_t(self) -> float | None:
        return self.current_total_t - self.optimal_total_t if self.has_current else None


@dataclass
class PlantResult:
    plant: str
    country: str
    region: str
    capacity_t_month: float | None
    load_t_month: float = 0.0
    conc_d_day: float = 0.0
    conc_ss_t: float = 0.0
    conc_cycle_t: float = 0.0
    conc_lead_time: float = 0.0
    conc_current_t: float | None = None
    conc_value_usd_t: float = 0.0

    @property
    def utilization(self) -> float | None:
        return self.load_t_month / self.capacity_t_month if self.capacity_t_month else None

    @property
    def conc_target_t(self) -> float:
        return self.conc_ss_t + self.conc_cycle_t

    @property
    def conc_coverage_days(self) -> float | None:
        return self.conc_current_t / self.conc_d_day if self.conc_current_t is not None and self.conc_d_day else None


@dataclass
class Analysis:
    ds: Dataset
    assignment: dict[tuple[str, str], str]
    lines: list[LineResult]
    nodes: list[NodeResult]
    plants: list[PlantResult]
    months: list[str] = field(default_factory=list)

    # --------------------------- agregados
    def totals(self) -> dict[str, float | None]:
        has_cur = any(n.has_current for n in self.nodes)
        cur_val = sum(n.current_total_t * n.value_usd_t for n in self.nodes)
        opt_val = sum(n.optimal_total_t * n.value_usd_t for n in self.nodes)
        conc_cur_known = [p for p in self.plants if p.conc_current_t is not None]
        conc_cur_val = sum(p.conc_current_t * p.conc_value_usd_t for p in conc_cur_known)
        conc_opt_val = sum(p.conc_target_t * p.conc_value_usd_t for p in self.plants)
        conc_opt_val_known = sum(p.conc_target_t * p.conc_value_usd_t for p in conc_cur_known)
        sales_val_month = sum(n.d_month * n.value_usd_t for n in self.nodes)
        rate = self.ds.params.tasa_costo_capital
        cur_total = cur_val + conc_cur_val if has_cur else None
        opt_total_cmp = opt_val + conc_opt_val_known
        return {
            "demanda_t_mes": sum(n.d_month for n in self.nodes),
            "ventas_usd_mes": sales_val_month,
            "inv_actual_t": sum(n.current_total_t for n in self.nodes) if has_cur else None,
            "inv_optimo_t": sum(n.optimal_total_t for n in self.nodes),
            "inv_actual_usd": cur_total,
            "inv_optimo_usd": opt_val + conc_opt_val,
            "pt_actual_usd": cur_val if has_cur else None,
            "pt_optimo_usd": opt_val,
            "conc_actual_usd": conc_cur_val if conc_cur_known else None,
            "conc_optimo_usd": conc_opt_val,
            "ciclos_mes_actual": sales_val_month / cur_val if has_cur and cur_val else None,
            "ciclos_mes_optimo": sales_val_month / opt_val if opt_val else None,
            "doi_actual": cur_val / (sales_val_month / 30) if has_cur and sales_val_month else None,
            "doi_optimo": opt_val / (sales_val_month / 30) if sales_val_month else None,
            "capital_liberable_usd": (cur_total - opt_total_cmp) if cur_total is not None else None,
            "ahorro_anual_usd": (cur_total - opt_total_cmp) * rate if cur_total is not None else None,
            "ss_pooling_t": sum(n.ss_unpooled_t - n.ss_t for n in self.nodes),
        }

    def line_annual_cost(self, line: LineResult) -> dict[str, float]:
        rate = self.ds.params.tasa_costo_capital
        freight = (line.freight_usd_t or 0.0) * line.annual_t
        conversion = (line.conversion_usd_t + line.raw_material_usd_t) * line.annual_t
        holding = (line.ss_t + line.cycle_t + line.transit_stock_t) * line.value_usd_t * rate
        return {"flete": freight, "conversion": conversion, "inventario": holding,
                "total": freight + conversion + holding}


# ---------------------------------------------------------------- helpers de datos

class Lookup:
    """Consultas a los datos con defaults y registro de advertencias."""

    def __init__(self, ds: Dataset):
        self.ds = ds
        self.p = ds.params
        self._lanes: dict[tuple[str, str], tuple] = {}
        self.has_freight = any(l.freight_usd_t is not None for l in ds.lanes)
        self.has_conv_cost = any(pr.cost_usd_t is not None for pr in ds.processes.values())
        self.producers: dict[str, set[str]] = defaultdict(set)
        for (pl, pr) in ds.processes:
            self.producers[pr].add(pl)

    def value(self, product: str) -> float:
        prod = self.ds.products.get(product)
        if prod and prod.value_usd_t:
            return prod.value_usd_t
        content = prod.mo_content if prod and prod.mo_content else DEFAULT_MO_CONTENT
        return self.p.precio_mo_usd_kg * 1000 * content

    def conc_factor(self, product: str) -> float:
        prod = self.ds.products.get(product)
        if prod and prod.conc_factor:
            return prod.conc_factor
        content = prod.mo_content if prod and prod.mo_content else DEFAULT_MO_CONTENT
        return content / self.p.contenido_mo_concentrado / CONVERSION_YIELD

    def conc_cost(self, plant: str) -> float:
        pl = self.ds.plants.get(plant)
        return pl.conc_cost_usd_t if pl and pl.conc_cost_usd_t else 0.0

    def can_make(self, plant: str, product: str) -> bool:
        if not self.producers.get(product):
            return True
        return plant in self.producers[product]

    def candidate_plants(self, product: str) -> list[str]:
        return sorted(self.producers[product]) if self.producers.get(product) else sorted(self.ds.plants)

    def process(self, plant: str, product: str) -> dict[str, float]:
        pr = self.ds.processes.get((plant, product))
        p = self.p
        if pr is None and self.ds.processes and self.producers.get(product):
            self.ds.warn(f"La planta '{plant}' no tiene definido el proceso de '{product}' pero abastece "
                         f"esa demanda; se usan tiempos por defecto.")
        get = (lambda attr, default: getattr(pr, attr) if pr and getattr(pr, attr) is not None else default)
        costs = [x.cost_usd_t for (pl, prd), x in self.ds.processes.items() if prd == product and x.cost_usd_t is not None]
        return {
            "process_days": get("process_days", p.tiempo_proceso_dias),
            "process_sd": get("process_sd_days", p.desv_proceso_dias),
            "qa_days": get("qa_days", p.tiempo_qa_dias),
            "campaign_days": get("campaign_interval_days", p.intervalo_campana_dias),
            "capacity": get("capacity_t_month", None),
            "cost": get("cost_usd_t", mean(costs) if costs else 0.0),
        }

    def lane(self, plant: str, customer: str) -> tuple[float, float, float | None, str]:
        key = (plant, customer)
        if key in self._lanes:
            return self._lanes[key]
        c = self.ds.customers.get(customer)
        targets = [(norm(customer), "cliente")]
        if c and c.country:
            targets.append((norm(c.country), "pais"))
        if c and c.region:
            targets.append((norm(c.region), "region"))
        plant_lanes = [l for l in self.ds.lanes if norm(l.plant) == norm(plant)]
        result = None
        for target, source in targets:
            for l in plant_lanes:
                if norm(l.destination) == target:
                    sd = l.transit_sd_days if l.transit_sd_days is not None else self.p.desv_transito_dias
                    result = (l.transit_days, sd, l.freight_usd_t, source)
                    break
            if result:
                break
        if result is None:
            result = (self.p.dias_transito, self.p.desv_transito_dias, None, "defecto")
        self._lanes[key] = result
        return result

    def has_lane(self, plant: str, customer: str) -> bool:
        return self.lane(plant, customer)[3] != "defecto"


def month_range(months: list[str]) -> list[str]:
    if not months:
        return []
    y, m = map(int, min(months).split("-"))
    ye, me = map(int, max(months).split("-"))
    out = []
    while (y, m) <= (ye, me):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def demand_stats(ds: Dataset) -> tuple[dict[tuple[str, str], tuple[float, float, int]], list[str]]:
    all_months = sorted({m for s in ds.demand.values() for m in s})
    full = month_range(all_months)
    stats = {}
    for key, series in ds.demand.items():
        start = min(series)
        vals = [series.get(m, 0.0) for m in full if m >= start]
        d = mean(vals) if vals else 0.0
        if len(vals) >= 3:
            sd = stdev(vals)
        else:
            sd = 0.3 * d
            ds.warn(f"Línea {key[0]} / {key[1]}: menos de 3 meses de historia; se asume CV de 30%.")
        stats[key] = (d, sd, len(vals))
    return stats, full


def classify_abc(ds: Dataset, stats) -> dict[str, str]:
    vol = defaultdict(float)
    for (cu, _), (d, _, _) in stats.items():
        vol[cu] += d
    total = sum(vol.values()) or 1.0
    out, acc = {}, 0.0
    for cu, v in sorted(vol.items(), key=lambda kv: -kv[1]):
        acc += v
        share_before = (acc - v) / total
        out[cu] = "A" if share_before < 0.80 else "B" if share_before < 0.95 else "C"
    return out


def classify_xyz(cv: float) -> str:
    return "X" if cv <= 0.25 else "Y" if cv <= 0.50 else "Z"


def service_level_for(ds: Dataset, customer: str, abc: str) -> float:
    p = ds.params
    c = ds.customers.get(customer)
    if c and c.service_level:
        return c.service_level
    seg = (c.segment if c and c.segment else abc) or ""
    return {"A": p.nivel_servicio_a, "B": p.nivel_servicio_b, "C": p.nivel_servicio_c}.get(seg, p.nivel_servicio)


# ---------------------------------------------------------------- cálculo

def build_line(lk: Lookup, customer: str, product: str, plant: str, stat, abc: str) -> LineResult:
    ds, p = lk.ds, lk.p
    d_month, sd_month, n = stat
    c = ds.customers.get(customer)
    sl = service_level_for(ds, customer, abc)
    z = z_value(sl)
    T, sdT, freight, src = lk.lane(plant, customer)
    pr = lk.process(plant, product)
    dispatch = c.dispatch_interval_days if c and c.dispatch_interval_days else p.intervalo_despacho_dias
    firm = c.firm_order_days if c and c.firm_order_days is not None else p.pedido_firme_dias
    lead_prod = pr["process_days"] + pr["qa_days"]
    exposure = max(pr["campaign_days"] + lead_prod + T - firm, 0.0)
    d_day = d_month / p.dias_mes
    sd_day = sd_month / math.sqrt(p.dias_mes)
    var_d = exposure * sd_day ** 2
    var_p = d_day ** 2 * pr["process_sd"] ** 2 if exposure > 0 else 0.0
    var_t = d_day ** 2 * sdT ** 2 if exposure > 0 else 0.0
    var = var_d + var_p + var_t
    ss = z * math.sqrt(var)
    return LineResult(
        customer=customer, product=product, plant=plant,
        country=c.country if c else "", region=c.region if c else "",
        abc=abc, xyz=classify_xyz(sd_month / d_month if d_month else 0.0), service_level=sl, z=z,
        months=n, d_month=d_month, sd_month=sd_month, cv=sd_month / d_month if d_month else 0.0,
        transit_days=T, transit_sd=sdT, freight_usd_t=freight, lane_source=src,
        process_days=pr["process_days"], process_sd=pr["process_sd"], qa_days=pr["qa_days"],
        campaign_days=pr["campaign_days"], dispatch_days=dispatch, firm_days=firm, exposure_days=exposure,
        ss_t=ss, cycle_t=d_day * max(pr["campaign_days"], dispatch) / 2, wip_t=d_day * pr["process_days"],
        transit_stock_t=d_day * T if p.propiedad_en_transito else 0.0,
        share_demand=var_d / var if var else 0.0, share_process=var_p / var if var else 0.0,
        share_transit=var_t / var if var else 0.0,
        value_usd_t=lk.value(product), conversion_usd_t=pr["cost"],
        raw_material_usd_t=lk.conc_cost(plant) * lk.conc_factor(product),
    )


def current_assignment(ds: Dataset) -> dict[tuple[str, str], str | None]:
    out = {}
    for (cu, pr) in ds.demand:
        c = ds.customers.get(cu)
        out[(cu, pr)] = ds.line_plant.get((cu, pr)) or (c.plant if c and c.plant else None)
    return out


def _current_stock(ds: Dataset) -> dict[tuple[str, str, str], float]:
    """Promedio mensual por (planta, producto, ubicación)."""
    by_month: dict[tuple[str, str, str], dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for s in ds.stock:
        by_month[(s.plant, s.product, s.location)][s.month] += s.qty_t
    return {k: mean(v.values()) for k, v in by_month.items()}


def analyze(ds: Dataset, assignment: dict[tuple[str, str], str | None] | None = None) -> Analysis:
    lk = Lookup(ds)
    stats, months = demand_stats(ds)
    abc = classify_abc(ds, stats)
    assignment = dict(assignment or current_assignment(ds))

    # Líneas sin planta: la de menor costo anual entre las que fabrican el producto.
    for key, plant in list(assignment.items()):
        if not plant:
            cands = lk.candidate_plants(key[1]) or sorted(ds.plants)
            if not cands:
                ds.warn("No hay plantas definidas.")
                continue
            costs = {pl: _quick_cost(lk, build_line(lk, key[0], key[1], pl, stats[key], abc[key[0]])) for pl in cands}
            assignment[key] = min(costs, key=costs.get)

    lines = [build_line(lk, cu, pr, assignment[(cu, pr)], stats[(cu, pr)], abc[cu])
             for (cu, pr) in sorted(ds.demand) if assignment.get((cu, pr))]

    nodes: dict[tuple[str, str], NodeResult] = {}
    for ln in lines:
        nd = nodes.setdefault((ln.plant, ln.product), NodeResult(ln.plant, ln.product, ln.value_usd_t))
        nd.d_month += ln.d_month
        nd.n_customers += 1
        nd.ss_t += ln.ss_t ** 2              # se saca la raíz al final (pooling)
        nd.ss_unpooled_t += ln.ss_t
        nd.cycle_t += ln.cycle_t
        nd.wip_t += ln.wip_t
        nd.transit_t += ln.transit_stock_t
        nd.sd_month += ln.sd_month ** 2
        nd.lead_prod_days = ln.process_days + ln.qa_days
        nd.campaign_days = ln.campaign_days

    stock = _current_stock(ds)
    for (pl, pr, loc) in stock:
        if loc != "concentrado" and (pl, pr) not in nodes:
            nodes[(pl, pr)] = NodeResult(pl, pr, lk.value(pr))

    p = ds.params
    for (pl, pr), nd in nodes.items():
        nd.ss_t = math.sqrt(nd.ss_t)
        nd.sd_month = math.sqrt(nd.sd_month)
        d_day = nd.d_month / p.dias_mes
        nd.min_t = nd.ss_t
        nd.target_t = nd.ss_t + nd.cycle_t
        nd.max_t = nd.ss_t + 2 * nd.cycle_t
        nd.rop_t = nd.ss_t + d_day * nd.lead_prod_days
        cur = {loc: stock.get((pl, pr, loc), 0.0) for loc in ("pt", "consignacion", "proceso", "transito")}
        nd.has_current = any((pl, pr, loc) in stock for loc in cur)
        nd.cur_fg_t, nd.cur_consig_t = cur["pt"], cur["consignacion"]
        nd.cur_wip_t, nd.cur_transit_t = cur["proceso"], cur["transito"]
        if not nd.has_current:
            nd.status = "Sin dato"
        elif nd.d_month == 0:
            nd.status = "Stock sin demanda"
        elif nd.cur_fg_total_t < nd.min_t:
            nd.status = "Riesgo de quiebre"
        elif nd.cur_fg_total_t > nd.max_t * p.umbral_sobrestock:
            nd.status = "Sobrestock"
        elif nd.cur_fg_total_t > nd.max_t:
            nd.status = "Sobre máximo"
        else:
            nd.status = "En rango"

    plants = _plants(ds, lk, nodes, stock)
    return Analysis(ds=ds, assignment=assignment, lines=lines,
                    nodes=sorted(nodes.values(), key=lambda n: (n.plant, n.product)),
                    plants=plants, months=months)


def _quick_cost(lk: Lookup, ln: LineResult) -> float:
    freight = ln.freight_usd_t if ln.freight_usd_t is not None else (1e6 if lk.has_freight else 0.0)
    holding = (ln.ss_t + ln.cycle_t + ln.transit_stock_t) * ln.value_usd_t * lk.p.tasa_costo_capital
    return ln.annual_t * (freight + ln.conversion_usd_t + ln.raw_material_usd_t) + holding


def _plants(ds: Dataset, lk: Lookup, nodes: dict, stock: dict) -> list[PlantResult]:
    p = ds.params
    out = []
    z = z_value(p.nivel_servicio_concentrado)
    for name in sorted(set(ds.plants) | {pl for pl, _ in nodes}):
        plant = ds.plants.get(name)
        pr = PlantResult(plant=name, country=plant.country if plant else "", region=plant.region if plant else "",
                         capacity_t_month=plant.capacity_t_month if plant else None)
        var_c = 0.0
        for (pl, prod), nd in nodes.items():
            if pl != name:
                continue
            pr.load_t_month += nd.d_month
            f = lk.conc_factor(prod)
            pr.conc_d_day += nd.d_month / p.dias_mes * f
            var_c += (nd.sd_month / math.sqrt(p.dias_mes) * f) ** 2
        lt = plant.conc_lead_time_days if plant and plant.conc_lead_time_days else p.lead_time_concentrado_dias
        sd_lt = (plant.conc_lead_time_sd_days if plant and plant.conc_lead_time_sd_days is not None
                 else p.desv_lead_time_concentrado_dias)
        rc = (plant.conc_receipt_interval_days if plant and plant.conc_receipt_interval_days
              else p.intervalo_recepcion_concentrado_dias)
        pr.conc_lead_time = lt
        pr.conc_ss_t = safety_stock(z, lt, math.sqrt(var_c), pr.conc_d_day, sd_lt ** 2)
        pr.conc_cycle_t = pr.conc_d_day * rc / 2
        pr.conc_value_usd_t = p.precio_mo_usd_kg * 1000 * p.contenido_mo_concentrado
        conc_records = [v for (pl, _, loc), v in stock.items() if pl == name and loc == "concentrado"]
        if conc_records:
            pr.conc_current_t = sum(conc_records)
        elif plant and plant.conc_stock_t is not None:
            pr.conc_current_t = plant.conc_stock_t
        out.append(pr)
    return out
