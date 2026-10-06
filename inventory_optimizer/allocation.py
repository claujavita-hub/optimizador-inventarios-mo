"""Reasignación cliente-producto -> planta.

Heurística de "arrepentimiento" (regret) con capacidades: en cada paso se
asigna la línea que más perdería si no obtiene su mejor planta, respetando
la capacidad total de cada planta y la específica de cada producto. El
costo anual de una línea en una planta es:

    flete + conversión + logística del concentrado + costo de capital del stock que la línea exige
    (seguridad + ciclo + tránsito)

Cambiar a un cliente de planta tiene un costo implícito (calificación del
producto, contratos, logística), así que solo se propone si el ahorro supera
`ahorro_minimo_reasignacion_usd`; ese umbral se suma como penalización a
toda planta distinta de la actual.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from .engine import Analysis, Lookup, analyze, build_line, classify_abc, demand_stats


@dataclass
class Move:
    customer: str
    product: str
    from_plant: str
    to_plant: str
    t_month: float
    cost_from: float
    cost_to: float
    transit_from: float
    transit_to: float
    reason: str

    @property
    def saving(self) -> float:
        return self.cost_from - self.cost_to


@dataclass
class AllocationResult:
    base: Analysis
    proposed: Analysis
    moves: list[Move]
    overloaded: list[str]


def _annual_cost(lk: Lookup, ln) -> float | None:
    if ln.freight_usd_t is None and lk.has_freight:
        return None   # sin flete conocido para esa ruta no se puede comparar
    rate = lk.p.tasa_costo_capital
    holding = (ln.ss_t + ln.cycle_t + ln.transit_stock_t) * ln.value_usd_t * rate
    return ln.annual_t * ((ln.freight_usd_t or 0.0) + ln.conversion_usd_t + ln.raw_material_usd_t) + holding


def optimize(base: Analysis) -> AllocationResult:
    ds = base.ds
    lk = Lookup(ds)
    stats, _ = demand_stats(ds)
    abc = classify_abc(ds, stats)
    hurdle = ds.params.ahorro_minimo_reasignacion_usd

    # Costo de cada línea en cada planta candidata.
    options: dict[tuple[str, str], dict[str, tuple[float, object]]] = {}
    for ln in base.lines:
        key = (ln.customer, ln.product)
        opts = {}
        for pl in lk.candidate_plants(ln.product):
            if pl != ln.plant and not (lk.has_lane(pl, ln.customer) or not ds.lanes):
                continue   # ruta no informada: no se propone
            cand = ln if pl == ln.plant else build_line(lk, ln.customer, ln.product, pl, stats[key], abc[ln.customer])
            cost = _annual_cost(lk, cand)
            if cost is None and pl != ln.plant:
                continue
            opts[pl] = (cost if cost is not None else float("inf"), cand)
        options[key] = opts

    plant_cap = {pl: p.capacity_t_month for pl, p in ds.plants.items() if p.capacity_t_month}
    proc_cap = {k: v.capacity_t_month for k, v in ds.processes.items() if v.capacity_t_month}
    load = defaultdict(float)
    proc_load = defaultdict(float)
    volume = {(ln.customer, ln.product): ln.d_month for ln in base.lines}
    current = {(ln.customer, ln.product): ln.plant for ln in base.lines}

    def fits(key, pl) -> bool:
        q = volume[key]
        if pl in plant_cap and load[pl] + q > plant_cap[pl] + 1e-9:
            return False
        if (pl, key[1]) in proc_cap and proc_load[(pl, key[1])] + q > proc_cap[(pl, key[1])] + 1e-9:
            return False
        return True

    def eff_cost(key, pl) -> float:
        cost = options[key][pl][0]
        return cost if pl == current[key] else cost + hurdle

    proposed: dict[tuple[str, str], str] = {}
    overloaded: list[str] = []
    pending = set(options)
    while pending:
        best_key, best_pl, best_regret = None, None, -1.0
        for key in pending:
            ranked = sorted((eff_cost(key, pl), pl) for pl in options[key] if fits(key, pl))
            if not ranked:
                continue
            regret = (ranked[1][0] - ranked[0][0]) if len(ranked) > 1 else float("inf")
            regret = regret if regret != float("inf") else 1e18 + volume[key]
            if regret > best_regret:
                best_key, best_pl, best_regret = key, ranked[0][1], regret
        if best_key is None:   # nadie cabe: se quedan en su planta actual (sobrecarga)
            for key in sorted(pending):
                proposed[key] = current[key]
                load[current[key]] += volume[key]
                proc_load[(current[key], key[1])] += volume[key]
                overloaded.append(f"{key[0]} / {key[1]} sin capacidad disponible; se mantiene en {current[key]}.")
            break
        proposed[best_key] = best_pl
        load[best_pl] += volume[best_key]
        proc_load[(best_pl, best_key[1])] += volume[best_key]
        pending.discard(best_key)

    moves = []
    for key, to_pl in sorted(proposed.items()):
        frm = current[key]
        if to_pl == frm:
            continue
        c_from, ln_from = options[key][frm]
        c_to, ln_to = options[key][to_pl]
        reasons = []
        if (ln_to.freight_usd_t or 0) < (ln_from.freight_usd_t or 0):
            reasons.append(f"flete {ln_from.freight_usd_t or 0:,.0f} -> {ln_to.freight_usd_t or 0:,.0f} USD/t")
        if ln_to.transit_days < ln_from.transit_days:
            reasons.append(f"tránsito {ln_from.transit_days:.0f} -> {ln_to.transit_days:.0f} días")
        if ln_to.conversion_usd_t < ln_from.conversion_usd_t:
            reasons.append(f"conversión {ln_from.conversion_usd_t:,.0f} -> {ln_to.conversion_usd_t:,.0f} USD/t")
        if ln_to.raw_material_usd_t < ln_from.raw_material_usd_t:
            reasons.append(f"logística concentrado {ln_from.raw_material_usd_t:,.0f} -> "
                           f"{ln_to.raw_material_usd_t:,.0f} USD/t")
        if c_to >= c_from:
            reasons.append(f"libera capacidad en {frm}")
        moves.append(Move(key[0], key[1], frm, to_pl, volume[key], c_from, c_to,
                          ln_from.transit_days, ln_to.transit_days, "; ".join(reasons)))

    proposed_analysis = analyze(ds, proposed) if moves else base
    return AllocationResult(base=base, proposed=proposed_analysis, moves=moves, overloaded=overloaded)
