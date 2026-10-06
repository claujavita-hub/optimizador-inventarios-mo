"""Propuestas de mejora a partir del análisis y de la reasignación óptima."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from .allocation import AllocationResult
from .engine import safety_stock


@dataclass
class Recommendation:
    prioridad: str          # Alta / Media / Baja
    tipo: str
    ambito: str
    diagnostico: str
    accion: str
    impacto_t: float | None = None
    capital_usd: float | None = None      # capital de trabajo liberado (+) o requerido (-)
    ahorro_anual_usd: float | None = None

    @property
    def sort_key(self):
        return ({"Alta": 0, "Media": 1, "Baja": 2}[self.prioridad], -(abs(self.ahorro_anual_usd or 0)))


def _prio(usd: float, high: float, mid: float) -> str:
    return "Alta" if usd >= high else "Media" if usd >= mid else "Baja"


def build(alloc: AllocationResult) -> list[Recommendation]:
    a = alloc.base
    p = a.ds.params
    rate = p.tasa_costo_capital
    recs: list[Recommendation] = []
    sales_year = sum(n.d_month * n.value_usd_t for n in a.nodes) * 12 or 1.0
    high, mid = sales_year * 0.001, sales_year * 0.0002     # 0,1% y 0,02% de las ventas anuales

    # 1. Desvíos de stock por nodo planta-producto
    for n in a.nodes:
        amb = f"{n.plant} / {n.product}"
        if n.status in ("Sobrestock", "Sobre máximo"):
            exceso = n.cur_fg_total_t - n.target_t
            usd = exceso * n.value_usd_t
            recs.append(Recommendation(
                _prio(usd * rate, high, mid) if n.status == "Sobrestock" else "Baja", "Reducir sobrestock", amb,
                f"PT {n.cur_fg_total_t:,.0f} t vs máximo {n.max_t:,.0f} t ({n.cur_fg_total_t / max(n.d_month, 1e-9):.1f} "
                f"meses de venta); rota {n.cycles_month_current or 0:.2f} ciclos/mes vs {n.cycles_month_optimal or 0:.2f} óptimo.",
                f"Bajar a {n.target_t:,.0f} t: postergar o achicar la próxima campaña, priorizar despachos "
                f"desde este stock y revisar asignación de clientes de otras plantas a este nodo.",
                exceso, usd, usd * rate))
        elif n.status == "Riesgo de quiebre":
            falta = n.target_t - n.cur_fg_total_t
            usd = falta * n.value_usd_t
            recs.append(Recommendation(
                "Alta", "Riesgo de quiebre", amb,
                f"PT {n.cur_fg_total_t:,.0f} t bajo el stock de seguridad {n.ss_t:,.0f} t "
                f"(cubre {n.cur_fg_total_t / max(n.d_month / 30, 1e-9):.0f} días de demanda).",
                f"Adelantar campaña: lanzar producción al llegar a {n.rop_t:,.0f} t (punto de reorden) y "
                f"reponer hasta {n.target_t:,.0f} t. Evaluar despacho desde otra planta mientras tanto.",
                -falta, -usd, None))
        elif n.status == "Stock sin demanda":
            usd = n.current_total_t * n.value_usd_t
            recs.append(Recommendation(
                _prio(usd * rate, high, mid), "Stock sin demanda", amb,
                f"{n.current_total_t:,.0f} t en stock sin demanda asignada a esta planta en la historia.",
                "Reasignar a clientes de otras plantas, reprocesar a otro producto o vender en spot.",
                n.current_total_t, usd, usd * rate))
        if n.has_current and n.cur_transit_t and n.transit_t and n.cur_transit_t > 1.3 * n.transit_t:
            recs.append(Recommendation(
                "Media", "Tránsito excesivo", amb,
                f"Stock en tránsito {n.cur_transit_t:,.0f} t vs {n.transit_t:,.0f} t esperado por demanda y "
                f"días de tránsito: hay despachos adelantados o demoras logísticas.",
                "Revisar demoras en puerto/aduana y sincronizar fechas de despacho con las necesidades del cliente.",
                n.cur_transit_t - n.transit_t, (n.cur_transit_t - n.transit_t) * n.value_usd_t, None))

    # 2. Concentrado
    for pl in a.plants:
        if pl.conc_current_t is None or not pl.conc_d_day:
            continue
        diff = pl.conc_current_t - pl.conc_target_t
        usd = diff * pl.conc_value_usd_t
        if pl.conc_current_t < pl.conc_ss_t:
            recs.append(Recommendation(
                "Alta", "Concentrado bajo seguridad", pl.plant,
                f"Concentrado {pl.conc_current_t:,.0f} t = {pl.conc_coverage_days:.0f} días de consumo; "
                f"seguridad requerida {pl.conc_ss_t:,.0f} t (lead time {pl.conc_lead_time:.0f} días).",
                f"Acelerar embarques de concentrado o reprogramar la carga de tostación; objetivo {pl.conc_target_t:,.0f} t.",
                diff, usd, None))
        elif pl.conc_current_t > (pl.conc_ss_t + 2 * pl.conc_cycle_t) * p.umbral_sobrestock:
            recs.append(Recommendation(
                _prio(usd * rate, high, mid), "Exceso de concentrado", pl.plant,
                f"Concentrado {pl.conc_current_t:,.0f} t = {pl.conc_coverage_days:.0f} días de consumo; "
                f"objetivo {pl.conc_target_t:,.0f} t.",
                "Diferir recepciones de concentrado o aumentar la frecuencia de embarques con lotes menores.",
                diff, usd, usd * rate))

    # 3. Reasignaciones de clientes entre plantas
    for m in alloc.moves:
        recs.append(Recommendation(
            _prio(m.saving, high, mid) if m.saving > 0 else "Media", "Reasignar abastecimiento",
            f"{m.customer} / {m.product}",
            f"Hoy desde {m.from_plant} ({m.transit_from:.0f} días de tránsito, costo anual {m.cost_from:,.0f} USD).",
            f"Abastecer desde {m.to_plant} ({m.transit_to:.0f} días; {m.reason}).",
            m.t_month, None, m.saving))
    for msg in alloc.overloaded:
        recs.append(Recommendation("Alta", "Capacidad insuficiente", msg.split(" sin ")[0], msg,
                                   "Ampliar capacidad, agregar turnos/campañas o tercerizar conversión."))

    # 4. Utilización de capacidad
    for pl in alloc.proposed.plants:
        u = pl.utilization
        if u is None:
            continue
        if u > p.utilizacion_alta:
            recs.append(Recommendation(
                "Alta" if u > 1 else "Media", "Capacidad saturada", pl.plant,
                f"Utilización {u:.0%} ({pl.load_t_month:,.0f} de {pl.capacity_t_month:,.0f} t/mes) con la asignación propuesta.",
                "Con alta utilización las campañas no se pueden acortar y el stock de seguridad real debe ser mayor; "
                "mover volumen flexible a plantas con holgura."))
        elif u < p.utilizacion_baja:
            recs.append(Recommendation(
                "Baja", "Capacidad ociosa", pl.plant,
                f"Utilización {u:.0%} ({pl.load_t_month:,.0f} de {pl.capacity_t_month:,.0f} t/mes).",
                "Usar la holgura para campañas más cortas y frecuentes (menos stock de ciclo) o captar demanda "
                "de plantas saturadas."))

    # 5. Palancas de política: campañas más frecuentes, pedidos firmes, tránsito
    by_node = defaultdict(list)
    for ln in a.lines:
        by_node[(ln.plant, ln.product)].append(ln)
    for (plant, product), lines in by_node.items():
        R = lines[0].campaign_days
        if R < 20:
            continue
        cycle_now = sum(ln.cycle_t for ln in lines)
        cycle_new = sum(ln.d_month / p.dias_mes * max(R / 2, ln.dispatch_days) / 2 for ln in lines)
        ss_now = sum(ln.ss_t ** 2 for ln in lines) ** 0.5
        ss_new = sum(safety_stock(ln.z, max(ln.exposure_days - R / 2, 0), ln.sd_month / p.dias_mes ** 0.5,
                                  ln.d_month / p.dias_mes, ln.process_sd ** 2 + ln.transit_sd ** 2) ** 2
                     for ln in lines) ** 0.5
        red = (cycle_now - cycle_new) + (ss_now - ss_new)
        usd = red * lines[0].value_usd_t
        if usd * rate >= mid:
            recs.append(Recommendation(
                _prio(usd * rate, high, mid), "Campañas más frecuentes", f"{plant} / {product}",
                f"Campaña cada {R:.0f} días: stock de ciclo {cycle_now:,.0f} t y seguridad {ss_now:,.0f} t.",
                f"Producir cada {R / 2:.0f} días reduce ~{red:,.0f} t de stock. Conviene si el costo extra de "
                f"cambio de campaña (setup, limpieza, QA) es menor a {usd * rate:,.0f} USD/año.",
                red, usd, usd * rate))

    firm_gain = []
    for ln in a.lines:
        if ln.firm_days >= 30 or ln.exposure_days <= 0:
            continue
        new_e = max(ln.exposure_days - (30 - ln.firm_days), 0)
        ss_new = safety_stock(ln.z, new_e, ln.sd_month / p.dias_mes ** 0.5, ln.d_month / p.dias_mes,
                              ln.process_sd ** 2 + ln.transit_sd ** 2)
        firm_gain.append((ln, ln.ss_t - ss_new))
    firm_gain.sort(key=lambda x: -x[1] * x[0].value_usd_t)
    for ln, red in firm_gain[:5]:
        usd = red * ln.value_usd_t
        if usd * rate < mid:
            continue
        recs.append(Recommendation(
            _prio(usd * rate, high, mid), "Pedido firme / forecast colaborativo", f"{ln.customer} / {ln.product}",
            f"Demanda {ln.xyz} (CV {ln.cv:.0%}) con {ln.firm_days:.0f} días de aviso firme; "
            f"exige {ln.ss_t:,.0f} t de seguridad.",
            "Acordar nominación firme con 30 días de anticipación (o VMI / forecast compartido): "
            f"reduce el stock de seguridad en ~{red:,.0f} t.",
            red, usd, usd * rate))

    transit_heavy = [ln for ln in a.lines if ln.share_transit > 0.4 and ln.transit_days >= 20 and ln.ss_t > 0]
    by_region = defaultdict(list)
    for ln in transit_heavy:
        by_region[(ln.plant, ln.region or ln.country or ln.customer)].append(ln)
    for (plant, region), lines in by_region.items():
        ss_t = sum(ln.ss_t * ln.share_transit for ln in lines)
        usd = ss_t * lines[0].value_usd_t
        if usd * rate < mid:
            continue
        recs.append(Recommendation(
            _prio(usd * rate, high, mid), "Variabilidad de tránsito", f"{plant} -> {region}",
            f"{len(lines)} línea(s) donde la variabilidad del tránsito explica >40% del stock de seguridad "
            f"(~{ss_t:,.0f} t).",
            "Negociar ventanas de zarpe fijas / servicio directo con la naviera, o evaluar un stock regional "
            "(hub o consignación) que absorba la variabilidad cerca del cliente.",
            ss_t, usd, usd * rate))

    pooling = sum(n.ss_unpooled_t - n.ss_t for n in a.nodes)
    if pooling > 0:
        val = sum((n.ss_unpooled_t - n.ss_t) * n.value_usd_t for n in a.nodes)
        recs.append(Recommendation(
            "Baja", "Pooling de riesgo (informativo)", "Todas las plantas",
            f"Manejar el stock de seguridad por planta-producto (y no por cliente) ahorra {pooling:,.0f} t "
            f"({val:,.0f} USD) frente a reservar stock por cliente.",
            "Mantener stock de seguridad común por planta-producto; evitar reservas por cliente salvo contratos "
            "que lo exijan.", pooling, val, val * rate))

    return sorted(recs, key=lambda r: r.sort_key)
