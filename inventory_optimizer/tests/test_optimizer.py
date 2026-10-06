"""python -m unittest discover -s inventory_optimizer/tests"""
from __future__ import annotations

import io
import math
from statistics import mean
import sys
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from inventory_optimizer import analyze, build_recommendations, load_excel, optimize, write_report  # noqa: E402
from inventory_optimizer.engine import z_value  # noqa: E402
from inventory_optimizer.loader import to_fraction, to_month  # noqa: E402
from inventory_optimizer.sample import _book, sample_data  # noqa: E402


def _xlsx(sheets: dict[str, list[list]]) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class ParsingTest(unittest.TestCase):
    def test_months(self):
        self.assertEqual(to_month("2025-03"), "2025-03")
        self.assertEqual(to_month("03/2025"), "2025-03")
        self.assertEqual(to_month("ene-25"), "2025-01")
        self.assertEqual(to_month("Septiembre 2024"), "2024-09")

    def test_fraction(self):
        self.assertAlmostEqual(to_fraction("57%"), 0.57)
        self.assertAlmostEqual(to_fraction(57), 0.57)
        self.assertAlmostEqual(to_fraction("0,95"), 0.95)


class SingleLineTest(unittest.TestCase):
    """Un cliente, una planta: verifica las fórmulas a mano."""

    def setUp(self):
        demand = [["Cliente", "Producto", "Mes", "Toneladas"]]
        demand += [["C1", "FeMo", f"2025-{m:02d}", q] for m, q in enumerate([90, 110, 100, 95, 105, 100], 1)]
        self.ds = load_excel(_xlsx({
            "Plantas": [["Planta", "País"], ["P1", "Chile"]],
            "Productos": [["Producto", "Contenido Mo %"], ["FeMo", 65]],
            "Proceso": [["Planta", "Producto", "Tiempo proceso dias", "Desv proceso dias", "Tiempo QA dias",
                         "Intervalo campana dias"], ["P1", "FeMo", 10, 0, 5, 30]],
            "Clientes": [["Cliente", "Planta asignada", "Nivel servicio", "Intervalo despacho dias"],
                         ["C1", "P1", "95%", 15]],
            "Demanda": demand,
            "Tránsito": [["Origen", "Destino", "Dias transito", "Desv transito dias", "Flete"],
                         ["P1", "C1", 20, 4, 100]],
            "Stock": [["Planta", "Producto", "Ubicación", "Cantidad"], ["P1", "FeMo", "Planta", 400]],
        }))

    def test_formulas(self):
        a = analyze(self.ds)
        ln = a.lines[0]
        d_day = 100 / 30
        sd_day = ln.sd_month / math.sqrt(30)
        exposure = 30 + 10 + 5 + 20
        self.assertEqual(ln.exposure_days, exposure)
        expected_ss = z_value(0.95) * math.sqrt(exposure * sd_day ** 2 + d_day ** 2 * 4 ** 2)
        self.assertAlmostEqual(ln.ss_t, expected_ss, places=6)
        self.assertAlmostEqual(ln.cycle_t, d_day * 30 / 2)
        self.assertAlmostEqual(ln.wip_t, d_day * 10)
        self.assertAlmostEqual(ln.transit_stock_t, d_day * 20)
        n = a.nodes[0]
        self.assertAlmostEqual(n.cycles_month_current, 100 / 400)
        self.assertEqual(n.status, "Sobrestock")
        self.assertEqual(ln.lane_source, "cliente")
        self.assertAlmostEqual(ln.value_usd_t, 50 * 1000 * 0.65)

    def test_firm_orders_reduce_exposure(self):
        self.ds.customers["C1"].firm_order_days = 30
        self.assertEqual(analyze(self.ds).lines[0].exposure_days, 35)


class WideDemandTest(unittest.TestCase):
    def test_wide_format_and_defaults(self):
        ds = load_excel(_xlsx({"Demanda": [["Cliente", "Producto", "2025-01", "2025-02", "2025-03"],
                                           ["C1", "TMO", 10, 12, 11], ["C2", "TMO", 5, 0, 7]]}))
        self.assertEqual(len(ds.demand), 2)
        self.assertEqual(ds.demand[("C2", "TMO")]["2025-02"], 0)
        self.assertTrue(any("Proceso" in w for w in ds.warnings))


class AllocationTest(unittest.TestCase):
    def test_capacity_respected_and_sample_runs(self):
        buf = io.BytesIO()
        _book(sample_data()).save(buf)
        buf.seek(0)
        ds = load_excel(buf)
        alloc = optimize(analyze(ds))
        for pl in alloc.proposed.plants:
            if pl.capacity_t_month:
                self.assertLessEqual(pl.load_t_month, pl.capacity_t_month + 1e-6)
        self.assertTrue(alloc.moves)
        self.assertTrue(all(m.saving > ds.params.ahorro_minimo_reasignacion_usd or "capacidad" in m.reason
                            for m in alloc.moves))
        recs = build_recommendations(alloc)
        self.assertTrue(any(r.tipo == "Reducir sobrestock" for r in recs))
        out = io.BytesIO()
        write_report(alloc, recs, out)
        self.assertGreater(len(out.getvalue()), 10_000)


if __name__ == "__main__":
    unittest.main()


class AggregateTest(unittest.TestCase):
    """Formato agregado (stock por etapa + ventas mensuales), con datos sintéticos."""

    def _book(self, stage_days=None):
        months = ["Ene-26", "Feb-26", "Mar-26", "Abr-26", "May-26", "Jun-26"]
        sales = [8, 9, 7, 8, 9, 7]
        fis = [["Inventario físico"], [], ["Mes", "Stock total MMlb", "Producto terminado MMlb", "Pre-PT MMlb",
                                           "Own Sales MMlb", "Safety stock MMlb"]]
        pt = [["PT"], [], ["Mes cierre", "Not Assigned", "Assigned", "In-Transit", "Warehouse", "Consignment / SS",
                           "PT total", "Own Sales mes siguiente"]]
        for i, m in enumerate(months):
            fis.append([m, 30 + i, 20, 10 + i, sales[i], 0.4])
            pt.append([m, 3.6, 4, 9, 3, 0.4, 20, sales[i + 1] if i + 1 < len(sales) else None])
        fis.append(["Promedio", 1, 1, 1, 1])
        ciclo = [[], [], ["Etapa", "Días actuales", "Días objetivo"]]
        for name in ("Compra/espera embarque", "Tránsito", "Puerto/aduana/recepción", "Espera pre-proceso",
                     "Proceso productivo", "Producto terminado/espera", "Despacho/entrega"):
            ciclo.append([name, None, (stage_days or {}).get(name)])
        return _xlsx({"00_Instrucciones": [["Modelo de prueba"]],
                      "01_Supuestos": [["Supuestos"], [], ["Variable", "Valor"], ["Valor promedio por libra", 30],
                                       ["Tasa deuda incremental", 0.08]],
                      "08_Inventario_Fisico": fis, "09_PT_Cobertura": pt, "03_Ciclo_Transito": ciclo})

    def test_benchmark_and_technical(self):
        from inventory_optimizer.aggregate import (analyze_aggregate, group_table, is_aggregate_workbook,
                                                   load_aggregate, write_aggregate_report)
        src = self._book()
        self.assertTrue(is_aggregate_workbook(src))
        data = load_aggregate(src)
        self.assertEqual(len(data.months), 6)
        self.assertEqual(data.params.precio_usd_lb, 30)
        self.assertAlmostEqual(data.params.tasa_costo_capital, 0.08)
        r = analyze_aggregate(data)
        self.assertAlmostEqual(r.d_month, 8.0)
        self.assertAlmostEqual(r.actual_total, mean(30 + i for i in range(6)), places=6)
        self.assertLess(r.benchmark_total, r.actual_total)
        self.assertIsNone(r.technical_total)   # sin días por etapa no hay nivel técnico
        out = io.BytesIO()
        write_aggregate_report(r, out)
        self.assertGreater(len(out.getvalue()), 5_000)

        r2 = analyze_aggregate(load_aggregate(self._book({"Despacho/entrega": 30})))
        disp = next(g for g in r2.groups if g.name.startswith("Asignado"))
        self.assertAlmostEqual(disp.technical, 8.0 / 30 * 30)
        self.assertIsNotNone(group_table(r2)["rows"][2][7])
