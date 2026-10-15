"""Detalle por lote de producto terminado (export de Celonis).

Lee el export crudo (columnas Market, Stage, Plant, Warehouse, Material,
Batch, Inventory (Qty.)) o la hoja '11_Celonis_Detalle' del modelo, en
libras, y lo resume por etapa, mercado, planta, material y bodega. Es una
foto (un corte) del producto terminado: no se suma al stock del dashboard,
lo desagrega. Si el libro trae también la apertura mensual de PT
(Not Assigned / Assigned / In-Transit / Warehouse), se concilia con el
último mes disponible.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from openpyxl import load_workbook

from .loader import to_float
from .schema import norm

COLS = {
    "market": ("market", "mercado"),
    "stage": ("stage", "etapa"),
    "plant": ("plant", "planta"),
    "warehouse": ("warehouse", "bodega"),
    "material": ("material",),
    "batch": ("batch", "lote"),
    "qty": ("inventory_qty", "inventory", "cantidad_lb", "cantidad", "qty"),
}
STAGES = ("Plant", "Transit", "Ext.WH")


@dataclass
class Lot:
    market: str
    stage: str
    plant: str
    warehouse: str
    material: str
    batch: str
    lb: float | None


@dataclass
class CelonisData:
    lots: list[Lot]
    sheet: str
    warnings: list[str] = field(default_factory=list)

    @property
    def total(self) -> float:
        return sum(l.lb or 0 for l in self.lots) / 1e6

    def by(self, *attrs) -> dict[tuple, float]:
        out: dict[tuple, float] = defaultdict(float)
        for l in self.lots:
            out[tuple(getattr(l, a) for a in attrs)] += (l.lb or 0) / 1e6
        return dict(out)

    @property
    def missing_qty(self) -> int:
        return sum(1 for l in self.lots if not l.lb)


def _find_header(ws):
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows[:10]):
        hdr = [norm(c) if c is not None else "" for c in row]
        idx = {}
        for key, aliases in COLS.items():
            for j, h in enumerate(hdr):
                if any(h == a or h.startswith(a) for a in aliases):
                    idx.setdefault(key, j)
                    break
        if {"stage", "material", "qty"} <= idx.keys() and ("market" in idx or "plant" in idx):
            return idx, rows[i + 1:]
    return None, []


def load_celonis(source) -> CelonisData | None:
    """Devuelve el detalle si el libro trae una hoja tipo Celonis; si no, None."""
    wb = load_workbook(source, data_only=True, read_only=True)
    try:
        names = sorted(wb.sheetnames, key=lambda n: 0 if "celonis" in norm(n) and "detalle" in norm(n) else 1)
        for name in names:
            idx, body = _find_header(wb[name])
            if not idx:
                continue
            lots = []
            for r in body:
                def get(k):
                    j = idx.get(k)
                    return r[j] if j is not None and j < len(r) else None
                stage = get("stage")
                if not stage:
                    continue
                lots.append(Lot(str(get("market") or ""), str(stage), str(get("plant") or ""),
                                str(get("warehouse") or ""), str(get("material") or ""), str(get("batch") or ""),
                                to_float(get("qty"))))
            if lots:
                data = CelonisData(lots, name)
                if data.missing_qty:
                    data.warnings.append(f"{data.missing_qty} lotes sin cantidad en Celonis: revisar antes de cerrar.")
                data.warnings.append("Unidad asumida: libras de Mo contenido (confirmar).")
                return data
        return None
    finally:
        wb.close()
        if hasattr(source, "seek"):
            source.seek(0)


def load_sales_by_market(source) -> dict[str, float]:
    """Promedio de venta mensual por mercado (MMlb/mes) desde una hoja de ventas (p. ej. 15_Ventas_ZSD120)."""
    wb = load_workbook(source, data_only=True, read_only=True)
    try:
        for name in wb.sheetnames:
            if "ventas" not in norm(name):
                continue
            rows = list(wb[name].iter_rows(values_only=True))
            for i, row in enumerate(rows):
                hdr = [norm(c) if c is not None else "" for c in row]
                if hdr and hdr[0] == "mercado" and any(h.startswith("promedio") for h in hdr):
                    j = next(k for k, h in enumerate(hdr) if h.startswith("promedio"))
                    out = {}
                    for r in rows[i + 1:]:
                        if not r or not r[0]:
                            break
                        v = to_float(r[j]) if j < len(r) else None
                        if v is not None:
                            out[str(r[0])] = v
                    return out
        return {}
    finally:
        wb.close()
        if hasattr(source, "seek"):
            source.seek(0)


def transit_days_table(c: CelonisData, sales: dict[str, float]) -> dict:
    """Días de venta de cada mercado que representan su tránsito y su bodega (foto Celonis / venta mensual)."""
    mk = c.by("market", "stage")
    rows = []
    for m, v in sorted(sales.items(), key=lambda x: -x[1]):
        tr, wh = mk.get((m, "Transit"), 0.0), mk.get((m, "Ext.WH"), 0.0)
        d = v / 30 if v else 0
        rows.append([m, v, tr, tr / d if d else None, wh, wh / d if d else None])
    return dict(headers=["Mercado", "Venta MMlb/mes", "Tránsito MMlb", "Días de venta en tránsito",
                         "Bodega externa MMlb", "Días de venta en bodega"],
                rows=rows, fmts=[None, "0.00", "0.00", "0", "0.00", "0"])


def is_celonis_only(source) -> bool:
    wb = load_workbook(source, read_only=True)
    names = [norm(n) for n in wb.sheetnames]
    wb.close()
    if hasattr(source, "seek"):
        source.seek(0)
    if any("inventario_fisico" in n for n in names):
        return False
    return load_celonis(source) is not None


# ---------------------------------------------------------------- tablas

def stage_market_table(c: CelonisData) -> dict:
    piv = c.by("market", "stage")
    markets = sorted({m for m, _ in piv}, key=lambda m: -sum(v for (mm, _), v in piv.items() if mm == m))
    rows = []
    for m in markets:
        vals = [piv.get((m, s), 0.0) for s in STAGES]
        rows.append([m, *vals, sum(vals)])
    tot = [sum(r[i] for r in rows) for i in range(1, 5)]
    rows.append(["TOTAL", *tot])
    return dict(headers=["Mercado", "Planta MMlb", "Tránsito MMlb", "Bodega externa MMlb", "Total MMlb"],
                rows=rows, fmts=[None, "0.00", "0.00", "0.00", "0.00"])


def material_table(c: CelonisData) -> dict:
    piv = c.by("material", "stage")
    mats = sorted({m for m, _ in piv}, key=lambda m: -sum(v for (mm, _), v in piv.items() if mm == m))
    total = c.total or 1
    rows = []
    for m in mats:
        vals = [piv.get((m, s), 0.0) for s in STAGES]
        rows.append([m, *vals, sum(vals), sum(vals) / total])
    return dict(headers=["Material", "Planta MMlb", "Tránsito MMlb", "Bodega externa MMlb", "Total MMlb", "% del total"],
                rows=rows, fmts=[None, "0.00", "0.00", "0.00", "0.00", "0%"])


def warehouse_table(c: CelonisData) -> dict:
    piv = c.by("warehouse", "market", "stage")
    rows = [[w, m, v, v / (c.total or 1)] for (w, m, s), v in piv.items() if s == "Ext.WH" and v > 0]
    rows.sort(key=lambda r: -r[2])
    return dict(headers=["Bodega", "Mercado", "MMlb", "% del total"], rows=rows, fmts=[None, None, "0.00", "0%"])


def plant_table(c: CelonisData) -> dict:
    piv = c.by("plant", "stage")
    plants = sorted({p for p, _ in piv}, key=lambda p: -sum(v for (pp, _), v in piv.items() if pp == p))
    rows = []
    for p in plants:
        vals = [piv.get((p, s), 0.0) for s in STAGES]
        rows.append([p, *vals, sum(vals)])
    return dict(headers=["Planta", "Planta MMlb", "Tránsito MMlb", "Bodega externa MMlb", "Total MMlb"],
                rows=rows, fmts=[None, "0.00", "0.00", "0.00", "0.00"])


def reconciliation_table(c: CelonisData, month) -> dict:
    """Celonis vs apertura de PT del último mes (MonthRow de aggregate)."""
    st = {s: v for (s,), v in c.by("stage").items()}
    rows = []
    if month is not None:
        pairs = [
            ("Tránsito", st.get("Transit", 0.0), month.in_transit, "In-Transit"),
            ("Bodega externa", st.get("Ext.WH", 0.0), (month.warehouse or 0) + (month.consignment or 0),
             "Warehouse + Consignment"),
            ("Planta", st.get("Plant", 0.0), (month.not_assigned or 0) + (month.assigned or 0),
             "Not Assigned + Assigned"),
            ("Total PT", c.total, month.pt, "PT total"),
        ]
        for name, cel, dash, ref in pairs:
            rows.append([name, cel, dash, (dash - cel) if dash is not None else None, f"{ref} ({month.label})"])
    else:
        rows = [[s, st.get(s, 0.0), None, None, "sin apertura mensual de PT"] for s in STAGES]
    return dict(headers=["Etapa", "Celonis MMlb", "Dashboard MMlb", "Diferencia MMlb", "Referencia dashboard"],
                rows=rows, fmts=[None, "0.00", "0.00", "0.00", None])


def insights(c: CelonisData, month=None, sales: dict[str, float] | None = None) -> list[list]:
    """Lecturas principales en el mismo formato que las recomendaciones del modo agregado."""
    out: list[list] = []
    tot = c.total or 1
    st = {s: v for (s,), v in c.by("stage").items()}
    tr = st.get("Transit", 0.0)
    mk = c.by("market", "stage")
    far = sorted(((m, v) for (m, s), v in mk.items() if s == "Transit"), key=lambda x: -x[1])
    if tr:
        top = ", ".join(f"{m} {v:.2f}" for m, v in far[:3])
        share = sum(v for _, v in far[:2]) / tr
        out.append(["Alta", "Tránsito concentrado en rutas largas",
                    f"El tránsito es {tr:.2f} MMlb ({tr / tot:.0%} del detalle). Principales mercados: {top} MMlb; "
                    f"los dos primeros explican {share:.0%} del tránsito.",
                    "Cruzar el tránsito de cada mercado con su venta mensual (días de tránsito por ruta) y con las "
                    "fechas de llegada requeridas (ETA vs fecha cliente).", None, None])
        for m, v in far[:2]:
            wh = mk.get((m, "Ext.WH"), 0.0)
            if wh and v / wh > 5:
                out.append(["Media", f"{m}: tránsito muy alto frente a su bodega",
                            f"{v:.2f} MMlb en tránsito vs {wh:.2f} MMlb en bodega ({v / wh:.0f} veces).",
                            "Revisar si los embarques a este mercado salen con más anticipación de la necesaria o "
                            "si se vende en destino directo desde el barco.", None, None])
    if sales:
        for m, v, tr, dtr, wh, dwh in transit_days_table(c, sales)["rows"]:
            if dtr and dtr > 60:
                out.append(["Alta", f"{m}: {dtr:.0f} días de venta en tránsito",
                            f"Tránsito {tr:.2f} MMlb vs venta de {v:.2f} MMlb/mes: el tránsito cubre {dtr:.0f} días de "
                            f"venta del mercado (bodega: {dwh or 0:.0f} días).",
                            "Validar si ese tránsito corresponde a ventas de este mercado o a reposición de bodegas que "
                            "abastecen otros; revisar ETA, fecha requerida por el cliente y lotes con más antigüedad.",
                            None, None])
            elif dwh and dwh > 30:
                out.append(["Media", f"{m}: {dwh:.0f} días de venta en bodegas",
                            f"Bodegas {wh:.2f} MMlb vs venta de {v:.2f} MMlb/mes ({dwh:.0f} días).",
                            "Revisar el nivel objetivo de las bodegas del mercado y la frecuencia de reposición.",
                            None, None])
    mats = sorted(((m, v) for (m,), v in c.by("material").items()), key=lambda x: -x[1])
    if mats:
        m, v = mats[0]
        out.append(["Media", "Concentración por producto",
                    f"{m}: {v:.2f} MMlb ({v / tot:.0%} del detalle).",
                    "Usar esta apertura para completar 02_Cliente_SKU por tipo de producto (venta y variabilidad por "
                    "producto) y calcular stock de seguridad por familia.", None, None])
    whs = sorted(((w, v) for (w, s), v in c.by("warehouse", "stage").items() if s == "Ext.WH"), key=lambda x: -x[1])
    if whs:
        ext = st.get("Ext.WH", 0.0) or 1
        top = ", ".join(f"{w} {v:.2f}" for w, v in whs[:3])
        out.append(["Baja", "Bodegas externas", f"Principales: {top} MMlb ({sum(v for _, v in whs[:3]) / ext:.0%} "
                    "del stock en bodegas externas).",
                    "Definir un nivel objetivo por bodega (venta del mercado x días) y revisar la frecuencia de "
                    "reposición de las bodegas más grandes.", None, None])
    if month is not None and month.pt is not None:
        plant = st.get("Plant", 0.0)
        gap = month.pt - c.total
        out.append(["Media", "Conciliación con el dashboard",
                    f"Celonis {c.total:.2f} MMlb vs PT {month.label} {month.pt:.2f} MMlb (diferencia {gap:+.2f}). "
                    f"La etapa Planta de Celonis ({plant:.2f}) se parece al Not Assigned ({month.not_assigned or 0:.2f}): "
                    f"el Assigned en planta ({month.assigned or 0:.2f}) no aparece en el export.",
                    "Confirmar con el dueño del export si el stock asignado en planta queda fuera del perímetro de "
                    "Celonis y si las fechas de corte coinciden; no sumar Celonis al dashboard.", None, None])
    if c.missing_qty:
        out.append(["Baja", "Calidad de datos Celonis", f"{c.missing_qty} lotes sin cantidad.",
                    "Revisarlos en SAP antes de usar el detalle para decisiones.", None, None])
    return out
