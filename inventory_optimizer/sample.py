"""Plantilla Excel de entrada y un caso de ejemplo con datos ficticios.

El ejemplo representa una red típica de conversión de Mo (4 plantas,
4 productos, 20 clientes en 6 regiones) con desequilibrios a propósito para
que el análisis tenga algo que proponer. Los nombres son genéricos.
"""
from __future__ import annotations

import random
from dataclasses import fields

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

from .model import Params
from .schema import PARAM_HELP, SHEETS

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
REQ_FILL = PatternFill("solid", fgColor="C00000")


def _book(data: dict[str, list[list]] | None, params: Params | None = None) -> Workbook:
    wb = Workbook()
    ins = wb.active
    ins.title = "Instrucciones"
    ins["A1"] = "Plantilla de entrada - Optimizador de inventarios de productos de molibdeno"
    ins["A1"].font = Font(bold=True, size=14)
    ins["A3"] = ("Completa las hojas que tengas. Solo 'Demanda' es indispensable; lo que falte se reemplaza por "
                 "los supuestos de la hoja 'Parametros' y queda informado en el reporte. Los encabezados en rojo "
                 "son obligatorios dentro de cada hoja. Unidades: toneladas de producto, días, USD.")
    ins["A3"].alignment = Alignment(wrap_text=True)
    ins.column_dimensions["A"].width = 28
    ins.column_dimensions["B"].width = 30
    ins.column_dimensions["C"].width = 80
    ins.merge_cells("A3:C3")
    ins.row_dimensions[3].height = 48
    r = 5
    for sh in SHEETS:
        ins.cell(r, 1, sh.name).font = Font(bold=True)
        ins.cell(r, 3, sh.help)
        r += 1
        for col in sh.cols:
            ins.cell(r, 2, col.key + (" *" if col.required else ""))
            ins.cell(r, 3, col.help)
            r += 1
        r += 1

    for sh in SHEETS:
        ws = wb.create_sheet(sh.name)
        for j, col in enumerate(sh.cols, start=1):
            c = ws.cell(1, j, col.key)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = REQ_FILL if col.required else HEADER_FILL
            if col.help:
                c.comment = Comment(col.help, "plantilla")
            ws.column_dimensions[c.column_letter].width = max(14, len(col.key) + 3)
        for row in (data or {}).get(sh.name, []):
            ws.append(row)
        ws.freeze_panes = "A2"

    ws = wb.create_sheet("Parametros")
    ws.append(["parametro", "valor", "descripcion"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
    params = params or Params()
    for f in fields(Params):
        v = getattr(params, f.name)
        ws.append([f.name, int(v) if isinstance(v, bool) else v, PARAM_HELP.get(f.name, "")])
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 70
    return wb


def write_template(path) -> None:
    _book(None).save(path)


# ---------------------------------------------------------------- ejemplo

PRODUCTS = [
    # producto, % Mo, factor concentrado, tiempo proceso, desv, QA, campaña
    ("Óxido técnico (TMO)", 57, 1.18, 6, 1.0, 3, 15),
    ("Ferromolibdeno (FeMo)", 65, 1.36, 10, 2.0, 4, 30),
    ("Molibdato de sodio", 39.6, 0.83, 8, 1.5, 5, 45),
    ("Óxido puro (MoO3)", 66, 1.39, 12, 2.0, 5, 45),
]

PLANTS = [
    # planta, país, región, capacidad, LT conc, desv, intervalo recepción, stock conc, logística conc USD/t
    ("Planta Chile", "Chile", "Sudamérica", 2200, 10, 3, 7, 1400, 25),
    ("Planta México", "México", "Norteamérica", 900, 25, 5, 15, 500, 90),
    ("Planta Bélgica", "Bélgica", "Europa", 1300, 50, 10, 30, 3900, 150),
    ("Planta China", "China", "Asia", 800, 20, 6, 15, 120, 130),
]

MAKES = {
    "Planta Chile": {"Óxido técnico (TMO)": 310, "Ferromolibdeno (FeMo)": 420, "Molibdato de sodio": 900},
    "Planta México": {"Óxido técnico (TMO)": 340, "Ferromolibdeno (FeMo)": 450},
    "Planta Bélgica": {"Óxido técnico (TMO)": 420, "Ferromolibdeno (FeMo)": 520, "Molibdato de sodio": 980,
                       "Óxido puro (MoO3)": 1400},
    "Planta China": {"Óxido técnico (TMO)": 300, "Ferromolibdeno (FeMo)": 400, "Óxido puro (MoO3)": 1250},
}

# Tránsito por planta -> región: días, desviación, flete USD/t
LANES = {
    "Planta Chile": {"Sudamérica": (8, 2, 45), "Norteamérica": (22, 4, 95), "Europa": (32, 6, 115),
                     "Asia": (36, 7, 120), "Oceanía": (28, 5, 110), "África": (30, 7, 125)},
    "Planta México": {"Norteamérica": (6, 1, 50), "Sudamérica": (20, 4, 90), "Europa": (24, 5, 100),
                      "Asia": (32, 6, 115)},
    "Planta Bélgica": {"Europa": (4, 1, 35), "Norteamérica": (18, 4, 90), "Asia": (38, 9, 125),
                       "África": (16, 4, 85), "Medio Oriente": (20, 5, 95), "Sudamérica": (25, 5, 105)},
    "Planta China": {"Asia": (7, 2, 40), "Oceanía": (16, 3, 75), "Europa": (40, 10, 120),
                     "Medio Oriente": (22, 5, 90), "Norteamérica": (28, 6, 105)},
}

# cliente, país, región, planta actual, segmento, despacho, pedido firme, {producto: t/mes medias}, CV
CUSTOMERS = [
    ("Acería Rhein", "Alemania", "Europa", "Planta Bélgica", "A", 15, 15, {"Ferromolibdeno (FeMo)": 260}, 0.18),
    ("Inox Nord", "Suecia", "Europa", "Planta Chile", "A", 30, 0, {"Óxido técnico (TMO)": 220}, 0.22),
    ("Aceros Ibéricos", "España", "Europa", "Planta Bélgica", "B", 30, 0, {"Ferromolibdeno (FeMo)": 70}, 0.35),
    ("Catalizadores EU", "Países Bajos", "Europa", "Planta Bélgica", "B", 30, 30,
     {"Óxido puro (MoO3)": 35, "Molibdato de sodio": 20}, 0.30),
    ("Steel Midwest", "Estados Unidos", "Norteamérica", "Planta Chile", "A", 30, 0,
     {"Óxido técnico (TMO)": 300, "Ferromolibdeno (FeMo)": 110}, 0.20),
    ("Alloy Texas", "Estados Unidos", "Norteamérica", "Planta México", "B", 15, 0, {"Ferromolibdeno (FeMo)": 120}, 0.28),
    ("Lubricantes Ontario", "Canadá", "Norteamérica", "Planta Bélgica", "C", 30, 0, {"Molibdato de sodio": 12}, 0.55),
    ("Fundición Monterrey", "México", "Norteamérica", "Planta México", "B", 15, 0, {"Óxido técnico (TMO)": 90}, 0.25),
    ("Usiminas Sul", "Brasil", "Sudamérica", "Planta Chile", "A", 30, 15, {"Ferromolibdeno (FeMo)": 180}, 0.20),
    ("Aceros del Plata", "Argentina", "Sudamérica", "Planta Chile", "C", 30, 0, {"Ferromolibdeno (FeMo)": 25}, 0.50),
    ("Fertilizantes Andinos", "Perú", "Sudamérica", "Planta Chile", "C", 30, 0, {"Molibdato de sodio": 15}, 0.45),
    ("Nippon Special Steel", "Japón", "Asia", "Planta Chile", "A", 30, 30, {"Óxido técnico (TMO)": 350}, 0.15),
    ("Korea Stainless", "Corea del Sur", "Asia", "Planta Bélgica", "A", 30, 0, {"Ferromolibdeno (FeMo)": 200}, 0.20),
    ("Shanghai Alloys", "China", "Asia", "Planta China", "B", 15, 0,
     {"Óxido técnico (TMO)": 150, "Ferromolibdeno (FeMo)": 80}, 0.30),
    ("India Tool Steel", "India", "Asia", "Planta Chile", "B", 30, 0, {"Ferromolibdeno (FeMo)": 90}, 0.40),
    ("Taiwan Chem", "Taiwán", "Asia", "Planta Bélgica", "C", 45, 0, {"Óxido puro (MoO3)": 18}, 0.50),
    ("Pilbara Process", "Australia", "Oceanía", "Planta Chile", "C", 45, 0, {"Óxido técnico (TMO)": 30}, 0.60),
    ("SA Ferroalloys", "Sudáfrica", "África", "Planta Chile", "B", 45, 0, {"Ferromolibdeno (FeMo)": 60}, 0.40),
    ("Gulf Steel", "Emiratos Árabes", "Medio Oriente", "Planta Bélgica", "B", 30, 0, {"Ferromolibdeno (FeMo)": 75}, 0.35),
    ("Anatolia Çelik", "Turquía", "Medio Oriente", "Planta Bélgica", "C", 30, 0, {"Óxido técnico (TMO)": 40}, 0.45),
]

# Stock actual: planta, producto, ubicación, t
STOCK = [
    ("Planta Chile", "Óxido técnico (TMO)", "Planta", 1900),
    ("Planta Chile", "Óxido técnico (TMO)", "Proceso", 60),
    ("Planta Chile", "Óxido técnico (TMO)", "Tránsito", 820),
    ("Planta Chile", "Ferromolibdeno (FeMo)", "Planta", 160),
    ("Planta Chile", "Ferromolibdeno (FeMo)", "Proceso", 45),
    ("Planta Chile", "Ferromolibdeno (FeMo)", "Tránsito", 330),
    ("Planta Chile", "Molibdato de sodio", "Planta", 35),
    ("Planta México", "Óxido técnico (TMO)", "Planta", 70),
    ("Planta México", "Ferromolibdeno (FeMo)", "Planta", 160),
    ("Planta México", "Ferromolibdeno (FeMo)", "Proceso", 35),
    ("Planta México", "Ferromolibdeno (FeMo)", "Tránsito", 30),
    ("Planta Bélgica", "Ferromolibdeno (FeMo)", "Planta", 1350),
    ("Planta Bélgica", "Ferromolibdeno (FeMo)", "Proceso", 180),
    ("Planta Bélgica", "Ferromolibdeno (FeMo)", "Tránsito", 360),
    ("Planta Bélgica", "Ferromolibdeno (FeMo)", "Consignación", 140),
    ("Planta Bélgica", "Óxido técnico (TMO)", "Planta", 90),
    ("Planta Bélgica", "Molibdato de sodio", "Planta", 85),
    ("Planta Bélgica", "Óxido puro (MoO3)", "Planta", 110),
    ("Planta China", "Óxido técnico (TMO)", "Planta", 140),
    ("Planta China", "Ferromolibdeno (FeMo)", "Planta", 45),
    ("Planta China", "Óxido puro (MoO3)", "Planta", 60),
]

COST_FACTOR = {"Planta Chile": 1.0, "Planta México": 1.05, "Planta Bélgica": 1.25, "Planta China": 0.95}


def sample_data(seed: int = 7, months: int = 24, last=(2026, 9)) -> dict[str, list[list]]:
    rng = random.Random(seed)
    data: dict[str, list[list]] = {s.name: [] for s in SHEETS}
    for pl in PLANTS:
        data["Plantas"].append(list(pl))
    for name, mo, factor, *_ in PRODUCTS:
        data["Productos"].append([name, mo, factor, None])
    prod = {p[0]: p for p in PRODUCTS}
    for plant, makes in MAKES.items():
        for product, cost in makes.items():
            _, _, _, t, sd, qa, camp = prod[product]
            camp = camp if plant != "Planta Bélgica" else camp * 1.5   # Bélgica produce campañas largas
            data["Proceso"].append([plant, product, t, sd, qa, camp, None, cost])
    for cu, country, region, plant, seg, disp, firm, _, _ in CUSTOMERS:
        data["Clientes"].append([cu, country, region, plant, seg, None, disp, firm])
    for plant, regions in LANES.items():
        for region, (days, sd, fr) in regions.items():
            data["Transito"].append([plant, region, days, sd, fr])

    y, m = last
    month_list = []
    for _ in range(months):
        month_list.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    month_list.reverse()
    for cu, _, _, _, _, _, _, products, cv in CUSTOMERS:
        for product, base in products.items():
            for i, month in enumerate(month_list):
                season = 1 + 0.08 * (1 if int(month[5:]) in (3, 4, 5, 9, 10) else -0.6)
                trend = 1 + 0.004 * (i - months / 2)
                qty = max(0.0, rng.gauss(base * season * trend, base * cv))
                data["Demanda"].append([cu, product, month, round(qty, 1), None])
    for plant, product, loc, qty in STOCK:
        data["Stock"].append([plant, product, loc, qty, None])
    return data


def write_sample(path) -> None:
    _book(sample_data()).save(path)
