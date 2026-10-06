"""Uso:
    python -m inventory_optimizer plantilla  entrada.xlsx        # plantilla vacía
    python -m inventory_optimizer ejemplo    ejemplo.xlsx        # caso de ejemplo
    python -m inventory_optimizer analizar   entrada.xlsx [-o reporte.xlsx]
"""
from __future__ import annotations

import argparse
from pathlib import Path

from . import run, write_report
from .aggregate import analyze_aggregate, is_aggregate_workbook, kpis, load_aggregate, write_aggregate_report
from .sample import write_sample, write_template


def _fmt(v, unit=""):
    return "s/d" if v is None else f"{v:,.2f}{unit}"


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(prog="python -m inventory_optimizer", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("comando", choices=["plantilla", "ejemplo", "analizar"])
    ap.add_argument("archivo")
    ap.add_argument("-o", "--salida", help="Reporte Excel (default: <entrada>_reporte.xlsx)")
    args = ap.parse_args(argv)

    if args.comando == "plantilla":
        write_template(args.archivo)
        print(f"Plantilla creada: {args.archivo}")
        return
    if args.comando == "ejemplo":
        write_sample(args.archivo)
        print(f"Ejemplo creado: {args.archivo}")
        return

    out = args.salida or str(Path(args.archivo).with_name(Path(args.archivo).stem + "_reporte.xlsx"))
    if is_aggregate_workbook(args.archivo):
        _aggregate(args.archivo, out)
        return
    alloc, recs = run(args.archivo)
    write_report(alloc, recs, out)
    t = alloc.base.totals()
    print(f"Ciclos por mes   actual {_fmt(t['ciclos_mes_actual'])} | óptimo {_fmt(t['ciclos_mes_optimo'])}")
    print(f"Días inventario  actual {_fmt(t['doi_actual'])} | óptimo {_fmt(t['doi_optimo'])}")
    print(f"Capital liberable {_fmt(t['capital_liberable_usd'], ' USD')} | ahorro anual {_fmt(t['ahorro_anual_usd'], ' USD')}")
    print(f"Reasignaciones propuestas: {len(alloc.moves)}")
    print("Principales recomendaciones:")
    for r in recs[:8]:
        print(f"  [{r.prioridad}] {r.tipo} - {r.ambito}")
    if alloc.base.ds.warnings:
        print(f"{len(alloc.base.ds.warnings)} advertencias de datos (ver hoja Supuestos).")
    print(f"Reporte: {out}")


def _aggregate(path: str, out: str) -> None:
    r = analyze_aggregate(load_aggregate(path))
    write_aggregate_report(r, out)
    print("Formato agregado (stock por etapa + ventas mensuales).")
    for row in kpis(r):
        print(f"  {row[0]:<40} actual {_fmt(row[1])} | demostrado {_fmt(row[2])} | técnico {_fmt(row[3])}")
    print("Principales recomendaciones:")
    for rec in r.recs[:8]:
        print(f"  [{rec[0]}] {rec[1]}")
    if r.data.warnings:
        print(f"{len(r.data.warnings)} advertencias (ver hoja Supuestos).")
    print(f"Reporte: {out}")


if __name__ == "__main__":
    main()
