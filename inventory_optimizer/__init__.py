"""Optimizador de inventarios para la conversión de productos de molibdeno.

Lee un Excel con plantas, productos, procesos, clientes, demanda, tránsitos
y stock, y calcula la rotación actual, el stock óptimo por planta-producto, los ciclos
por mes y propuestas de mejora (incluida la reasignación de clientes entre
plantas).
"""
from __future__ import annotations

from .allocation import AllocationResult, optimize
from .engine import Analysis, analyze
from .loader import load_excel
from .recommendations import Recommendation, build as build_recommendations
from .report import write_report


def run(source) -> tuple[AllocationResult, list[Recommendation]]:
    """Excel de entrada (ruta o archivo abierto) -> (asignación base/propuesta, recomendaciones)."""
    ds = load_excel(source)
    alloc = optimize(analyze(ds))
    return alloc, build_recommendations(alloc)


__all__ = ["run", "load_excel", "analyze", "optimize", "build_recommendations", "write_report",
           "Analysis", "AllocationResult", "Recommendation"]
