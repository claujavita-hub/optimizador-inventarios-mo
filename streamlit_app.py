"""Punto de entrada para Streamlit Cloud (archivo principal por defecto)."""
import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).resolve().parent / "inventory_optimizer" / "streamlit_app.py"), run_name="__main__")
