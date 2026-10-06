"""App web: subir el Excel, revisar el análisis y descargar el reporte.

    streamlit run inventory_optimizer/streamlit_app.py
"""
from __future__ import annotations

import sys
from pathlib import Path

# Streamlit agrega la carpeta de este script (la del paquete) a sys.path. Eso
# expone report.py, engine.py, etc. como módulos sueltos y, si alguien los
# importa así, fallan sus imports relativos (ImportError visto en Streamlit
# Cloud). Se quita esa carpeta y se importa el paquete desde la raíz del repo.
_PKG = Path(__file__).resolve().parent
sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != _PKG]
sys.path.insert(0, str(_PKG.parent))

import io  # noqa: E402
from dataclasses import replace  # noqa: E402

import pandas as pd  # noqa: E402
import streamlit as st  # noqa: E402

from inventory_optimizer import analyze, build_recommendations, load_excel, optimize, write_report  # noqa: E402
from inventory_optimizer.report import (PLANT_HEADERS, detail_tables, kpi_rows, load_table,  # noqa: E402
                                        plant_summary, scenario)
from inventory_optimizer.aggregate import is_aggregate_workbook  # noqa: E402
from inventory_optimizer.sample import _book, sample_data  # noqa: E402

st.set_page_config(page_title="Inventarios Mo", layout="wide")
st.title("Rotación y stock óptimo · productos de molibdeno")


def _xlsx(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data
def template_bytes() -> bytes:
    return _xlsx(_book(None))


@st.cache_data
def sample_bytes() -> bytes:
    return _xlsx(_book(sample_data()))


with st.sidebar:
    st.header("Datos")
    upload = st.file_uploader("Excel de entrada", type=["xlsx", "xlsm"])
    use_sample = upload is None and st.toggle("Sin archivo: usar caso de ejemplo", value=True)
    st.download_button("Descargar plantilla vacía", template_bytes(), "plantilla_inventarios_mo.xlsx")
    st.download_button("Descargar caso de ejemplo", sample_bytes(), "ejemplo_inventarios_mo.xlsx")

source = upload.getvalue() if upload is not None else sample_bytes() if use_sample else None
if source is None:
    st.info("Sube tu Excel (o activa el caso de ejemplo). La plantilla explica cada hoja y columna; "
            "solo la hoja Demanda es indispensable.")
    st.stop()



def aggregate_page(raw: bytes) -> None:
    from inventory_optimizer import aggregate as ag

    data = ag.load_aggregate(io.BytesIO(raw))
    p = data.params
    with st.sidebar:
        st.header("Supuestos")
        p.precio_usd_lb = st.number_input("Valor promedio (US$/lb)", 1.0, 200.0, float(p.precio_usd_lb), 0.5)
        p.tasa_costo_capital = st.number_input("Tasa deuda incremental (% anual)", 0.0, 30.0,
                                               float(p.tasa_costo_capital * 100), 0.25) / 100
        p.nivel_servicio = st.slider("Nivel de servicio", 0.80, 0.995, float(p.nivel_servicio), 0.005)
        p.lt_reposicion_pt_dias = st.number_input("Lead time reposición PT (días)", 1.0, 180.0,
                                                  float(p.lt_reposicion_pt_dias), 1.0)
        p.desv_lt_pt_dias = st.number_input("Desviación LT PT (días)", 0.0, 60.0, float(p.desv_lt_pt_dias), 1.0)
        p.segmentos_independientes = st.number_input("Segmentos independientes (mix)", 1, 100,
                                                     int(p.segmentos_independientes), 1)
        st.header("Días necesarios por etapa")
        st.caption("0 = sin dato. Con todos los días de un grupo se calcula su cobertura técnica.")
        for stg in p.stages:
            val = st.number_input(stg.name, 0.0, 365.0, float(stg.days) if stg.known else 0.0, 1.0, help=stg.help)
            if val > 0 and (not stg.known or val != stg.days):
                stg.days, stg.source = val, "ingresado en la app"
            elif val == 0:
                stg.source = "sin dato"
        bod = st.number_input("Días de venta en bodegas destino", 0.0, 120.0, 0.0, 1.0)
        if bod > 0:
            p.dias_bodega_destino, p.sources["dias_bodega_destino"] = bod, "ingresado en la app"
        p.desv_lt_mp_dias = st.number_input("Desviación LT materia prima (días)", 0.0, 60.0,
                                            float(p.desv_lt_mp_dias), 1.0)

    r = ag.analyze_aggregate(data)
    st.caption(f"{data.title} · {data.months[0].label} a {data.months[-1].label} · vista agregada (MMlb de Mo)")
    k = {row[0]: row for row in ag.kpis(r)}
    c = st.columns(5)
    c[0].metric("Ciclos/mes actual", fmt(k["Ciclos de rotación por mes"][1]))
    c[1].metric("Ciclos/mes demostrado", fmt(k["Ciclos de rotación por mes"][2]))
    c[2].metric("Días de cobertura", fmt(k["Días de cobertura"][1]),
                f"{k['Días de cobertura'][2] - k['Días de cobertura'][1]:+.0f} al demostrado", delta_color="inverse")
    c[3].metric("1 día de cobertura", f"US$ {k['Valor de 1 día de cobertura (US$mm)'][1]:,.1f} mm")
    c[4].metric("Capital liberable (demostrado)", f"US$ {k['Capital liberable vs actual (US$mm)'][2]:,.0f} mm")

    buf = io.BytesIO()
    ag.write_aggregate_report(r, buf)
    st.download_button("Descargar reporte Excel", buf.getvalue(), "reporte_rotacion_agregado.xlsx", type="primary")

    tabs = st.tabs(["Resumen", "Recomendaciones", "Rotación mensual", "Etapas", "Sensibilidad SS",
                    f"Advertencias ({len(data.warnings)})"])
    with tabs[0]:
        st.dataframe(pd.DataFrame(ag.kpis(r), columns=["Indicador", "Actual", "Demostrado (P25)", "Técnico"]),
                     hide_index=True, width="stretch")
        st.subheader("Cobertura por grupo de stock")
        st.dataframe(df(ag.group_table(r)), hide_index=True, width="stretch")
        st.caption("Demostrado = percentil 25 de los meses observados. Técnico = días por etapa x venta diaria + "
                   "stock de seguridad (requiere días por etapa).")
    with tabs[1]:
        for rec in r.recs:
            icon = {"Alta": "🔴", "Media": "🟠", "Baja": "🔵"}[rec[0]]
            extra = f" · US$ {rec[5]:,.0f} mm" if rec[5] else ""
            with st.expander(f"{icon} {rec[1]}{extra}"):
                st.write(f"**Diagnóstico:** {rec[2]}")
                st.write(f"**Acción:** {rec[3]}")
    with tabs[2]:
        mt = df(ag.month_table(r))
        st.dataframe(mt, hide_index=True, width="stretch")
        series = ["Stock total MMlb", "PT MMlb", "Pre-PT MMlb", "Ventas MMlb"]
        if mt["Producción MMlb"].notna().any():
            series.append("Producción MMlb")
        st.line_chart(mt.set_index("Mes")[series])
    with tabs[3]:
        st.dataframe(df(ag.stage_table(r)), hide_index=True, width="stretch")
    with tabs[4]:
        st.dataframe(df(ag.sensitivity_table(r)), hide_index=True, width="stretch")
        st.caption("Stock de seguridad en días de venta según nivel de servicio, lead time y mix.")
        if r.sku_rows:
            st.subheader("Cliente / SKU")
            st.dataframe(df(ag.sku_table(r)), hide_index=True, width="stretch")
    with tabs[5]:
        for w in data.warnings:
            st.warning(w)


def fmt(v, kind="n"):
    if v is None:
        return "s/d"
    if kind == "usd":
        return f"{v / 1e6:,.1f} M USD"
    return f"{v:,.2f}"


def df(tb: dict) -> pd.DataFrame:
    return pd.DataFrame(tb["rows"], columns=tb["headers"])


if is_aggregate_workbook(io.BytesIO(source)):
    aggregate_page(source)
    st.stop()

ds = load_excel(io.BytesIO(source))
p0 = ds.params
with st.sidebar:
    st.header("Supuestos")
    ds.params = replace(
        p0,
        nivel_servicio_a=st.slider("Nivel de servicio clientes A", 0.80, 0.999, float(p0.nivel_servicio_a), 0.005),
        nivel_servicio_b=st.slider("Nivel de servicio clientes B", 0.80, 0.999, float(p0.nivel_servicio_b), 0.005),
        nivel_servicio_c=st.slider("Nivel de servicio clientes C", 0.75, 0.999, float(p0.nivel_servicio_c), 0.005),
        tasa_costo_capital=st.number_input("Costo de capital anual (%)", 0.0, 50.0,
                                           float(p0.tasa_costo_capital * 100), 0.5) / 100,
        precio_mo_usd_kg=st.number_input("Precio Mo (USD/kg Mo)", 1.0, 500.0, float(p0.precio_mo_usd_kg), 1.0),
        ahorro_minimo_reasignacion_usd=st.number_input("Ahorro mínimo para reasignar (USD/año)", 0.0, 1e8,
                                                       float(p0.ahorro_minimo_reasignacion_usd), 5000.0),
        propiedad_en_transito=st.toggle("El stock en tránsito es de la empresa (CIF/DAP)", p0.propiedad_en_transito),
    )

if not ds.demand:
    st.error("No se encontraron datos de demanda válidos.")
    for w in ds.warnings:
        st.warning(w)
    st.stop()

alloc = optimize(analyze(ds))
recs = build_recommendations(alloc)
a = alloc.base
t = a.totals()


c = st.columns(5)
c[0].metric("Ciclos/mes actual", fmt(t["ciclos_mes_actual"]))
c[1].metric("Ciclos/mes óptimo", fmt(t["ciclos_mes_optimo"]),
            None if t["ciclos_mes_actual"] is None else f"{t['ciclos_mes_optimo'] - t['ciclos_mes_actual']:+.2f}")
c[2].metric("Días de inventario", fmt(t["doi_actual"]), None if t["doi_actual"] is None else
            f"{t['doi_optimo'] - t['doi_actual']:+.0f} días al óptimo", delta_color="inverse")
c[3].metric("Capital liberable", fmt(t["capital_liberable_usd"], "usd"))
c[4].metric("Ahorro anual", fmt(t["ahorro_anual_usd"], "usd"))
st.caption(f"Historia {a.months[0]} a {a.months[-1]} · {len(a.lines)} líneas cliente-producto · "
           f"{len(a.plants)} plantas · {len(alloc.moves)} reasignaciones propuestas")

buf = io.BytesIO()
write_report(alloc, recs, buf)
st.download_button("Descargar reporte Excel completo", buf.getvalue(), "reporte_inventarios_mo.xlsx",
                   type="primary")


tables = detail_tables(alloc, recs)
tabs = st.tabs(["Resumen", "Recomendaciones", "Rotación", "Stock óptimo", "Concentrado", "Clientes",
                "Reasignación", f"Advertencias ({len(ds.warnings)})"])

with tabs[0]:
    st.dataframe(pd.DataFrame(kpi_rows(a), columns=["Indicador", "Actual", "Óptimo", "Diferencia"]),
                 hide_index=True, width="stretch")
    plants = pd.DataFrame(plant_summary(a), columns=PLANT_HEADERS)
    st.subheader("Por planta")
    st.dataframe(plants, hide_index=True, width="stretch")
    st.bar_chart(plants.set_index("Planta")[["Ciclos/mes actual", "Ciclos/mes óptimo"]], stack=False)
    st.subheader("Escenario: asignación actual vs propuesta")
    base, prop = scenario(alloc.base), scenario(alloc.proposed)
    st.dataframe(pd.DataFrame([[k, base[k], prop[k]] for k in base],
                              columns=["Indicador", "Actual", "Propuesta"]), hide_index=True, width="stretch")

with tabs[1]:
    for r in recs:
        icon = {"Alta": "🔴", "Media": "🟠", "Baja": "🔵"}[r.prioridad]
        extra = f" · ahorro {r.ahorro_anual_usd:,.0f} USD/año" if r.ahorro_anual_usd else ""
        with st.expander(f"{icon} {r.tipo} — {r.ambito}{extra}"):
            st.write(f"**Diagnóstico:** {r.diagnostico}")
            st.write(f"**Acción:** {r.accion}")

for i, name in enumerate(["Rotacion", "Stock_Optimo", "Concentrado", "Clientes", "Reasignacion"], start=2):
    with tabs[i]:
        st.dataframe(df(tables[name]), hide_index=True, width="stretch")
        if name == "Reasignacion":
            st.caption("Carga por planta antes y después de la reasignación")
            st.dataframe(df(load_table(alloc)), hide_index=True, width="stretch")

with tabs[7]:
    if not ds.warnings:
        st.success("Sin advertencias: todos los datos necesarios vinieron en el Excel.")
    for w in ds.warnings:
        st.warning(w)
