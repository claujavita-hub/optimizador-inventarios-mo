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
# Streamlit Cloud actualiza este archivo al hacer push pero mantiene en memoria los módulos ya importados:
# se descartan para que cada ejecución use el motor de cálculo vigente (evita mezclar versiones).
for _mod in [m for m in sys.modules if m == "inventory_optimizer" or m.startswith("inventory_optimizer.")]:
    del sys.modules[_mod]

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
    upload_tab = st.file_uploader("Tablero de solvencia y liquidez (opcional)", type=["xlsx", "xlsm"],
                                  help="Hojas Tablero y Umbrales. Con él la app calcula el impacto de cada acción "
                                       "de inventario en deuda neta/EBITDA, covenant, días de inventario y liquidez.")
    use_sample = upload is None and st.toggle("Sin archivo: usar caso de ejemplo", value=True)
    st.download_button("Descargar plantilla vacía", template_bytes(), "plantilla_inventarios_mo.xlsx")
    st.download_button("Descargar caso de ejemplo", sample_bytes(), "ejemplo_inventarios_mo.xlsx")

source = upload.getvalue() if upload is not None else sample_bytes() if use_sample else None
tablero_raw = upload_tab.getvalue() if upload_tab is not None else None
if source is None and tablero_raw is None:
    st.info("Sube tu Excel (o activa el caso de ejemplo). La plantilla explica cada hoja y columna; "
            "solo la hoja Demanda es indispensable.")
    st.stop()



def load_tab(*raws):
    from inventory_optimizer.solvency import load_tablero

    for raw in raws:
        if raw:
            try:
                tab = load_tablero(io.BytesIO(raw))
            except Exception as exc:   # archivo que no es un tablero
                st.sidebar.warning(f"No pude leer el tablero de solvencia: {exc}")
                continue
            if tab is not None and tab.latest:
                return tab
    return None


def solvency_view(tab, r=None) -> None:
    """Tablero de solvencia y, si hay análisis de inventario, el impacto de las acciones."""
    from inventory_optimizer import solvency as sv

    periods = [p for p, v in tab.periods.items() if v.get("ebitda") is not None and v.get("deuda") is not None]
    c = st.columns(3)
    period = c[0].selectbox("Período del tablero", periods, index=len(periods) - 1)
    price = r.data.params.precio_usd_lb if r is not None else 33.0
    value_lb = c[1].number_input("Valor del inventario (US$/lb)", 1.0, 200.0, float(price), 0.5,
                                 help="Por defecto el valor promedio por libra de 01_Supuestos. Si el inventario "
                                      "del balance está a costo, usa el valor contable por libra.")
    use = c[2].radio("Uso de la caja liberada", ["Mantener en caja", "Pagar deuda"], horizontal=True)
    use = "deuda" if use == "Pagar deuda" else "caja"

    actions: list = []
    if r is not None:
        st.subheader("Acciones de inventario")
        c = st.columns(3)
        t_tr = c[0].number_input("Días objetivo de tránsito por mercado", 10.0, 120.0, 45.0, 5.0)
        t_bod = c[1].number_input("Días objetivo en bodegas por mercado", 5.0, 90.0, 20.0, 5.0)
        t_asg = c[2].number_input("Días objetivo de asignado esperando despacho", 2.0, 30.0, 10.0, 1.0)
        acts = sv.inventory_actions(r, t_tr, t_bod, t_asg)
        if not acts:
            st.info("El análisis no encontró acciones cuantificables con los datos actuales.")
        ed = st.data_editor(pd.DataFrame([[a.default, a.code, a.name, a.mmlb, a.usd(value_lb), a.status, a.basis]
                                          for a in acts],
                                         columns=["Incluir", "Código", "Acción", "MMlb", "US$ mm", "Estado",
                                                  "Base del cálculo"]),
                            hide_index=True, width="stretch", disabled=["Código", "Acción", "US$ mm", "Estado",
                                                                        "Base del cálculo"],
                            column_config={"MMlb": st.column_config.NumberColumn(format="%.2f", min_value=0.0),
                                           "US$ mm": st.column_config.NumberColumn(format="%.1f")},
                            key=f"acts_{len(acts)}")
        for a, (_, row) in zip(acts, ed.iterrows()):
            if row["Incluir"]:
                a.mmlb = float(row["MMlb"])
                actions.append(a)
        st.caption("Marcadas por defecto: acciones con datos validados. Tránsito, asignado y spot quedan sin marcar "
                   "hasta validarlas. Puedes editar las MMlb de cada acción.")
    release = sv.actions_release(actions, value_lb)
    extra = st.number_input("Otra liberación de capital de trabajo (US$ mm)", 0.0, 2000.0, 0.0, 5.0,
                            help="Por ejemplo, cobranza o proveedores. Se suma a las acciones de inventario.")
    release += extra
    rate = r.data.params.tasa_costo_capital if r is not None else 0.08
    m = st.columns(4)
    base = tab.ind(period)
    sim = tab.ind(period, release, use, rate)
    m[0].metric("Liberación total", f"US$ {release:,.0f} mm")
    m[1].metric("Deuda neta / EBITDA", sv.fmt("nd_ebitda", sim["nd_ebitda"]),
                f"{sim['nd_ebitda'] - base['nd_ebitda']:+.2f}" if release else None, delta_color="inverse")
    m[2].metric("Holgura covenant", sv.fmt("holgura", sim["holgura"]),
                f"{(sim['holgura'] - base['holgura']) * 100:+.1f} pts" if release else None)
    m[3].metric("Ahorro de intereses", f"US$ {release * rate:,.1f} mm/año")

    tb = sv.scenario_table(tab, period, release, use, rate)
    rows = [[x[0], x[1], f"{sv.ICON[x[2]]} {x[2]}", x[3], f"{sv.ICON[x[4]]} {x[4]}"] for x in tb["rows"]]
    st.dataframe(pd.DataFrame(rows, columns=tb["headers"]), hide_index=True, width="stretch")
    st.caption(f"{tab.holgura_desc}. Los gastos financieros bajan en liberación x tasa "
               f"({rate:.1%}): intereses que se dejan de pagar o que gana la caja. El flujo libre suma la "
               "liberación una sola vez (caja de única vez, no flujo recurrente). Con 'Pagar deuda' baja también "
               "el pasivo; con 'Mantener en caja' mejoran la razón ácida y caja/deuda CP.")
    if tab.covenants:
        cov = sv.indicators_covenants(tab, period, release, use, rate)
        st.subheader("Covenants financieros")
        st.dataframe(pd.DataFrame(cov, columns=["Covenant", "Límite", f"{period} actual", "Con acciones",
                                                "Cumple"]), hide_index=True, width="stretch")
    need = sv.thresholds_needed(tab, period, use, rate)
    if need:
        st.subheader("¿Cuánto hay que liberar para llegar a verde?")
        st.dataframe(pd.DataFrame([[n[0], n[1], f"{sv.ICON[n[2]]} {n[2]}",
                                    f"US$ {n[3]:,.0f} mm" if n[3] is not None else sv._no_need(n)]
                                   for n in need], columns=["Indicador", "Actual", "Semáforo", "Liberación necesaria"]),
                     hide_index=True, width="stretch")
    buf = io.BytesIO()
    sv.write_solvency_report(tab, period, actions, value_lb, rate, use, buf)
    st.download_button("Descargar escenario de solvencia (Excel)", buf.getvalue(), "escenario_solvencia.xlsx")
    for w in tab.warnings:
        st.warning(w)


def aggregate_page(raw: bytes, tablero_raw: bytes | None = None) -> None:
    from inventory_optimizer import aggregate as ag

    data = ag.load_aggregate(io.BytesIO(raw))
    tab = load_tab(tablero_raw, raw)
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
        st.header("Bodegas destino")
        p.stock_spot = st.number_input("Stock para ventas spot (MMlb)", 0.0, 20.0, float(p.stock_spot), 0.1)
        p.buffer_stock = st.number_input("Buffer stock en bodegas (MMlb)", 0.0, 20.0, float(p.buffer_stock), 0.1)
        p.pct_venta_bodega = st.slider("% de la venta que pasa por bodega", 0.0, 1.0, float(p.pct_venta_bodega), 0.05)
        p.despacho_incluye_bodega = st.toggle("Despacho/entrega incluye la estadía en bodega destino",
                                              p.despacho_incluye_bodega)
        if not p.despacho_incluye_bodega:
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

    names = ["Resumen", "Recomendaciones", "Rotación mensual", "Etapas", "Sensibilidad SS",
             f"Advertencias ({len(data.warnings)})"]
    if data.celonis is not None:
        names.append("Detalle Celonis")
    if tab is not None:
        names.append("Solvencia")
    tabs = st.tabs(names)
    if data.celonis is not None:
        with tabs[6]:
            celonis_view(data.celonis, ag.latest_pt_month(data), data.sales_market, p.precio_usd_lb)
    if tab is not None:
        with tabs[-1]:
            solvency_view(tab, r)
    elif tablero_raw is None:
        st.sidebar.caption("Sube el tablero de solvencia para ver el impacto en deuda neta/EBITDA y covenant.")
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
        st.subheader("Balance de masa")
        st.dataframe(df(ag.balance_table(r)), hide_index=True, width="stretch")
        st.caption("Producción = venta propia + maquila + variación de stock. Una diferencia grande indica que algún "
                   "dato no corresponde (producción con intermedios, ventas fuera de Own Sales, saldo inicial).")
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


def celonis_view(c, month=None, sales=None, price: float = 33.0) -> None:
    from inventory_optimizer import celonis as cz

    st.caption(f"Foto por lote del producto terminado ({len(c.lots)} lotes, {c.total:.2f} MMlb). Desagrega el PT, "
               "no se suma al stock del dashboard.")
    st.subheader("Por mercado y etapa")
    st.dataframe(df(cz.stage_market_table(c)), hide_index=True, width="stretch")
    if c.has_targets:
        st.subheader("Stock vs metas de días de Celonis")
        st.dataframe(df(cz.target_table(c, price)), hide_index=True, width="stretch")
        st.caption("Exceso = stock x (1 − meta/días) en cada material que supera su meta. Los días los calcula "
                   "Celonis sobre el consumo de 12 meses. El 'Not Assigned' tiene días muy altos porque su base de "
                   "consumo es distinta: no se usa como acción.")
    if sales:
        st.subheader("Días de venta en tránsito y en bodega, por mercado")
        st.dataframe(df(cz.transit_days_table(c, sales)), hide_index=True, width="stretch")
    st.subheader("Conciliación con el dashboard")
    st.dataframe(df(cz.reconciliation_table(c, month)), hide_index=True, width="stretch")
    col = st.columns(2)
    with col[0]:
        st.subheader("Por material")
        st.dataframe(df(cz.material_table(c)), hide_index=True, width="stretch")
    with col[1]:
        st.subheader("Por planta")
        st.dataframe(df(cz.plant_table(c)), hide_index=True, width="stretch")
    st.subheader("Bodegas externas")
    st.dataframe(df(cz.warehouse_table(c)), hide_index=True, width="stretch")
    for rec in cz.insights(c, month, sales):
        with st.expander(f"{ {'Alta': '🔴', 'Media': '🟠', 'Baja': '🔵'}[rec[0]] } {rec[1]}"):
            st.write(f"**Diagnóstico:** {rec[2]}")
            st.write(f"**Acción:** {rec[3]}")
    for w in c.warnings:
        st.warning(w)


def fmt(v, kind="n"):
    if v is None:
        return "s/d"
    if kind == "usd":
        return f"{v / 1e6:,.1f} M USD"
    return f"{v:,.2f}"


def df(tb: dict) -> pd.DataFrame:
    return pd.DataFrame(tb["rows"], columns=tb["headers"])


from inventory_optimizer.celonis import is_celonis_only, load_celonis  # noqa: E402

_only_tab = upload is None and tablero_raw is not None          # solo se subió el tablero
_only_tab = _only_tab or (upload is not None and not is_aggregate_workbook(io.BytesIO(source))
                          and load_tab(source) is not None)    # el tablero se subió como archivo principal
if _only_tab:
    # solo el tablero: indicadores y semáforos sin acciones de inventario
    _tab = load_tab(tablero_raw, source)
    if _tab is None:
        st.error("No encontré las hojas Tablero/Umbrales con datos en el archivo.")
        st.stop()
    st.info("Solo tablero de solvencia. Sube también el modelo de inventario para calcular el impacto de cada acción.")
    solvency_view(_tab)
    st.stop()

if is_celonis_only(io.BytesIO(source)):
    st.info("Archivo de Celonis (detalle por lote). Para cruzarlo con la rotación, súbelo dentro del modelo "
            "(hoja 11_Celonis_Detalle) o sube el modelo completo.")
    celonis_view(load_celonis(io.BytesIO(source)))
    st.stop()

if is_aggregate_workbook(io.BytesIO(source)):
    aggregate_page(source, tablero_raw)
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
