"""Definición de la plantilla Excel de entrada (hojas, columnas, sinónimos).

La usa tanto el cargador (para reconocer columnas aunque vengan con otro
nombre) como el generador de plantilla. Los nombres se comparan
normalizados: minúsculas, sin acentos, y cualquier separador -> "_".
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass


def norm(text) -> str:
    s = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


@dataclass(frozen=True)
class Col:
    key: str
    aliases: tuple[str, ...] = ()
    required: bool = False
    help: str = ""


@dataclass(frozen=True)
class Sheet:
    name: str
    aliases: tuple[str, ...]
    cols: tuple[Col, ...]
    help: str


SHEETS: tuple[Sheet, ...] = (
    Sheet("Plantas", ("plantas", "plants", "planta", "sites", "plant"), (
        Col("planta", ("plant", "site", "nombre", "nombre_planta"), True, "Nombre de la planta"),
        Col("pais", ("country",), help="País"),
        Col("region", ("continente", "zona", "region_planta"), help="Región / continente"),
        Col("capacidad_t_mes", ("capacidad", "capacity", "capacidad_mensual_t", "capacidad_t"),
            help="Capacidad total de conversión (t producto / mes)"),
        Col("lead_time_concentrado_dias", ("lt_concentrado", "lead_time_mp", "lead_time_concentrado"),
            help="Días desde la orden hasta la llegada del concentrado"),
        Col("desv_lead_time_concentrado_dias", ("desv_lt_concentrado", "desv_lead_time_mp"),
            help="Desviación estándar de ese lead time (días)"),
        Col("intervalo_recepcion_concentrado_dias", ("frecuencia_recepcion_concentrado", "intervalo_compra_dias"),
            help="Cada cuántos días llega un embarque de concentrado"),
        Col("stock_concentrado_t", ("stock_concentrado", "inventario_concentrado_t", "stock_mp_t"),
            help="Stock actual de concentrado (t)"),
        Col("costo_concentrado_usd_t", ("flete_concentrado_usd_t", "costo_logistico_concentrado"),
            help="Costo logístico de llevar el concentrado a la planta (USD/t concentrado)"),
    ), "Una fila por planta de conversión."),
    Sheet("Productos", ("productos", "products", "producto", "sku"), (
        Col("producto", ("product", "sku", "nombre"), True, "Ej: Óxido técnico, FeMo, Molibdato de sodio"),
        Col("contenido_mo_pct", ("contenido_mo", "mo_pct", "ley_mo", "pct_mo"), help="% de Mo contenido (ej. 57)"),
        Col("factor_concentrado", ("t_concentrado_por_t", "consumo_concentrado", "factor_mp"),
            help="t de concentrado consumidas por t de producto"),
        Col("valor_usd_t", ("precio_usd_t", "valor", "costo_usd_t", "valor_unitario"),
            help="Valor de inventario (USD/t). Si falta: precio Mo x contenido"),
    ), "Una fila por producto terminado."),
    Sheet("Proceso", ("proceso", "produccion", "process", "conversion", "procesos"), (
        Col("planta", ("plant", "site"), True),
        Col("producto", ("product", "sku"), True),
        Col("tiempo_proceso_dias", ("tiempo_proceso", "lead_time_produccion", "ciclo_produccion_dias", "process_days"),
            help="Días de conversión (tostación / reducción / química)"),
        Col("desv_proceso_dias", ("desv_tiempo_proceso", "desv_proceso"), help="Desviación estándar (días)"),
        Col("tiempo_qa_dias", ("tiempo_qa", "qa_dias", "liberacion_dias", "laboratorio_dias"),
            help="Días de muestreo, análisis y liberación"),
        Col("intervalo_campana_dias", ("frecuencia_campana_dias", "ciclo_campana", "campana_dias", "frecuencia_produccion_dias"),
            help="Cada cuántos días se produce una campaña de este producto"),
        Col("capacidad_t_mes", ("capacidad", "capacity"), help="Capacidad específica para este producto (t/mes)"),
        Col("costo_conversion_usd_t", ("costo_usd_t", "costo_produccion_usd_t", "costo"),
            help="Costo de conversión (USD/t), para comparar plantas"),
    ), "Una fila por combinación planta-producto que la planta puede fabricar."),
    Sheet("Clientes", ("clientes", "customers", "cliente"), (
        Col("cliente", ("customer", "nombre", "nombre_cliente"), True),
        Col("pais", ("country",)),
        Col("region", ("continente", "zona", "mercado")),
        Col("planta_asignada", ("planta", "plant", "planta_origen", "origen")),
        Col("segmento", ("clase", "abc", "segment"), help="A / B / C (define nivel de servicio)"),
        Col("nivel_servicio", ("service_level", "ns"), help="Override de nivel de servicio (ej. 97%)"),
        Col("intervalo_despacho_dias", ("frecuencia_despacho_dias", "frecuencia_despacho", "intervalo_despacho"),
            help="Cada cuántos días se despacha al cliente"),
        Col("pedido_firme_dias", ("aviso_firme_dias", "horizonte_pedido_firme", "anticipacion_pedido_dias"),
            help="Con cuántos días de anticipación el cliente confirma su pedido"),
    ), "Una fila por cliente."),
    Sheet("Demanda", ("demanda", "demand", "ventas", "despachos", "historial"), (
        Col("cliente", ("customer",), True),
        Col("producto", ("product", "sku"), True),
        Col("mes", ("periodo", "fecha", "month", "date"), True, "Mes (fecha o AAAA-MM)"),
        Col("cantidad_t", ("cantidad", "toneladas", "volumen", "qty", "t", "demanda_t", "ventas_t"), True),
        Col("planta", ("plant", "planta_origen"), help="Opcional: planta que abastece esta línea"),
    ), "Historia mensual de demanda/despachos. También se acepta formato ancho "
       "(cliente, producto y una columna por mes)."),
    Sheet("Transito", ("transito", "transit", "rutas", "lanes", "fletes", "logistica"), (
        Col("planta", ("plant", "origen"), True),
        Col("destino", ("cliente", "pais", "region", "destination"), True,
            "Nombre de cliente, país o región destino"),
        Col("dias_transito", ("transit_time", "tiempo_transito", "transito_dias", "dias"), True),
        Col("desv_transito_dias", ("desv_transito", "variabilidad_transito"), help="Desviación estándar (días)"),
        Col("flete_usd_t", ("flete", "costo_flete", "freight", "costo_flete_usd_t")),
    ), "Tiempos de tránsito planta -> destino (cliente, país o región)."),
    Sheet("Stock", ("stock", "inventario", "inventarios", "inventory", "existencias"), (
        Col("planta", ("plant", "site", "bodega"), True),
        Col("producto", ("product", "sku"), True),
        Col("ubicacion", ("estado", "tipo", "location", "tipo_stock"),
            help="Planta (PT), Tránsito, Proceso, Consignación o Concentrado"),
        Col("cantidad_t", ("cantidad", "toneladas", "stock_t", "qty"), True),
        Col("mes", ("fecha", "periodo", "month"), help="Opcional: si hay varios meses se promedia"),
    ), "Stock actual (o histórico mensual) por planta y producto."),
)

PARAM_HELP = {
    "nivel_servicio": "Nivel de servicio por defecto (clientes sin segmento)",
    "nivel_servicio_a": "Nivel de servicio clientes A",
    "nivel_servicio_b": "Nivel de servicio clientes B",
    "nivel_servicio_c": "Nivel de servicio clientes C",
    "nivel_servicio_concentrado": "Nivel de servicio del abastecimiento de concentrado",
    "dias_mes": "Días por mes",
    "tasa_costo_capital": "Costo anual de mantener inventario (% del valor)",
    "precio_mo_usd_kg": "Precio Mo (USD/kg Mo) para valorizar",
    "contenido_mo_concentrado": "Ley de Mo del concentrado",
    "propiedad_en_transito": "1 si el stock en tránsito es de la empresa (CIF/DAP), 0 si no (FOB/FCA)",
    "pedido_firme_dias": "Días de aviso firme por defecto de los clientes",
    "tiempo_proceso_dias": "Default: días de conversión",
    "desv_proceso_dias": "Default: desviación del tiempo de conversión",
    "tiempo_qa_dias": "Default: días de QA/liberación",
    "intervalo_campana_dias": "Default: días entre campañas",
    "intervalo_despacho_dias": "Default: días entre despachos a cliente",
    "dias_transito": "Default: días de tránsito",
    "desv_transito_dias": "Default: desviación del tránsito",
    "lead_time_concentrado_dias": "Default: lead time del concentrado",
    "desv_lead_time_concentrado_dias": "Default: desviación lead time concentrado",
    "intervalo_recepcion_concentrado_dias": "Default: días entre recepciones de concentrado",
    "umbral_sobrestock": "Stock > máximo x umbral = sobrestock",
    "ahorro_minimo_reasignacion_usd": "Ahorro anual mínimo para proponer cambiar de planta",
    "utilizacion_alta": "Utilización de capacidad considerada alta",
    "utilizacion_baja": "Utilización de capacidad considerada baja",
}

PARAMS_SHEET_ALIASES = ("parametros", "parameters", "supuestos", "config", "configuracion")
