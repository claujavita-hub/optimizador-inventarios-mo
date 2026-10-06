"""Estructuras de datos del optimizador de inventarios de productos de Mo.

Unidades: toneladas de producto (t), días, USD. Las cantidades de demanda y
stock se expresan en toneladas del producto terminado (no Mo contenido);
el contenido de Mo de cada producto se usa solo para valorizar.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Params:
    nivel_servicio: float = 0.95          # por defecto si el cliente no tiene segmento
    nivel_servicio_a: float = 0.98
    nivel_servicio_b: float = 0.95
    nivel_servicio_c: float = 0.90
    nivel_servicio_concentrado: float = 0.95
    dias_mes: float = 30.0
    tasa_costo_capital: float = 0.10       # anual, sobre el valor del inventario
    precio_mo_usd_kg: float = 50.0         # para valorizar si el producto no trae valor
    contenido_mo_concentrado: float = 0.50
    propiedad_en_transito: bool = True     # True si el stock en tránsito es de la empresa (CIF/DAP)
    pedido_firme_dias: float = 0.0         # días de aviso firme del cliente (reduce exposición)
    # Defaults cuando falta información
    tiempo_proceso_dias: float = 7.0
    desv_proceso_dias: float = 1.0
    tiempo_qa_dias: float = 3.0
    intervalo_campana_dias: float = 30.0
    intervalo_despacho_dias: float = 30.0
    dias_transito: float = 30.0
    desv_transito_dias: float = 5.0
    lead_time_concentrado_dias: float = 45.0
    desv_lead_time_concentrado_dias: float = 7.0
    intervalo_recepcion_concentrado_dias: float = 30.0
    # Umbrales de diagnóstico / propuestas
    umbral_sobrestock: float = 1.25        # actual > máximo*umbral => sobrestock
    ahorro_minimo_reasignacion_usd: float = 25_000.0
    utilizacion_alta: float = 0.90
    utilizacion_baja: float = 0.50


@dataclass
class Plant:
    name: str
    country: str = ""
    region: str = ""
    capacity_t_month: float | None = None
    conc_lead_time_days: float | None = None
    conc_lead_time_sd_days: float | None = None
    conc_receipt_interval_days: float | None = None
    conc_stock_t: float | None = None
    conc_cost_usd_t: float | None = None   # costo logístico del concentrado puesto en planta (USD/t conc.)


@dataclass
class Product:
    name: str
    mo_content: float | None = None        # fracción 0-1
    conc_factor: float | None = None       # t concentrado por t producto
    value_usd_t: float | None = None


@dataclass
class Process:
    plant: str
    product: str
    process_days: float | None = None
    process_sd_days: float | None = None
    qa_days: float | None = None
    campaign_interval_days: float | None = None
    capacity_t_month: float | None = None
    cost_usd_t: float | None = None


@dataclass
class Customer:
    name: str
    country: str = ""
    region: str = ""
    plant: str = ""
    segment: str = ""
    service_level: float | None = None
    dispatch_interval_days: float | None = None
    firm_order_days: float | None = None


@dataclass
class Lane:
    plant: str
    destination: str
    transit_days: float
    transit_sd_days: float | None = None
    freight_usd_t: float | None = None


@dataclass
class StockRecord:
    plant: str
    product: str
    location: str      # "pt" | "transito" | "proceso" | "consignacion" | "concentrado"
    qty_t: float
    month: str = ""


@dataclass
class Dataset:
    params: Params = field(default_factory=Params)
    plants: dict[str, Plant] = field(default_factory=dict)
    products: dict[str, Product] = field(default_factory=dict)
    processes: dict[tuple[str, str], Process] = field(default_factory=dict)
    customers: dict[str, Customer] = field(default_factory=dict)
    lanes: list[Lane] = field(default_factory=list)
    # (cliente, producto) -> {mes "YYYY-MM": t}
    demand: dict[tuple[str, str], dict[str, float]] = field(default_factory=dict)
    # (cliente, producto) -> planta (override de la asignación del cliente)
    line_plant: dict[tuple[str, str], str] = field(default_factory=dict)
    stock: list[StockRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
