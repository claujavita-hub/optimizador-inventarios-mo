# Optimizador de inventarios: conversión de productos de molibdeno

A partir de un Excel
con la red de plantas, clientes, tránsitos y procesos calcula:

1. **Rotación actual**: ciclos por mes y por año, días de inventario y
   estado (en rango / sobrestock / riesgo de quiebre / stock sin demanda)
   por planta y producto.
2. **Stock óptimo**: stock de seguridad, de ciclo, en proceso y en
   tránsito; niveles mínimo, objetivo y máximo de producto terminado, más
   el punto de reorden para lanzar cada campaña.
3. **Concentrado**: stock de seguridad y objetivo de materia prima por
   planta, según su lead time de abastecimiento.
4. **Reasignación de clientes entre plantas**: busca la planta que
   minimiza flete + conversión + logística del concentrado + costo de
   capital del stock, respetando la capacidad de cada planta.
5. **Recomendaciones priorizadas**, cada una con su impacto en toneladas,
   capital de trabajo y ahorro anual: reducir sobrestock, adelantar
   campañas, reasignar clientes, hacer campañas más frecuentes, acordar
   pedidos firmes, atacar la variabilidad del tránsito y usar la capacidad
   ociosa.

## Uso

App web (sirve desde el iPad si se publica en Streamlit Cloud con
*Main file* `inventory_optimizer/streamlit_app.py`):

```bash
pip install -r requirements.txt
streamlit run inventory_optimizer/streamlit_app.py
```

Línea de comandos:

```bash
python -m inventory_optimizer plantilla mi_plantilla.xlsx     # plantilla vacía con instrucciones
python -m inventory_optimizer ejemplo   ejemplo.xlsx          # caso ficticio de 4 plantas y 20 clientes
python -m inventory_optimizer analizar  mis_datos.xlsx -o reporte.xlsx
```

En `inventory_optimizer/plantillas/` están la plantilla vacía y el caso de ejemplo.

## Modo agregado (stock por etapa a nivel empresa)

Si el Excel trae la foto física mensual de toda la empresa, en lugar del
detalle por planta y cliente, el módulo cambia solo a **modo agregado**.
Lo reconoce porque tiene una hoja `*Inventario_Fisico` con stock total,
producto terminado, Pre-PT y Own Sales en MMlb, y opcionalmente
`*PT_Cobertura` (Not Assigned / Assigned / In-Transit / Warehouse /
Consignment), `*Supuestos`, `*Ciclo_Transito` y `*Cliente_SKU`. Este modo
entrega:

- **Rotación mensual:** ciclos por mes y días de cobertura en total, en PT y en Pre-PT, y días por categoría de PT.
- **Nivel demostrado:** el percentil 25 de los propios meses de cada grupo de stock. Es una meta que la operación ya
  alcanzó, por lo que no depende de supuestos.
- **Nivel técnico:** días necesarios por etapa × venta diaria, más el stock de seguridad. Solo se calcula cuando hay
  días reales por etapa, ya sea porque la hoja `Ciclo_Transito` está completa o porque se ingresan en la app.
- **Sensibilidad del stock de seguridad** frente al nivel de servicio, el lead time y el mix (segmentos independientes).
- **Recomendaciones:** tendencias como un Pre-PT que crece sin que crezca la venta, un stock que no sigue a la venta,
  el tránsito implícito, la cobertura de la venta del mes siguiente y el valor de un día de cobertura en US$.

## Excel de entrada (modo detallado)

| Hoja | Contenido | ¿Obligatoria? |
|---|---|---|
| Demanda | cliente, producto, mes, toneladas (también se acepta una columna por mes) | **Sí** |
| Clientes | país, región, planta asignada, segmento A/B/C, nivel de servicio, frecuencia de despacho, días de pedido firme | Recomendada |
| Transito | planta → destino (cliente, país o región): días, desviación, flete USD/t | Recomendada |
| Proceso | planta-producto: días de proceso y QA, intervalo de campaña, capacidad, costo de conversión | Recomendada |
| Stock | planta, producto, ubicación (planta / proceso / tránsito / consignación / concentrado), toneladas | Para medir rotación actual |
| Plantas | capacidad, lead time y stock de concentrado, costo logístico del concentrado | Opcional |
| Productos | % de Mo, t de concentrado por t de producto, valor USD/t | Opcional |
| Parametros | niveles de servicio, costo de capital, precio del Mo, valores por defecto | Opcional |

Los encabezados se reconocen aunque vengan con otro nombre, mayúsculas o
acentos (por ejemplo, "Toneladas" se lee como `cantidad_t` y "Origen" como
`planta`). Si falta algún dato se usa un supuesto por defecto, y el reporte
lo indica en la hoja **Supuestos**.

## Método (resumen)

- Exposición = campaña + proceso + QA + tránsito − aviso firme del cliente.
- Stock de seguridad = z · √(exposición · σd² + d² · (σproceso² + σtránsito²)), agregado por planta-producto con pooling (√Σ SS²).
- Stock de ciclo = d · max(campaña, despacho) / 2; en proceso = d · tiempo de proceso; en tránsito = d · días de tránsito.
- Ciclos por mes = demanda mensual / inventario total (producto terminado + proceso + tránsito).

El detalle completo está en la hoja **Metodologia** del reporte. Pruebas:
`python -m unittest discover -s inventory_optimizer/tests`.
