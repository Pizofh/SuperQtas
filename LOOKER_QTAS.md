# SuperQTAS x Looker Studio

This guide uses the canonical operational sheets exported by SuperQTAS. Connect each sheet as its own Google Sheets data source; do not create spreadsheet formulas or duplicate data for reporting.

## Reporting Rules

- A sale is revenue. A payment is cash collection. Never add `Ventas` and `Pagos` to calculate revenue.
- A completed costed sale uses `Venta_Detalle_Costos_Calc`; only rows with `Estado_Costo = Completo` belong in the gross-margin calculation.
- A purchase is not automatically cost of goods sold. Purchase lines update historical costs; the resulting cost reaches the margin through `Costo_Total_Estimado` when the product is sold.
- `Compra_Origenes_Fondos` records who funded a purchase. It is useful for contributor/reimbursement reporting, not as a company expense total.
- `Distribucion_Ingresos` has one row for the sale and one for every payment. Filter `Fuente_Tipo` to either `Venta` or `Pago`; never aggregate both in a single total.

## Dashboard 1: Sales And Collections

Sources: `Ventas`, `Pagos`, and `Venta_Detalle`.

Use `Ventas.Fecha_Venta` for revenue dates and `Pagos.Fecha_Pago` for collection dates.

Recommended scorecards:

- Net sales: `SUM(Total_Venta)` from `Ventas`.
- Cash collected: `SUM(Monto_Pago)` from `Pagos`.
- Open balance: `SUM(Saldo)` from `Ventas`.
- Number of sales: `COUNT_DISTINCT(Venta_ID)` from `Ventas`.
- Average ticket: `SUM(Total_Venta) / COUNT_DISTINCT(Venta_ID)` from `Ventas`.

Recommended charts:

- Time series for sales and collections in separate charts.
- Sales by client and payment status.
- Product quantity, discounts, and net sales from `Venta_Detalle`.

## Dashboard 2: Gross Margin

Source: `Venta_Detalle_Costos_Calc`.

Apply this chart-level filter before calculating margin:

```text
Estado_Costo = Completo
```

Metrics:

```text
Net revenue          = SUM(Subtotal_Neto)
Estimated COGS       = SUM(Costo_Total_Estimado)
Gross profit         = SUM(Margen_Bruto_Estimado)
Gross margin percent = SUM(Margen_Bruto_Estimado) / SUM(Subtotal_Neto)
```

Dimensions: `Fecha_Venta`, `Producto_Estandar`, `Cliente_ID`, `Metodo_Costo`, and `Unidad`.

Use `Estado_Costo` as a quality-control filter. Rows marked `Sin receta`, `Sin costo`, or `No costeable` should be reported separately instead of silently treated as zero cost.

For a simple cash-oriented operating view, show `Compra_Detalle` lines where `Tipo_Item = Gasto` in a separate expense card. Do not subtract all purchases from revenue, because material purchases are already recognized through product cost when sold.

## Dashboard 3: Purchases And Funding

Sources: `Compras`, `Compra_Detalle`, and `Compra_Origenes_Fondos`.

Recommended metrics:

- Purchases: `SUM(Total_Compra)` from `Compras`.
- Purchase lines: `SUM(Costo_Total_Linea)` from `Compra_Detalle`.
- Funding by contributor: `SUM(Monto_Asignado)` from `Compra_Origenes_Fondos`.
- Purchase count: `COUNT_DISTINCT(Compra_ID)` from `Compras`.

Dimensions: `Fecha_Compra`, `Proveedor`, `Tipo_Item`, `Item`, `Origen_Fondos`, and `Aportante`.

Use this dashboard to reconcile a purchase against its sources of funds and to track reimbursements. It is a cash/funding dashboard, not a profit-and-loss calculation.

## Dashboard 4: Inventory And Production

Sources: `Inventario_Snapshot`, `Inventario_Movimientos`, `Producciones`, and `Produccion_Detalle`.

Recommended scorecards:

- Current stock: `SUM(Stock_Actual)` from `Inventario_Snapshot`.
- Items below minimum: count records where `Estado_Stock = Bajo`.
- Items out of stock: count records where `Estado_Stock = Agotado`.
- Produced units: `SUM(Cantidad_Producida)` from `Producciones`.

Use `Inventario_Snapshot` for the latest state and `Inventario_Movimientos` for history. Do not add both as if they were independent stock totals.

## Dashboard 5: Distribution

Source: `Distribucion_Ingresos`.

Choose one perspective at a time:

- Accrued sale distribution: filter `Fuente_Tipo = Venta` and use `Fecha_Venta`.
- Collected-cash distribution: filter `Fuente_Tipo = Pago` and use `Fecha_Pago`.

Metrics: `SUM(Steve_Valor)`, `SUM(Majo_Valor)`, `SUM(Mush_Valor)`, and `SUM(Monto_Base)`.

Dimensions: `Fecha_Base`, `Regla_ID_Usada`, `Medio_Pago`, and `Estado_Pago`.

## Source Classification

Operator-facing sheets:

- `Ventas`, `Pagos`, `Compras`, `Ventas_Envio`, `Producciones`, and `Inventario_Snapshot`.

Business-intelligence sheets:

- `Venta_Detalle`, `Venta_Detalle_Costos_Calc`, `Compra_Detalle`, `Compra_Origenes_Fondos`, `Distribucion_Ingresos`, `Inventario_Movimientos`, and `Produccion_Detalle`.

Backend/configuration sheets:

- `Productos`, `Precios_Referencia`, `Clientes`, `Costos_Referencia`, `Producto_Componentes`, `Producto_Reglas_Costo`, `Costo_Producto_Calc`, `Inventario_Control`, `Distribucion_Reglas`, `Origenes_Fondos_Reglas`, and `Config_MediosPago`.
