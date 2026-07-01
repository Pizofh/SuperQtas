# QTAS x Looker

Usa `Fecha_Venta`, `Fecha_Pago` y `Fecha_Base` como fuentes canonicas de tiempo.

## Fuentes recomendadas

`Ventas`
- Para ventas por dia, semana, mes, cliente, estado de pago y ticket promedio.
- Campo de tiempo recomendado: `Fecha_Venta`.

`Venta_Detalle`
- Para analizar productos, cantidades, descuentos, precio vendido y subtotal neto.
- Campo de tiempo recomendado: `Fecha_Venta`.

`Pagos`
- Para flujo de caja, medios de pago y cobros posteriores.
- Campo de tiempo recomendado: `Fecha_Pago`.

`Costo_Producto_Calc`
- Para costo unitario vigente por producto, cobertura de receta y metodo de costeo.
- Campo de tiempo recomendado: `Fecha_Calculo`.

`Venta_Detalle_Costos_Calc`
- Para margen bruto estimado por linea vendida, comparando precio de venta vs costo calculado.
- Campo de tiempo recomendado: `Fecha_Venta`.

`Compra_Origenes_Fondos`
- Para analizar de donde salieron los gastos/compras y como se repartieron por aportante.
- Campo de tiempo recomendado: `Fecha_Compra`.

`Distribucion_Ingresos`
- Para analizar reparto Steve / Majo / Mush.
- Campo de tiempo recomendado: `Fecha_Base`.
- Tambien puedes usar `Fecha_Venta` o `Fecha_Pago` si quieres separar venta vs cobro real.

## Campos que no deberian ser canonicos en Looker

No bases tus filtros principales en columnas auxiliares, backups o hojas historicas.

El analisis robusto debe salir de estos datetime canonicos:
- `Ventas.Fecha_Venta`
- `Venta_Detalle.Fecha_Venta`
- `Pagos.Fecha_Pago`
- `Distribucion_Ingresos.Fecha_Base`
- `Distribucion_Ingresos.Fecha_Venta`
- `Distribucion_Ingresos.Fecha_Pago`

## Uso sugerido por hoja

`Ventas`
- Dimensiones: fecha, cliente, estado de pago.
- Metricas: `SUM(Total_Venta)`, `SUM(Total_Pagado)`, `SUM(Saldo)`, `COUNT_DISTINCT(Venta_ID)`.

`Venta_Detalle`
- Dimensiones: fecha, producto, unidad, cliente.
- Metricas: `SUM(Cantidad)`, `SUM(Subtotal_Neto)`, `SUM(Descuento_Linea)`.

`Pagos`
- Dimensiones: fecha, medio de pago.
- Metricas: `SUM(Monto_Pago)`, `COUNT_DISTINCT(Pago_ID)`.

`Costo_Producto_Calc`
- Dimensiones: producto, unidad, metodo, estado.
- Metricas: `AVG(Costo_Unitario_Total)`, `AVG(Cobertura_Costo_Pct)`.

`Venta_Detalle_Costos_Calc`
- Dimensiones: fecha, cliente, producto, metodo, estado.
- Metricas: `SUM(Subtotal_Neto)`, `SUM(Costo_Total_Estimado)`, `SUM(Margen_Bruto_Estimado)`.

`Compra_Origenes_Fondos`
- Dimensiones: fecha, origen de fondos, aportante.
- Metricas: `SUM(Monto_Asignado)`, `COUNT_DISTINCT(Compra_Origen_ID)`.

`Distribucion_Ingresos`
- Dimensiones: fecha, fuente tipo, medio de pago.
- Metricas: `SUM(Steve_Valor)`, `SUM(Majo_Valor)`, `SUM(Mush_Valor)`, `SUM(Monto_Base)`.

## Hojas operativas vs analiticas

Operativas:
- `Productos`
- `Precios_Referencia`
- `Producto_Componentes`
- `Compras`
- `Compra_Detalle`
- `Costos_Referencia`
- `Clientes`
- `Ventas`
- `Venta_Detalle`
- `Pagos`
- `Ventas_Envio`
- `Config_MediosPago`

Configuracion / reglas:
- `Distribucion_Reglas`
- `Origenes_Fondos_Reglas`

Analiticas / salida:
- `Costo_Producto_Calc`
- `Venta_Detalle_Costos_Calc`
- `Compra_Origenes_Fondos`
- `Distribucion_Ingresos`

Historicas / soporte:
- `ARCHIVO__Data_Limpia`
- `Venta_Detalle_Costos`
- cualquier hoja `ARCHIVO__...`
- backups de precios

Legacy para migracion o respaldo:
- `Insumos`
- `Costo_Producto`
- `compras_resumen`
- `costos_referencia_sugeridos`
- `gastos_aportes_desglosado`
- `gastos_compras_limpio`
- `reglas_origen_fondos`

## Vista sugerida del libro

Deja visibles:
- `Productos`
- `Precios_Referencia`
- `Producto_Componentes`
- `Compras`
- `Compra_Detalle`
- `Costos_Referencia`
- `Clientes`
- `Ventas`
- `Venta_Detalle`
- `Pagos`
- `Costo_Producto_Calc`
- `Venta_Detalle_Costos_Calc`
- `Compra_Origenes_Fondos`
- `Distribucion_Reglas`
- `Origenes_Fondos_Reglas`
- `Config_MediosPago`
- `Distribucion_Ingresos`

Puedes ocultar sin problema operativo:
- `Insumos`
- `Costo_Producto`
- `Venta_Detalle_Costos`
- `compras_resumen`
- `costos_referencia_sugeridos`
- `gastos_aportes_desglosado`
- `gastos_compras_limpio`
- `reglas_origen_fondos`
- `Ventas_Envio` si no necesitas verla en el libro; su uso es operativo para seguimiento de envios, no analitico

Candidatas a borrar despues, solo cuando ya no las necesites como respaldo:
- `Config` legado, despues de migrar a `Config_MediosPago`
- `Ventas_Resumen`
- `Deudores`
- `ARCHIVO__INV`
- `ARCHIVO__Data_Limpia`
- `ARCHIVO__mapa_stock_estandar`
- `Precios_Referencia_Backup_20260`

## Recomendacion final

Para Looker, conecta minimo estas 4 hojas:
- `Ventas`
- `Venta_Detalle`
- `Pagos`
- `Distribucion_Ingresos`

Si quieres analisis real de costo vs precio de venta, suma tambien:
- `Producto_Componentes`
- `Costo_Producto_Calc`
- `Venta_Detalle_Costos_Calc`
- `Compras`
- `Compra_Detalle`
- `Costos_Referencia`

Si quieres simplicidad maxima, puedes empezar solo con:
- `Ventas`
- `Venta_Detalle`
- `Pagos`
