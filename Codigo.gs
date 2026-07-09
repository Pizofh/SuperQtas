/****************************************************
 * QTAS ERP
 *
 * Configuracion compartida del mini ERP.
 ****************************************************/

var QTAS = {
  sheets: {
    productos: 'Productos',
    precios: 'Precios_Referencia',
    compras: 'Compras',
    compraDetalle: 'Compra_Detalle',
    costosReferencia: 'Costos_Referencia',
    productoComponentes: 'Producto_Componentes',
    productoReglasCosto: 'Producto_Reglas_Costo',
    costoProductoCalculado: 'Costo_Producto_Calc',
    ventaDetalleCostosCalculado: 'Venta_Detalle_Costos_Calc',
    origenesFondosReglas: 'Origenes_Fondos_Reglas',
    compraOrigenesFondos: 'Compra_Origenes_Fondos',
    clientes: 'Clientes',
    ventas: 'Ventas',
    detalle: 'Venta_Detalle',
    pagos: 'Pagos',
    ventasEnvio: 'Ventas_Envio',
    distribucionReglas: 'Distribucion_Reglas',
    distribucionIngresos: 'Distribucion_Ingresos',
    config: 'Config_MediosPago'
  },
  status: {
    pago: {
      pagado: 'Pagado',
      parcial: 'Parcial',
      pendiente: 'Pendiente',
      anulado: 'Anulado'
    },
    envio: {
      pendiente: 'Pendiente',
      enviado: 'Enviado'
    },
    registro: {
      activo: 'Activo',
      anulado: 'Anulado'
    }
  },
  schemas: {
    Productos: [
      'Producto_Estandar',
      'Unidad_Default',
      'Activo',
      'Nota'
    ],
    Precios_Referencia: [
      'Precio_ID',
      'Producto_Estandar',
      'Precio',
      'Unidad',
      'Fecha_Desde',
      'Fecha_Hasta',
      'Activo',
      'Nota'
    ],
    Compras: [
      'Compra_ID',
      'Fecha_Compra',
      'Proveedor',
      'Items_Resumen',
      'Total_Compra',
      'Medio_Pago',
      'Comentario_Compra',
      'Estado_Registro'
    ],
    Compra_Detalle: [
      'Compra_Detalle_ID',
      'Compra_ID',
      'Fecha_Compra',
      'Proveedor',
      'Tipo_Item',
      'Item',
      'Cantidad',
      'Unidad',
      'Costo_Total_Linea',
      'Costo_Unitario',
      'Impacta_Costo',
      'Comentario_Linea',
      'Estado_Registro'
    ],
    Costos_Referencia: [
      'Costo_ID',
      'Compra_ID',
      'Item',
      'Tipo_Item',
      'Unidad',
      'Costo_Unitario',
      'Proveedor',
      'Fecha_Desde',
      'Fecha_Hasta',
      'Activo',
      'Fuente_Tipo',
      'Fuente_ID',
      'Nota'
    ],
    Producto_Componentes: [
      'Componente_ID',
      'Producto_Estandar',
      'Unidad_Venta',
      'Orden',
      'Tipo_Componente',
      'Item_Componente',
      'Cantidad_Componente',
      'Unidad_Componente',
      'Merma_Pct',
      'Activo',
      'Nota'
    ],
    Producto_Reglas_Costo: [
      'Regla_Costo_ID',
      'Producto_Estandar',
      'Unidad_Venta',
      'Fecha_Desde',
      'Fecha_Hasta',
      'Cantidad_Min',
      'Cantidad_Max',
      'Orden',
      'Tipo_Componente',
      'Item_Componente',
      'Cantidad_Componente',
      'Unidad_Componente',
      'Aplicacion',
      'Merma_Pct',
      'Activo',
      'Nota'
    ],
    Costo_Producto_Calc: [
      'Costo_Producto_ID',
      'Fecha_Calculo',
      'Producto_Estandar',
      'Unidad_Venta',
      'Metodo_Costo',
      'Costo_Unitario_Total',
      'Costo_Unitario_Componentes',
      'Componentes_Activos',
      'Componentes_Con_Costo',
      'Componentes_Sin_Costo',
      'Cobertura_Costo_Pct',
      'Estado_Costo',
      'Nota'
    ],
    Venta_Detalle_Costos_Calc: [
      'Detalle_Costo_ID',
      'Detalle_ID',
      'Venta_ID',
      'Fecha_Venta',
      'Cliente_ID',
      'Nombre',
      'Producto_Estandar',
      'Cantidad',
      'Unidad',
      'Subtotal_Neto',
      'Costo_Unitario_Usado',
      'Costo_Total_Estimado',
      'Margen_Bruto_Estimado',
      'Margen_Porcentaje_Estimado',
      'Metodo_Costo',
      'Componentes_Con_Costo',
      'Componentes_Sin_Costo',
      'Cobertura_Costo_Pct',
      'Estado_Costo',
      'Actualizado_En',
      'Nota'
    ],
    Origenes_Fondos_Reglas: [
      'Regla_ID',
      'Origen_Fondos',
      'Fecha_Desde',
      'Fecha_Hasta',
      'Aportante',
      'Porcentaje',
      'Nota'
    ],
    Compra_Origenes_Fondos: [
      'Compra_Origen_ID',
      'Compra_ID',
      'Compra_Detalle_ID',
      'Fecha_Compra',
      'Origen_Fondos',
      'Aportante',
      'Porcentaje',
      'Monto_Asignado',
      'Fuente_Registro',
      'Nota'
    ],
    Clientes: [
      'Cliente_ID',
      'Nombre',
      'Activo',
      'Creado_En',
      'Actualizado_En'
    ],
    Ventas: [
      'Venta_ID',
      'Fecha_Venta',
      'Cliente_ID',
      'Nombre',
      'Productos_Resumen',
      'Comentario_Venta',
      'Total_Venta',
      'Total_Pagado',
      'Saldo',
      'Estado_Pago',
      'Regla_Distribucion_Venta_ID',
      'Steve_Pct_Venta',
      'Majo_Pct_Venta',
      'Mush_Pct_Venta',
      'Estado_Registro'
    ],
    Venta_Detalle: [
      'Detalle_ID',
      'Venta_ID',
      'Fecha_Venta',
      'Cliente_ID',
      'Nombre',
      'Producto_Estandar',
      'Cantidad',
      'Unidad',
      'Precio_Lista',
      'Precio_Vendido_Unitario',
      'Descuento_Linea',
      'Subtotal_Bruto',
      'Subtotal_Neto',
      'Comentario_Linea',
      'Estado_Registro'
    ],
    Pagos: [
      'Pago_ID',
      'Venta_ID',
      'Fecha_Pago',
      'Medio_Pago',
      'Monto_Pago',
      'Comentario_Pago',
      'Regla_Distribucion_Pago_ID',
      'Steve_Pct_Pago',
      'Majo_Pct_Pago',
      'Mush_Pct_Pago',
      'Estado_Registro'
    ],
    Ventas_Envio: [
      'Venta_ID',
      'Estado_Envio',
      'Fecha_Pendiente_Envio',
      'Fecha_Envio',
      'Comentario_Envio',
      'Creado_En',
      'Actualizado_En'
    ],
    Distribucion_Reglas: [
      'Regla_ID',
      'Fecha_Desde',
      'Fecha_Hasta',
      'Steve_Pct',
      'Majo_Pct',
      'Mush_Pct',
      'Activo',
      'Nota'
    ],
    Distribucion_Ingresos: [
      'Distribucion_ID',
      'Fuente_Tipo',
      'Fuente_ID',
      'Venta_ID',
      'Pago_ID',
      'Fecha_Base',
      'Fecha_Venta',
      'Fecha_Pago',
      'Nombre',
      'Cliente_ID',
      'Productos_Resumen',
      'Base_Distribucion',
      'Monto_Base',
      'Regla_ID_Usada',
      'Steve_Pct',
      'Majo_Pct',
      'Mush_Pct',
      'Steve_Valor',
      'Majo_Valor',
      'Mush_Valor',
      'Medio_Pago',
      'Estado_Pago',
      'Estado_Registro'
    ],
    Config_MediosPago: [
      'Medio_Pago',
      'Activo',
      'Nota'
    ]
  }
};

var PRECIOS_INICIALES = [
  { producto: 'AcAlt', precio: 18000, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'AcSup', precio: 20000, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'AcMed', precio: 12000, unidad: 'g', desde: '2024-01-01', hasta: '2025-02-28' },
  { producto: 'AcMed', precio: 14000, unidad: 'g', desde: '2025-03-01', hasta: '' },
  { producto: 'Lm', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'Cordy', precio: 2500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'Gano', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'ColaDP', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'Shii', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'LmPow', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'CordyPow', precio: 2500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'GanoPow', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'ColaDPPow', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: 'ShiiPow', precio: 1500, unidad: 'g', desde: '2024-01-01', hasta: '' },
  { producto: '50mg', precio: 1000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: '100mg', precio: 2000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: '150mg', precio: 2500, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: '200mg', precio: 3000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: '300mg', precio: 4000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: '500mg', precio: '', unidad: 'und', desde: '', hasta: '' },
  { producto: 'Choco', precio: 20000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'LmExt', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'CordyExt', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'GanoExt', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'ColaDPExt', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'ShiiExt', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'Tin', precio: 3000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'Vino', precio: 50000, unidad: 'und', desde: '2024-01-01', hasta: '' },
  { producto: 'Chocordy', precio: 35000, unidad: 'und', desde: '2024-01-01', hasta: '' }
];

var MEDIOS_PAGO = [
  'Efectivo',
  'NequiSteve',
  'NequiMajo',
  'Daviplata',
  'Bancolombia',
  'Otro'
];

var TIPOS_COMPRA_ITEM = [
  'Producto',
  'Insumo',
  'Gasto'
];

var DISTRIBUCION_REGLAS_INICIALES = [
  {
    desde: '2000-01-01',
    hasta: '2025-08-31',
    steve: 40,
    majo: 40,
    mush: 20,
    nota: 'Regla historica 40/40/20'
  },
  {
    desde: '2025-09-01',
    hasta: '',
    steve: 35,
    majo: 45,
    mush: 20,
    nota: 'Regla actual 35/45/20'
  }
];
