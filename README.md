# SuperQTAS

SuperQTAS es un ERP ligero construido sobre Google Sheets y Google Apps Script para operar ventas, compras, cartera, costos historicos y configuracion comercial desde una interfaz web integrada en la hoja.

No busca reemplazar un ERP corporativo; busca resolver bien la operacion diaria de un negocio pequeno con trazabilidad, reglas claras y mantenimiento simple.

## Que hace

- Registra ventas con varios productos, descuentos, pagos iniciales y saldo pendiente.
- Lleva cartera viva con ventas pendientes, pagos posteriores y deudores resumidos.
- Hace seguimiento de envios pendientes dentro del mismo flujo comercial.
- Registra compras por lineas con proveedor, medio de pago y comentario operativo.
- Actualiza costos de referencia historicos desde compras reales.
- Calcula costos y margenes para productos compuestos usando recetas y reglas de costo.
- Permite configurar productos, precios, medios de pago y reglas de distribucion sin editar codigo.
- Mantiene separados los ambientes de QA y produccion para probar antes de desplegar.

## Como funciona

La solucion esta dividida en cuatro capas:

1. Google Sheets como base operativa.
2. Apps Script como backend y modelo transaccional.
3. HTML/CSS/JavaScript como interfaz embebida dentro del spreadsheet.
4. `clasp`, Node.js y GitHub Actions para despliegue y automatizacion.

En la practica, el usuario abre la hoja, entra al menu `QTAS ERP`, trabaja desde la interfaz y los scripts actualizan las hojas canonicas del modelo por debajo.

## Modulos principales

### Ventas

- Registro de ventas con multiples lineas.
- Pagos completos o parciales desde el inicio.
- Registro posterior de pagos pendientes.
- Panel de deudores y ventas abiertas.
- Seguimiento de ventas marcadas para envio.

### Compras

- Registro de compras por proveedor y medio de pago.
- Catalogo sugerido de items para reducir digitacion manual.
- Estandarizacion de items conocidos al guardar.
- Opcion para decidir si una linea impacta o no el costo historico.

### Costos y analitica

- Costos de referencia vigentes e historicos.
- Productos compuestos con componentes y reglas variables.
- Calculo incremental de `Costo_Producto_Calc`.
- Calculo incremental de `Venta_Detalle_Costos_Calc`.
- Cobertura de costo, metodo usado y margen estimado por linea vendida.

### Configuracion operativa

- Catalogo de productos activos e inactivos.
- Historial de precios con vigencia por fecha.
- Medios de pago configurables.
- Reglas de distribucion por fecha.

## Flujo operativo

### Venta

1. Se seleccionan productos, cantidades, precios y descuentos.
2. Se registran pagos iniciales si existen.
3. Se actualizan `Ventas`, `Venta_Detalle` y `Pagos`.
4. Se sincroniza la distribucion de ingresos.
5. Se actualiza la analitica incremental de costos de esa venta.

### Compra

1. Se registra proveedor, medio de pago y lineas.
2. Cada linea puede impactar o no el costo de referencia.
3. Se actualizan `Compras`, `Compra_Detalle` y `Costos_Referencia`.
4. Se refresca el snapshot operativo de costos de producto.

### Configuracion

1. Se editan productos, precios, medios de pago o distribuciones desde la UI.
2. Los cambios se aplican sin romper historicos ya registrados.

## Hojas principales del modelo

Estas son las hojas operativas mas importantes:

- `Productos`
- `Precios_Referencia`
- `Clientes`
- `Ventas`
- `Venta_Detalle`
- `Pagos`
- `Ventas_Envio`
- `Compras`
- `Compra_Detalle`
- `Costos_Referencia`
- `Producto_Componentes`
- `Producto_Reglas_Costo`
- `Costo_Producto_Calc`
- `Venta_Detalle_Costos_Calc`
- `Distribucion_Reglas`
- `Distribucion_Ingresos`
- `Config_MediosPago`

## Arquitectura del repositorio

- [Codigo.gs](./Codigo.gs) define constantes, esquemas y datos base.
- [QTAS_Ventas.gs](./QTAS_Ventas.gs) contiene el flujo comercial de ventas, pagos y envios.
- [QTAS_Compras.gs](./QTAS_Compras.gs) maneja compras y costos historicos.
- [QTAS_CostosProducto.gs](./QTAS_CostosProducto.gs) resuelve recetas, costo compuesto y margen.
- [QTAS_Distribucion.gs](./QTAS_Distribucion.gs) calcula la distribucion de ingresos.
- [QTAS_Admin.gs](./QTAS_Admin.gs) expone la configuracion avanzada al front.
- [QTAS_Modelo.gs](./QTAS_Modelo.gs) asegura y valida la estructura canonica.
- [QTAS_Utils.gs](./QTAS_Utils.gs) concentra utilidades transversales.
- [App.html](./App.html) contiene la interfaz web embebida.

## Calidad, QA y despliegue

El repositorio ya trabaja con separacion real entre QA y produccion.

- `qa` despliega al proyecto de prueba.
- `main` se reserva para produccion.
- Cada push a `qa` puede correr pruebas headless sobre Apps Script.
- La suite smoke corre como validacion rapida.
- La suite completa se puede activar de forma manual o por variable.
- Produccion se despliega manualmente desde GitHub Actions.

Ademas, produccion usa un bundle reducido:

- Excluye archivos de testing, migracion, backup y export.
- Recorta funciones manuales o destructivas que no pertenecen al flujo normal del ERP.

## Scripts utiles

```bash
npm run push:gas:test
npm run test:qtas:probe:qa
npm run test:qtas:smoke:qa
npm run test:qtas:qa
npm run build:gas:prod
npm run deploy:prod:push
```

## Puesta en marcha

### Requisitos

- Node.js 20+
- `clasp`
- Un spreadsheet para produccion
- Una copia separada para QA
- Un proyecto de Apps Script ligado a cada hoja

### Paso a paso

1. Crea una hoja de produccion y una copia separada para QA.
2. Crea o vincula un proyecto de Apps Script a cada archivo.
3. Configura `.clasp.json` para produccion.
4. Configura `.clasp.test.json` para QA.
5. Instala dependencias con `npm install`.
6. Haz push a QA y valida el probe y la smoke suite.
7. Despliega a produccion solo cuando QA haya pasado.

## Interfaz

La app corre dentro de Google Sheets y hoy incluye tres vistas principales:

- Ventas
- Compras
- Configuracion avanzada

La experiencia esta pensada para uso operativo diario, con estados visibles, catalogos guiados y formularios rapidos.

## Documentacion complementaria

- [TESTING_QTAS.md](./TESTING_QTAS.md)
- [OPERACION_QTAS.md](./OPERACION_QTAS.md)
- [LOOKER_QTAS.md](./LOOKER_QTAS.md)

## Estado actual del proyecto

Hoy el sistema ya cuenta con:

- flujo de ventas operativo
- flujo de compras operativo
- costos historicos desde compras
- productos compuestos y reglas de costo
- analitica incremental
- QA automatizado
- despliegue separado para QA y produccion
- poda de bundle para produccion

## Nota

El repositorio contiene codigo y estructura tecnica. Los datos reales del negocio deben permanecer en hojas separadas y privadas.
