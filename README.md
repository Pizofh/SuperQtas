# SuperQTAS

SuperQTAS is a spreadsheet-native ERP built with Google Sheets and Google Apps Script for Psylo Scibio operations. It centralizes sales, purchases, receivables, historical costs, and operational configuration in a single interface embedded directly inside the workbook.

The project is intentionally lightweight. It does not try to mimic a large enterprise ERP; it focuses on operational reliability, traceability, incremental automation, and low maintenance overhead for a small business workflow.

## What It Does

- Registers sales with multiple products, discounts, upfront payments, and open balances.
- Keeps receivables live with pending sales, later payments, and debtor summaries.
- Tracks pending shipments inside the same commercial workflow.
- Registers purchases by line item, supplier, payment method, funding source, and operational note.
- Updates historical reference costs from real purchase data.
- Calculates costs and margins for composite products using recipes and cost rules.
- Tracks direct, manufactured, and recipe-driven inventory through a stock ledger and snapshot.
- Registers production batches and consumes the corresponding recipe components.
- Lets users manage products, prices, payment methods, and distribution rules without editing code.
- Keeps QA and production environments separated so changes can be tested before release.

## How It Works

The solution is split into four layers:

1. Google Sheets as the operational data store.
2. Apps Script as the backend and transaction layer.
3. HTML/CSS/JavaScript as the embedded interface.
4. `clasp`, Node.js, and GitHub Actions for deployment and automation.

In practice, the user opens the spreadsheet, launches the `QTAS ERP` menu, works from the SuperQTAS interface, and the scripts update the canonical sheets behind the scenes.

## Core Modules

### Sales

- Multi-line sales registration.
- Full or partial payment at the time of sale.
- Later payment registration for open balances.
- Debtors and open-sales dashboard.
- Shipment follow-up for sales marked as pending shipment.

### Purchases

- Purchase registration by supplier and payment method.
- Suggested item catalog to reduce manual typing.
- Canonicalization of known items when saving.
- Per-line option to decide whether the purchase affects historical costs.

### Costs and Analytics

- Current and historical reference costs.
- Composite products with components and variable cost rules.
- Incremental calculation of `Costo_Producto_Calc`.
- Incremental calculation of `Venta_Detalle_Costos_Calc`.
- Cost coverage, costing method, and estimated margin per sold line.

### Inventory and Production

- Inventory controls with direct, manufactured, recipe-driven, and non-controlled modes.
- Immutable movement ledger and current-stock snapshot.
- Production batches with component consumption and finished-product entries.
- Inventory status for supply-chain monitoring.
- A reversible historical-load pause that records commercial data without changing stock until operations are current.

### Operational Configuration

- Active and inactive product catalog.
- Price history with date-based validity.
- Configurable payment methods.
- Date-based income distribution rules.

## Operational Flow

### Sales

1. Products, quantities, prices, and discounts are selected.
2. Upfront payments are registered when applicable.
3. `Ventas`, `Venta_Detalle`, and `Pagos` are updated.
4. Income distribution is synchronized.
5. Incremental cost analytics are refreshed for that sale.

### Purchases

1. Supplier, payment method, and purchase lines are registered.
2. Each line can affect or ignore reference cost updates.
3. `Compras`, `Compra_Detalle`, and `Costos_Referencia` are updated.
4. The operational product-cost snapshot is refreshed.

### Configuration

1. Products, prices, payment methods, or distribution rules are edited from the UI.
2. Changes are applied without breaking historical records.

## Main Model Sheets

These are the most important operational sheets in the model:

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
- `Inventario_Control`
- `Inventario_Movimientos`
- `Inventario_Snapshot`
- `Producciones`
- `Produccion_Detalle`
- `Distribucion_Reglas`
- `Distribucion_Ingresos`
- `Origenes_Fondos_Reglas`
- `Compra_Origenes_Fondos`
- `Config_MediosPago`

## Repository Architecture

- [Codigo.gs](./Codigo.gs) defines constants, schemas, and base data.
- [QTAS_Ventas.gs](./QTAS_Ventas.gs) contains the commercial flow for sales, payments, and shipments.
- [QTAS_Compras.gs](./QTAS_Compras.gs) handles purchases and historical costs.
- [QTAS_CostosProducto.gs](./QTAS_CostosProducto.gs) resolves recipes, composite costing, and margins.
- [QTAS_Inventario.gs](./QTAS_Inventario.gs) manages inventory movements, current stock, and production batches.
- [QTAS_Distribucion.gs](./QTAS_Distribucion.gs) calculates income distribution.
- [QTAS_Admin.gs](./QTAS_Admin.gs) exposes advanced configuration to the front end.
- [QTAS_Modelo.gs](./QTAS_Modelo.gs) ensures and validates the canonical structure.
- [QTAS_Utils.gs](./QTAS_Utils.gs) contains shared utilities.
- [App.html](./App.html) contains the embedded web interface.

## Quality, QA, and Deployment

The repository uses **one GitHub Actions pipeline**: [Quality and Deploy](./.github/workflows/quality-deploy.yml), and **one test entrypoint**: `tests/suite.mjs`.

- Pull requests to `main`: every unit test + production build/surface verification. No deployment.
- Push to `qa`: same local quality gates + deployment to isolated QA + Apps Script probe + critical smoke suite.
- Manual `full_qa`: quality gates + complete QA E2E suite of 24 scenarios.
- Manual `release` on `main`: local quality gates + **complete QA E2E suite**, then (and only then) production deploy with `DEPLOY_PROD` confirmation and `production` environment gate.
- Manual `diagnose`: select an existing scenario from the same suite for troubleshooting; no independent tests or production release.

QA credentials are mandatory; missing credentials fail the gate. Smoke scenarios are already included in the full suite and are never repeated in the same release run. E2E tests reset an isolated QA spreadsheet, **never the production spreadsheet**. All QA E2E jobs are serialized through one concurrency group.

Production is deployed from a reduced verified bundle excluding test and destructive helpers. See [TESTING_QTAS.md](./TESTING_QTAS.md) for commands, secrets, and release criteria.

## Useful Scripts

```bash
npm run test:suite -- --level unit
npm run test:suite -- --level list
npm run test:suite -- --level smoke --config tests/qtas.test.config.json
npm run test:suite -- --level full --config tests/qtas.test.config.json
npm run test:suite -- --level diagnose --scenario venta_con_deuda --config tests/qtas.test.config.json
npm run build:gas:prod
npm run verify:gas:prod
```

For deployment, use **Quality and Deploy** in GitHub Actions, not a separate manual production workflow.

## Setup

### Requirements

- Node.js 20+
- `clasp`
- One spreadsheet for production
- A separate copy for QA
- One Apps Script project linked to each spreadsheet

### Step by Step

1. Create a production spreadsheet and a separate QA copy.
2. Create or link one Apps Script project to each file.
3. Configure `.clasp.json` for production.
4. Configure `.clasp.test.json` for QA.
5. Install dependencies with `npm install`.
6. Push to `qa` to execute the integrated smoke gate.
7. Use the manual `release` operation on `main` to run full E2E before production.

## Interface

The application runs inside Google Sheets and currently includes three main views:

- Sales
- Purchases
- Advanced configuration

The experience is designed for daily operational use, with guided catalogs, direct status feedback, and fast forms for both desktop and side-panel usage.

## Additional Documentation

- [Architecture and release diagrams](./docs/ARCHITECTURE_OVERVIEW.md)
- [TESTING_QTAS.md](./TESTING_QTAS.md)
- [OPERACION_QTAS.md](./OPERACION_QTAS.md)
- [LOOKER_QTAS.md](./LOOKER_QTAS.md)

## Current State

The system already includes:

- operational sales flow
- operational purchase flow
- historical cost tracking from purchases
- composite products and cost rules
- incremental analytics
- automated QA
- separated QA and production deployment
- production bundle pruning
- removal of old migration and export modules from the active codebase
- inventory and production control

## Note

This repository contains code and technical structure. Real business data should remain in separate private spreadsheets.
