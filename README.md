# SuperQTAS

**A lightweight Google Sheets ERP with an Apps Script interface and controlled release workflows.**

SuperQTAS brings common business operations into a single spreadsheet interface. The public repository demonstrates application organization, workflow testing, and delivery automation with Google Apps Script, JavaScript, clasp, and GitHub Actions.

## Project Highlights

- An embedded HTML, CSS, and JavaScript interface backed by Apps Script.
- Modules for sales, purchases, product costs, inventory, and production.
- Separate QA and production Apps Script projects.
- Headless test tooling for validating workflows in QA.
- Manually triggered production deployment with a reduced runtime bundle.

## Engineering Decisions

| Decision | Purpose |
| --- | --- |
| Google Sheets and Apps Script | Keep the application accessible inside a familiar operational tool |
| Separate QA and production | Validate changes in a dedicated test environment |
| Headless workflow tests | Check behavior before a release |
| Manual production releases | Make deployment an explicit action after validation |
| Production bundle verification | Keep test helpers outside the deployed runtime |

## Stack

Google Sheets · Google Apps Script · JavaScript · HTML/CSS · Node.js · clasp · GitHub Actions

## Getting Started

Requirements: Node.js 20+, a Google account, and separate spreadsheets and Apps Script projects for QA and production.

```bash
git clone https://github.com/Pizofh/SuperQtas.git
cd SuperQtas
npm ci
```

Copy the clasp example files into local configuration files and connect them to your own Apps Script projects. Configure QA before attempting deployment.

Dependencies alone do not create a working ERP instance. Spreadsheet setup, authorization, and test configuration are described in the guides below.

## Documentation

- [Operations and environment setup](./OPERACION_QTAS.md)
- [Testing and release validation](./TESTING_QTAS.md)
- [Analytics integration](./LOOKER_QTAS.md)

## Repository Guide

| Path | Responsibility |
| --- | --- |
| `Codigo.gs` | Shared definitions and base model |
| `QTAS_*.gs` | Application modules and supporting utilities |
| `App.html` | Embedded interface |
| `scripts/` | Bundle generation and runtime verification |
| `tests/` | Test harness and scenarios |
| `.github/workflows/` | QA and production delivery workflows |

## Scope

This repository presents application code and technical documentation. Each operational instance uses its own private spreadsheets and local configuration.
