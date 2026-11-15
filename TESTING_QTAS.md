# SuperQTAS - suite unificada de calidad y despliegue

**Un solo workflow:** `.github/workflows/quality-deploy.yml`. Los escenarios se mantienen modularizados para mantenimiento, pero todos se ejecutan desde el mismo punto de entrada: `tests/suite.mjs`.

## Puertas de calidad

| Disparador | Validaciones | Despliegue |
|---|---|---|
| Pull request hacia `main` | Todas las pruebas unitarias + build y verificación de superficie de producción | Ninguno |
| Push a `qa` | Unitarias + build/verificación + deploy QA + probe + 10 escenarios smoke (mínimo exigido) | Solo QA |
| Actions → Quality and Deploy → `full_qa` | Unitarias + build/verificación + deploy QA + probe + 24 escenarios E2E (mínimo exigido) | Solo QA |
| Actions → Quality and Deploy → `release` en `main` | Todas las puertas anteriores **con E2E completa**; producción solo después de aprobarlas | QA y luego producción |
| Actions → Quality and Deploy → `diagnose` | Puertas locales + deploy QA + probe + escenario indicado de la misma suite | Solo QA |

**No hay pipeline alterno para saltarse QA:** producción requiere selección `release` desde `main`, confirmación `DEPLOY_PROD`, todas las pruebas aprobadas y el entorno GitHub `production`. Configura *required reviewers* en ese entorno para una aprobación humana adicional.

**La E2E completa incluye los escenarios smoke**, por lo que no se ejecutan dos veces en un release.

## Comandos locales

```bash
npm ci
npm run test:suite -- --level unit
npm run test:suite -- --level list
npm run test:suite -- --level probe --config tests/qtas.test.config.json
npm run test:suite -- --level smoke --config tests/qtas.test.config.json
npm run test:suite -- --level full --config tests/qtas.test.config.json
npm run test:suite -- --level diagnose --scenario venta_eliminacion_reciente_reusa_id --config tests/qtas.test.config.json
```

**Un solo runner:** `tests/suite.mjs`.

- `unit`: descubre automáticamente todos los archivos `tests/*.test.mjs`. Incluye seguridad de datos, integridad de inventario y gastos sin recálculo.
- `probe`: valida permisos y conectividad del proyecto y despliegue Apps Script **de QA**.
- `smoke`: escenario crítico etiquetado `smoke`, mínimo 10.
- `full`: todos los escenarios de `tests/scenarios.mjs`, mínimo 24; conserva aislamiento con `testResetEntornoQTAS`.
- `diagnose`: ejecuta **uno de esos mismos escenarios**, no otra suite. Nunca autoriza producción.

Todos los niveles E2E usan el runner de Apps Script `tests/run-qtas-tests.mjs`, sin duplicar casos. Por defecto se activa *fail-fast*: al encontrar un fallo se detiene la ejecución, pero un release solo es válido si pasan **todos** los escenarios del nivel requerido. El runner registra las llamadas lentas a Apps Script.

## Configuración de QA

Mantener ambientes separados:

| Ambiente | Datos | Apps Script | Configuración |
|---|---|---|---|
| QA | Copia aislada del libro | Script de QA | `.clasp.test.json`, `tests/qtas.test.config.json` |
| Producción | Libro operativo real | Script de producción | `.clasp.json` |

**Secrets obligatorios** en GitHub Actions:

- `CLASPRC_JSON`
- `CLASP_JSON_TEST`
- `QTAS_TEST_CONFIG_JSON`
- `GOOGLE_AUTHORIZED_USER_JSON_TEST`
- `CLASP_JSON_PROD` (solo producción)

La falta de credenciales de QA provoca un **fallo**, nunca un skip exitoso. No guardar secretos ni JSON OAuth en Git.

Para bootstrap local OAuth:

```bash
npm run auth:qtas:bootstrap
```

El runner de QA utiliza Apps Script API y la configuración indicada mediante `--config`; debe corresponder a la copia QA. `devMode: true` ejecuta el código desplegado mediante `clasp push` en ese proyecto.

## Seguridad operativa

- Los escenarios E2E pueden **borrar y restablecer datos de QA**. Nunca apuntarlos al libro productivo.
- La ejecución QA está serializada en GitHub Actions para impedir dos suites destructivas simultáneas desde este workflow.
- Una falla de CI, del probe o de E2E bloquea el job de producción; los deployments no deben considerarse validados por un simple `clasp push`.
- El build de producción excluye `QTAS_Testing.gs` y verifica funciones prohibidas; ninguna prueba E2E se ejecuta contra producción.
- Para despliegues sensibles, revisar cambios de scripts temporales, respaldos e historial antes de activar `release`.

## Criterio de rendimiento

No duplicamos la suite smoke dentro de la E2E completa. La validación de PR se resuelve sin servicios Google y los E2E solo se corren cuando corresponda. Si un escenario falla, se detiene antes de consumir tiempo en los restantes. Las trazas `[API SLOW]`, los tiempos por escenario y los artefactos de fallo permiten investigar la causa; el tiempo total de 24 escenarios sigue condicionado a la latencia de Google Sheets/Apps Script. No se promete una duración hasta medir el nuevo flujo.

## Antes de una liberación

1. Merge del código a `main` tras aprobación de CI.
2. Ejecutar `Quality and Deploy` con `operation=release` desde `main` y `confirm_prod=DEPLOY_PROD`.
3. Verificar aprobación de unitarias, bundle, QA probe y **24 escenarios E2E**.
4. Autorizar el environment de producción y verificar el despliegue.
5. Hacer las operaciones de corrección de inventario **separadamente** con respaldo y validación de datos reales.

El workflow no ejecuta correcciones de datos por sí mismo.
