# Testing QTAS

Este repo ahora incluye un runner local para probar flujos reales de Apps Script contra una hoja de prueba.

## Que cubre

- Ventas con deuda
- Ventas sin deuda
- Ventas con varios productos
- Ventas con varios medios de pago
- Pagos pendientes
- Compras de producto con impacto en costos
- Compras de gasto sin impacto en costos
- Compras mixtas
- Reglas de distribucion
- Configuracion de productos, precios y medios de pago

## Recomendacion importante

No apuntes esto a tu archivo de produccion.

La forma segura es:

1. Dejar `.clasp.json` apuntando a produccion.
2. Crear otra copia de tu spreadsheet de QTAS solo para pruebas.
3. Crear un Apps Script project ligado a esa copia de pruebas.
4. Conectar el runner y `clasp` de pruebas a ese proyecto separado.

## Esquema recomendado de ambientes

- `produccion`
  - Spreadsheet: tu hoja principal
  - Apps Script: el proyecto normal que usas dia a dia
  - Archivo local: `.clasp.json`
- `pruebas`
  - Spreadsheet: una copia separada
  - Apps Script: otro proyecto ligado a esa copia
  - Archivo local: `.clasp.test.json`
  - Config del runner: `tests/qtas.test.config.json`

## Archivos nuevos

- `QTAS_Testing.gs`
  - Helpers para resetear el entorno y tomar snapshots completos de las hojas.
- `tests/run-qtas-tests.mjs`
  - Runner local que ejecuta escenarios via Apps Script API.
- `tests/scenarios.mjs`
  - Escenarios automatizados.
- `.clasp.test.json.example`
  - Plantilla para conectar `clasp` al proyecto de pruebas.
- `tests/qtas.test.config.example.json`
  - Plantilla de configuracion local del runner de pruebas.

## Setup

### 1. Instalar dependencias

```powershell
npm.cmd install
```

### 2. Configurar `clasp`

```powershell
npx clasp login
```

Crea o conserva estos archivos asi:

- `.clasp.json`
  - `scriptId` del proyecto de produccion
- `.clasp.test.json`
  - a partir de `.clasp.test.json.example`
  - `scriptId` del proyecto de pruebas

Comandos utiles:

```powershell
npm.cmd run push:gas
npm.cmd run push:gas:prod
npm.cmd run push:gas:test
```

### 3. Habilitar Apps Script API

Google exige que el proyecto se despliegue como API executable y que la app cliente use OAuth.

Pasos recomendados:

1. Vincula tu Apps Script project a un Google Cloud project estandar.
2. Habilita Apps Script API en ese Cloud project.
3. Crea credenciales OAuth para Desktop app.
4. Descarga el JSON y guardalo como `tests/google-oauth.credentials.json`.
5. En Apps Script, crea un deployment tipo `API executable`.

Referencia oficial:

- Apps Script API `scripts.run`: https://developers.google.com/apps-script/api/how-tos/execute
- `clasp`: https://developers.google.com/apps-script/guides/clasp

## Autenticacion interactiva vs headless

Hasta ahora el runner usaba `@google-cloud/local-auth`, que abre navegador para que inicies sesion.

Eso sirve bien en tu maquina, pero no sirve bien en CI porque GitHub Actions no puede abrir ese login interactivo.

Ahora el runner soporta tres modos:

- `authMode: "auto"`
  - primero intenta `authorizedUserPath`
  - luego `serviceAccountPath`
  - si no encuentra nada, cae a login interactivo con `credentialsPath`
- `authMode: "authorized_user"`
  - usa una credencial OAuth reutilizable con refresh token
  - es la mejor opcion para trabajar comodo y para CI
- `authMode: "service_account"`
  - util si quieres un robot dedicado, pero requiere compartir script y spreadsheet con esa cuenta

### Bootstrap recomendado para modo headless

Genera una vez el archivo reutilizable:

```powershell
npm.cmd run auth:qtas:bootstrap
```

Eso crea:

- `tests/google-oauth.authorized-user.json`

Luego en tu config de pruebas puedes dejar:

```json
{
  "deploymentId": "YOUR_TEST_API_EXECUTABLE_DEPLOYMENT_ID",
  "scriptProjectId": "YOUR_TEST_APPS_SCRIPT_PROJECT_ID",
  "claspProjectFile": "./.clasp.test.json",
  "credentialsPath": "./tests/google-oauth.credentials.json",
  "authorizedUserPath": "./tests/google-oauth.authorized-user.json",
  "authMode": "auto",
  "devMode": true
}
```

Con eso, los siguientes `npm run test:qtas` y `npm run test:qtas:probe` ya no deberian pedir navegador cada vez.

### 4. Crear configuracion local

Copia `tests/qtas.test.config.example.json` a `tests/qtas.test.config.json`.

Compatibilidad:

- Si ya vienes usando `tests/qtas.config.json`, el runner tambien lo acepta.
- Si existe `.clasp.test.json`, el runner lo prefiere sobre `.clasp.json`.

Campos importantes:

- `deploymentId`: ID del deployment `API executable`.
- `scriptProjectId`: opcional. ID del proyecto de Apps Script de pruebas.
- `claspProjectFile`: archivo `clasp` que el runner debe usar para inferir el script de pruebas.
- `credentialsPath`: ruta al JSON OAuth descargado desde Google Cloud.
- `devMode`: si `true`, ejecuta la ultima version guardada del proyecto.
- `failureArtifactPath`: donde guardar el snapshot del ultimo fallo.

Importante:

- `.clasp.json` es produccion.
- `.clasp.test.json` es pruebas.
- `tests/qtas.test.config.json` apunta al deployment de pruebas.
- `scriptId` y `deploymentId` no son el mismo valor.

## Como correr

Listar escenarios:

```powershell
npm.cmd run test:qtas:list
```

Correr toda la suite:

```powershell
npm.cmd run test:qtas
```

Hacer un diagnostico rapido de conexion:

```powershell
npm.cmd run test:qtas:probe
```

Ese probe revisa:

- si el `scriptProjectId` responde
- si el `deploymentId` existe dentro de ese proyecto
- si `testPingQTAS()` realmente ejecuta

## CI/CD recomendado

El repo ya puede quedar con flujo de despliegue por GitHub Actions sin volver a copiar codigo manualmente al editor web.

Flujo sugerido:

1. Trabajar cambios en git.
2. Hacer merge a `main`.
3. GitHub Actions hace `push` al Apps Script de pruebas.
4. Validar pruebas funcionales contra la hoja de pruebas.
5. Lanzar deploy manual a produccion.

Workflows incluidos:

- `.github/workflows/deploy-test.yml`
  - corre en `push` a `main` y tambien manualmente
  - instala dependencias
  - escribe `.clasprc.json` y `.clasp.test.json` desde secrets
  - hace `push` al proyecto de pruebas
  - si tambien existen secrets de auth/config de pruebas, corre `probe` y suite E2E en modo headless
- `.github/workflows/deploy-prod.yml`
  - corre solo manualmente
  - usa el environment `production`
  - hace `push` al proyecto de produccion

Secrets esperados en GitHub:

- `CLASPRC_JSON`
- `CLASP_JSON_TEST`
- `CLASP_JSON_PROD`
- `QTAS_TEST_CONFIG_JSON`
- `GOOGLE_AUTHORIZED_USER_JSON_TEST`

### Limite actual importante

El modo interactivo sigue existiendo como fallback.

Pero si quieres verdadera comodidad:

- local: usa `authorized_user`
- CI: usa el secret `GOOGLE_AUTHORIZED_USER_JSON_TEST`

Asi el deploy y las pruebas funcionales ya pueden correr sin abrir navegador.

Correr un escenario puntual:

```powershell
node tests/run-qtas-tests.mjs --scenario venta_con_deuda
```

Correr por tag:

```powershell
node tests/run-qtas-tests.mjs --tag ventas
node tests/run-qtas-tests.mjs --tag compras
```

## Como funciona

El runner llama funciones reales del proyecto con Apps Script API:

- `registrarVentaQTAS`
- `registrarPagoPendienteQTAS`
- `registrarCompraQTAS`
- `guardarProductoConfiguracionQTAS`
- `guardarCambioPrecioFrontendQTAS`
- `guardarMedioPagoQTAS`
- `cambiarEstadoMedioPagoQTAS`
- `guardarReglaDistribucionFrontendQTAS`

Antes de cada escenario, `testResetEntornoQTAS()` limpia hojas transaccionales, resetea secuencias persistentes y vuelve a sembrar:

- productos
- precios
- medios de pago
- reglas de distribucion

Despues, `testSnapshotQTAS()` devuelve el estado serializado de:

- hojas principales
- dashboard de ventas
- costos vigentes
- compras recientes
- configuracion avanzada

Si un escenario falla, el runner guarda un snapshot en `tests/.artifacts/last-failure.json`.

## Nota sobre OAuth

La primera vez, Google va a abrir el navegador para autorizar el acceso. El runner necesita permisos suficientes para ejecutar el script y acceder a la hoja que usa el proyecto.

## Migracion desde el esquema viejo

Si antes tenias todo mezclado en una sola copia:

1. Deja `.clasp.json` apuntando a la hoja que vas a usar como produccion.
2. Crea otra copia solo para test.
3. Crea `.clasp.test.json` con el `scriptId` de ese proyecto de pruebas.
4. Crea `tests/qtas.test.config.json` con el `deploymentId` de pruebas.
5. Usa `npm.cmd run push:gas:test` y `npm.cmd run test:qtas` solo contra ese ambiente de pruebas.

Si todavia usas `tests/qtas.config.json`, no hay problema: el runner lo toma como fallback. Cuando quieras dejarlo mas ordenado, migralo a `tests/qtas.test.config.json`. Si ya tienes `.clasp.json` real configurado para produccion, `.clasp.json.example` tambien pasa a ser opcional.

## Conviene clonar la copia de pruebas

Si tu copia de pruebas ya esta funcionando bien con el runner, normalmente conviene dejarla quieta como ambiente estable de test.

Patron recomendado:

1. `produccion`: una copia estable para uso real.
2. `pruebas`: una copia estable conectada a `clasp`, al runner y al deployment de Apps Script API.
3. `clones temporales`: solo para experimentos riesgosos, migraciones o limpiezas grandes.

Importante:

- Clonar la hoja de pruebas no evita todos los cambios de configuracion.
- Una copia nueva normalmente termina con otro `spreadsheetId`.
- Si el proyecto de Apps Script tambien cambia, tendras otro `scriptId`.
- Si quieres correr el runner contra esa copia nueva, tambien tendras que crear o actualizar su `deploymentId`.
- Lo que usualmente si puedes reutilizar es el archivo OAuth de desktop (`tests/google-oauth.credentials.json`), siempre que sigas usando el mismo cliente OAuth y la misma cuenta autorizada.

En otras palabras:

- Si lo que quieres es no volver a tocar IDs ni configs, no reemplaces tu ambiente de pruebas actual.
- Si lo que quieres es abrir un ambiente nuevo aislado, clona, pero asume que tendras que actualizar al menos `scriptId` y `deploymentId` de ese clon.

## Siguiente mejora natural

Si despues quieres automatizar tambien la UI HTML, el siguiente paso seria sumar Playwright encima de este harness. Primero conviene dejar estable esta capa de pruebas funcionales reales contra Sheets/App Script.
