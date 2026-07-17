import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';

const PROD_DIR = path.join(process.cwd(), 'dist', 'prod-appscript');

const FORBIDDEN_FILES = [
  'QTAS_Export.gs',
  'QTAS_Migracion.gs',
  'QTAS_Testing.gs'
];

const FORBIDDEN_FUNCTIONS = {
  'QTAS_CostosProducto.gs': [
    'reconstruirCostoProductoCalculadoQTAS',
    'reconstruirVentaDetalleCostosCalculadoQTAS',
    'reconstruirAnaliticaCostosQTAS',
    'reconstruirAnaliticaCostosQTAS_Log',
    'sincronizarPlantillaCosteoBaseQTAS',
    'sincronizarPlantillaCosteoBaseQTAS_Log',
    'sembrarCostosDirectosBaseQTAS',
    'sembrarCostosDirectosBaseQTAS_Log',
    'sembrarCostosDirectosBaseQTAS_Historico',
    'sembrarCostosDirectosBaseQTAS_Historico_Log',
    'repararCoberturaHistoricaCostosQTAS',
    'repararCoberturaHistoricaCostosQTAS_Log',
    'sincronizarPlantillaYAnaliticaCostosQTAS',
    'sincronizarPlantillaYAnaliticaCostosQTAS_Log'
  ],
  'QTAS_Admin.gs': [
    'auditarIntegridadFinancieraQTAS',
    'corregirVenta2342TemporalQTAS',
    'corregirVenta2342TemporalQTAS_Log',
    'corregirVentasHistoricasConfirmadasQTAS',
    'corregirVentasHistoricasConfirmadasQTAS_Log',
    'repararIntegridadFinancieraQTAS',
    'continuarReparacionIntegridadFinancieraQTAS',
    'estadoReparacionIntegridadFinancieraQTAS'
  ],
  'QTAS_Distribucion.gs': [
    'reconstruirDistribucionIngresosQTAS_'
  ],
  'QTAS_Modelo.gs': [
    'asegurarModeloCompletoQTAS_',
    'asegurarModeloCompletoQTAS',
    'diagnosticarModeloQTAS',
    'snapshotEstructuraModeloQTAS'
  ],
  'QTAS_Ventas.gs': [
    'recalcularSaldosQTAS'
  ]
};

async function main() {
  await assertProdBundleExists_();

  const entries = await fs.readdir(PROD_DIR, { withFileTypes: true });
  const files = entries.filter(entry => entry.isFile()).map(entry => entry.name);

  const forbiddenFilesFound = FORBIDDEN_FILES.filter(file => files.includes(file));
  if (forbiddenFilesFound.length) {
    throw new Error(
      `Archivos no permitidos en prod: ${forbiddenFilesFound.join(', ')}`
    );
  }

  const forbiddenFunctionsFound = [];
  for (const [fileName, functionNames] of Object.entries(FORBIDDEN_FUNCTIONS)) {
    const filePath = path.join(PROD_DIR, fileName);
    const content = await readIfExists_(filePath);
    if (!content) continue;

    functionNames.forEach(functionName => {
      const pattern = new RegExp(`\\bfunction\\s+${escapeRegExp_(functionName)}\\s*\\(`);
      if (pattern.test(content)) {
        forbiddenFunctionsFound.push(`${fileName}:${functionName}`);
      }
    });
  }

  if (forbiddenFunctionsFound.length) {
    throw new Error(
      `Funciones no permitidas en prod: ${forbiddenFunctionsFound.join(', ')}`
    );
  }

  console.log(JSON.stringify({
    ok: true,
    prodDir: path.relative(process.cwd(), PROD_DIR).replace(/\\/g, '/'),
    checkedFiles: files.sort((a, b) => a.localeCompare(b)),
    forbiddenFiles: FORBIDDEN_FILES,
    forbiddenFunctions: FORBIDDEN_FUNCTIONS
  }, null, 2));
}

async function assertProdBundleExists_() {
  try {
    const stat = await fs.stat(PROD_DIR);
    if (!stat.isDirectory()) {
      throw new Error('dist/prod-appscript no es un directorio.');
    }
  } catch (error) {
    throw new Error(
      'No existe dist/prod-appscript. Ejecuta primero "npm run build:gas:prod".'
    );
  }
}

async function readIfExists_(filePath) {
  try {
    return await fs.readFile(filePath, 'utf8');
  } catch (error) {
    if (error && error.code === 'ENOENT') {
      return '';
    }
    throw error;
  }
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

main().catch(error => {
  console.error(error.message || error);
  process.exitCode = 1;
});
