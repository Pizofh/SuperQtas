import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';

const TARGETS = {
  prod: {
    sourceProjectFile: '.clasp.json',
    generatedProjectFile: '.clasp.prod-build.json',
    outputDir: 'dist/prod-appscript',
    includeFiles: [
      'App.html',
      'appsscript.json',
      'Codigo.gs',
      'QTAS_Admin.gs',
      'QTAS_Compras.gs',
      'QTAS_CostosProducto.gs',
      'QTAS_Distribucion.gs',
      'QTAS_Modelo.gs',
      'QTAS_Precios.gs',
      'QTAS_UI.gs',
      'QTAS_Utils.gs',
      'QTAS_Ventas.gs'
    ]
  }
};

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const target = TARGETS[args.target];

  if (!target) {
    throw new Error(`Target desconocido: ${args.target}. Usa uno de: ${Object.keys(TARGETS).join(', ')}`);
  }

  const rootDir = process.cwd();
  const outputDirAbs = path.join(rootDir, target.outputDir);
  const generatedProjectFileAbs = path.join(rootDir, target.generatedProjectFile);
  const sourceProjectFileAbs = path.join(rootDir, target.sourceProjectFile);

  await fs.rm(outputDirAbs, { recursive: true, force: true });
  await fs.mkdir(outputDirAbs, { recursive: true });

  await Promise.all(target.includeFiles.map(async relativeFile => {
    const sourceAbs = path.join(rootDir, relativeFile);
    const targetAbs = path.join(outputDirAbs, relativeFile);
    await fs.copyFile(sourceAbs, targetAbs);
  }));

  const sourceProject = JSON.parse(await fs.readFile(sourceProjectFileAbs, 'utf8'));
  const generatedProject = {
    ...sourceProject,
    rootDir: target.outputDir.replace(/\\/g, '/')
  };

  await fs.writeFile(
    generatedProjectFileAbs,
    `${JSON.stringify(generatedProject, null, 2)}\n`,
    'utf8'
  );

  const sourceFiles = await listarArchivosAppsScriptRaiz_(rootDir);
  const includeSet = new Set(target.includeFiles);
  const excludedFiles = sourceFiles.filter(file => !includeSet.has(file));

  console.log(JSON.stringify({
    ok: true,
    target: args.target,
    outputDir: target.outputDir,
    generatedProjectFile: target.generatedProjectFile,
    includedFiles: target.includeFiles,
    excludedFiles
  }, null, 2));
}

function parseArgs(argv) {
  const args = {
    target: 'prod'
  };

  for (let index = 0; index < argv.length; index += 1) {
    const current = argv[index];
    if (current === '--target' && argv[index + 1]) {
      args.target = String(argv[index + 1]).trim();
      index += 1;
    }
  }

  return args;
}

async function listarArchivosAppsScriptRaiz_(rootDir) {
  const entries = await fs.readdir(rootDir, { withFileTypes: true });
  return entries
    .filter(entry => entry.isFile())
    .map(entry => entry.name)
    .filter(name => name === 'appsscript.json' || /\.(gs|html)$/i.test(name))
    .filter(name => !name.startsWith('.clasp'))
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
}

main().catch(error => {
  console.error(error.message || error);
  process.exitCode = 1;
});
