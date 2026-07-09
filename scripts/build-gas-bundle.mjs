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
      'QTAS_Backups.gs',
      'QTAS_Compras.gs',
      'QTAS_CostosProducto.gs',
      'QTAS_Distribucion.gs',
      'QTAS_Libro.gs',
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
    const sourceContent = await fs.readFile(sourceAbs, 'utf8');
    const stripFunctions = ((target.stripFunctionsByFile || {})[relativeFile] || []).slice();
    const targetContent = stripFunctions.length
      ? stripTopLevelFunctions_(sourceContent, stripFunctions, relativeFile)
      : sourceContent;

    await fs.writeFile(targetAbs, targetContent, 'utf8');
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
    strippedFunctionsByFile: target.stripFunctionsByFile || {},
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

function stripTopLevelFunctions_(source, functionNames, relativeFile) {
  return functionNames.reduce((currentSource, functionName) =>
    stripSingleTopLevelFunction_(currentSource, functionName, relativeFile), source);
}

function stripSingleTopLevelFunction_(source, functionName, relativeFile) {
  const signature = new RegExp(`(^|\\r?\\n)([\\t ]*)function\\s+${escapeRegExp_(functionName)}\\s*\\(`, 'm');
  const match = signature.exec(source);
  if (!match) {
    throw new Error(`No se encontro function ${functionName} en ${relativeFile}.`);
  }

  const functionStart = match.index + match[1].length;
  const openParenIndex = source.indexOf('(', functionStart);
  if (openParenIndex < 0) {
    throw new Error(`No se encontro la firma de ${functionName} en ${relativeFile}.`);
  }

  const bodyStart = findFunctionBodyStart_(source, openParenIndex, relativeFile, functionName);
  const functionEnd = findMatchingBraceEnd_(source, bodyStart, relativeFile, functionName);

  return source.slice(0, functionStart) + source.slice(functionEnd + 1);
}

function findFunctionBodyStart_(source, openParenIndex, relativeFile, functionName) {
  let parenDepth = 0;

  for (let index = openParenIndex; index < source.length; index += 1) {
    const skipIndex = skipLiteralOrComment_(source, index);
    if (skipIndex > index) {
      index = skipIndex;
      continue;
    }

    const current = source[index];
    if (current === '(') {
      parenDepth += 1;
      continue;
    }

    if (current === ')') {
      parenDepth -= 1;
      if (parenDepth === 0) {
        return findNextOpeningBrace_(source, index + 1, relativeFile, functionName);
      }
    }
  }

  throw new Error(`No se encontro el cuerpo de ${functionName} en ${relativeFile}.`);
}

function findNextOpeningBrace_(source, startIndex, relativeFile, functionName) {
  for (let index = startIndex; index < source.length; index += 1) {
    const skipIndex = skipLiteralOrComment_(source, index);
    if (skipIndex > index) {
      index = skipIndex;
      continue;
    }

    const current = source[index];
    if (/\s/.test(current)) continue;
    if (current === '{') return index;

    throw new Error(`Se esperaba "{" en ${functionName} (${relativeFile}).`);
  }

  throw new Error(`No se encontro "{" en ${functionName} (${relativeFile}).`);
}

function findMatchingBraceEnd_(source, bodyStart, relativeFile, functionName) {
  let braceDepth = 0;

  for (let index = bodyStart; index < source.length; index += 1) {
    const skipIndex = skipLiteralOrComment_(source, index);
    if (skipIndex > index) {
      index = skipIndex;
      continue;
    }

    const current = source[index];
    if (current === '{') {
      braceDepth += 1;
      continue;
    }

    if (current === '}') {
      braceDepth -= 1;
      if (braceDepth === 0) {
        return index;
      }
    }
  }

  throw new Error(`No se encontro el cierre de ${functionName} en ${relativeFile}.`);
}

function skipLiteralOrComment_(source, startIndex) {
  const current = source[startIndex];
  const next = source[startIndex + 1];

  if ((current === '\'' || current === '"' || current === '`')) {
    return skipQuotedLiteral_(source, startIndex, current);
  }

  if (current === '/' && next === '/') {
    let index = startIndex + 2;
    while (index < source.length && source[index] !== '\n') {
      index += 1;
    }
    return index - 1;
  }

  if (current === '/' && next === '*') {
    let index = startIndex + 2;
    while (index < source.length - 1) {
      if (source[index] === '*' && source[index + 1] === '/') {
        return index + 1;
      }
      index += 1;
    }
    return source.length - 1;
  }

  return startIndex;
}

function skipQuotedLiteral_(source, startIndex, quoteChar) {
  let escaping = false;

  for (let index = startIndex + 1; index < source.length; index += 1) {
    const current = source[index];

    if (escaping) {
      escaping = false;
      continue;
    }

    if (current === '\\') {
      escaping = true;
      continue;
    }

    if (current === quoteChar) {
      return index;
    }
  }

  return source.length - 1;
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

main().catch(error => {
  console.error(error.message || error);
  process.exitCode = 1;
});
