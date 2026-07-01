import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { authenticate } from '@google-cloud/local-auth';
import { google } from 'googleapis';

const DEFAULT_SCOPES = [
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.deployments.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/userinfo.email'
];

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.left || !args.right) {
    throw new Error('Uso: node tests/compare-qtas-structure.mjs --left <configA> --right <configB>');
  }

  const leftConfig = await loadConfig(args.left);
  const rightConfig = await loadConfig(args.right);
  const auth = await authenticate({
    scopes: unirValoresUnicos_([
      ...(leftConfig.scopes || []),
      ...(rightConfig.scopes || []),
      ...DEFAULT_SCOPES
    ]),
    keyfilePath: leftConfig.credentialsPath
  });
  const scriptClient = google.script({ version: 'v1', auth });

  const left = await callAppsScript(scriptClient, leftConfig, 'snapshotEstructuraModeloQTAS');
  const right = await callAppsScript(scriptClient, rightConfig, 'snapshotEstructuraModeloQTAS');
  const report = compareSnapshots(left, right, {
    leftLabel: args.leftLabel || path.basename(args.left),
    rightLabel: args.rightLabel || path.basename(args.right)
  });

  printReport(report);
}

function parseArgs(argv) {
  const args = {
    left: '',
    right: '',
    leftLabel: '',
    rightLabel: ''
  };

  for (let index = 0; index < argv.length; index += 1) {
    const current = argv[index];
    if (current === '--left') {
      args.left = argv[index + 1] || '';
      index += 1;
      continue;
    }
    if (current === '--right') {
      args.right = argv[index + 1] || '';
      index += 1;
      continue;
    }
    if (current === '--left-label') {
      args.leftLabel = argv[index + 1] || '';
      index += 1;
      continue;
    }
    if (current === '--right-label') {
      args.rightLabel = argv[index + 1] || '';
      index += 1;
    }
  }

  return args;
}

async function loadConfig(configPath) {
  const target = resolveProjectPath(configPath);
  const raw = await fs.readFile(target, 'utf8');
  const userConfig = JSON.parse(raw);
  const claspConfig = await loadClaspConfig(userConfig.claspProjectFile);

  return {
    scopes: unirValoresUnicos_([...(userConfig.scopes || []), ...DEFAULT_SCOPES]),
    credentialsPath: resolveProjectPath(userConfig.credentialsPath || './tests/google-oauth.credentials.json'),
    deploymentId: userConfig.deploymentId || '',
    scriptProjectId: userConfig.scriptProjectId || userConfig.scriptId || claspConfig.scriptId || ''
  };
}

async function loadClaspConfig(configPath) {
  const target = resolveProjectPath(configPath || './.clasp.json');
  try {
    const raw = await fs.readFile(target, 'utf8');
    return JSON.parse(raw);
  } catch (error) {
    return {};
  }
}

async function callAppsScript(scriptClient, config, functionName, parameters = []) {
  const response = await scriptClient.scripts.run({
    scriptId: config.deploymentId,
    requestBody: {
      function: functionName,
      parameters,
      devMode: true
    }
  });
  const operation = response.data || {};

  if (operation.error) {
    const details = operation.error.details || [];
    const scriptError = details.find(item => item.errorMessage);
    throw new Error(scriptError ? scriptError.errorMessage : operation.error.message || 'Error de Apps Script.');
  }

  return operation.response ? operation.response.result : null;
}

function compareSnapshots(left, right, labels) {
  const leftSheets = indexByLogicalName(left?.sheets || []);
  const rightSheets = indexByLogicalName(right?.sheets || []);
  const sheetNames = unirValoresUnicos_([
    ...Object.keys(leftSheets),
    ...Object.keys(rightSheets)
  ]).sort();

  const differences = [];
  const rowCountDifferences = [];

  sheetNames.forEach(sheetName => {
    const leftSheet = leftSheets[sheetName] || emptySheetSnapshot_(sheetName);
    const rightSheet = rightSheets[sheetName] || emptySheetSnapshot_(sheetName);

    if (Boolean(leftSheet.exists) !== Boolean(rightSheet.exists)) {
      differences.push({
        type: 'exists',
        sheetName,
        left: leftSheet.exists,
        right: rightSheet.exists
      });
    }

    if (!arraysEqual_(leftSheet.actualHeaders || [], rightSheet.actualHeaders || [])) {
      differences.push({
        type: 'headers',
        sheetName,
        left: leftSheet.actualHeaders || [],
        right: rightSheet.actualHeaders || []
      });
    }

    if (Number(leftSheet.dataRows || 0) !== Number(rightSheet.dataRows || 0)) {
      rowCountDifferences.push({
        sheetName,
        left: Number(leftSheet.dataRows || 0),
        right: Number(rightSheet.dataRows || 0)
      });
    }
  });

  return {
    leftLabel: labels.leftLabel,
    rightLabel: labels.rightLabel,
    leftSpreadsheet: {
      name: left?.spreadsheetName || '',
      id: left?.spreadsheetId || ''
    },
    rightSpreadsheet: {
      name: right?.spreadsheetName || '',
      id: right?.spreadsheetId || ''
    },
    leftExtraSheets: (left?.extraSheets || []).slice().sort(),
    rightExtraSheets: (right?.extraSheets || []).slice().sort(),
    structuralDifferences: differences,
    rowCountDifferences: rowCountDifferences
  };
}

function printReport(report) {
  console.log(`Comparando ${report.leftLabel} vs ${report.rightLabel}\n`);
  console.log(`${report.leftLabel}: ${report.leftSpreadsheet.name} (${report.leftSpreadsheet.id})`);
  console.log(`${report.rightLabel}: ${report.rightSpreadsheet.name} (${report.rightSpreadsheet.id})\n`);

  if (!report.structuralDifferences.length) {
    console.log('Estructura: sin diferencias de hojas/headers entre ambos ambientes.\n');
  } else {
    console.log('Diferencias estructurales:\n');
    report.structuralDifferences.forEach(item => {
      if (item.type === 'exists') {
        console.log(`- ${item.sheetName}: existe en ${report.leftLabel}=${item.left} | ${report.rightLabel}=${item.right}`);
        return;
      }

      console.log(`- ${item.sheetName}: headers distintos`);
      console.log(`  ${report.leftLabel}: ${JSON.stringify(item.left)}`);
      console.log(`  ${report.rightLabel}: ${JSON.stringify(item.right)}`);
    });
    console.log('');
  }

  console.log('Diferencias de cantidad de filas:\n');
  if (!report.rowCountDifferences.length) {
    console.log('- Ninguna\n');
  } else {
    report.rowCountDifferences.forEach(item => {
      console.log(`- ${item.sheetName}: ${report.leftLabel}=${item.left} | ${report.rightLabel}=${item.right}`);
    });
    console.log('');
  }

  console.log('Hojas extra fuera del modelo canonico:\n');
  console.log(`- ${report.leftLabel}: ${report.leftExtraSheets.length ? report.leftExtraSheets.join(', ') : 'ninguna'}`);
  console.log(`- ${report.rightLabel}: ${report.rightExtraSheets.length ? report.rightExtraSheets.join(', ') : 'ninguna'}`);
}

function indexByLogicalName(rows) {
  return (rows || []).reduce((acc, row) => {
    if (row && row.logicalName) {
      acc[row.logicalName] = row;
    }
    return acc;
  }, {});
}

function emptySheetSnapshot_(sheetName) {
  return {
    logicalName: sheetName,
    exists: false,
    actualHeaders: [],
    dataRows: 0
  };
}

function arraysEqual_(left, right) {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}

function resolveProjectPath(relativePath) {
  return path.resolve(process.cwd(), relativePath);
}

function unirValoresUnicos_(values) {
  return [...new Set((values || []).filter(Boolean))];
}

main().catch(error => {
  console.error(error.message || error);
  process.exitCode = 1;
});
