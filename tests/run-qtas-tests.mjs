import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { authenticate } from '@google-cloud/local-auth';
import { google } from 'googleapis';
import { SCENARIOS } from './scenarios.mjs';

const DEFAULT_CONFIG = {
  credentialsPath: './tests/google-oauth.credentials.json',
  devMode: true,
  failureArtifactPath: './tests/.artifacts/last-failure.json',
  scopes: [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.deployments.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/userinfo.email'
  ]
};

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const selected = selectScenarios(args);

  if (args.list) {
    printScenarioList(selected);
    return;
  }

  const config = await loadConfig(args.configPath);
  const auth = await authenticate({
    scopes: config.scopes,
    keyfilePath: config.credentialsPath
  });
  const scriptClient = google.script({ version: 'v1', auth });
  const context = createScenarioContext(config, scriptClient);

  if (args.probe) {
    const report = await context.probe();
    console.log(JSON.stringify(report, null, 2));
    if (!report.ok) {
      process.exitCode = 1;
    }
    return;
  }

  const results = [];

  for (const scenario of selected) {
    const startedAt = Date.now();
    console.log(`\n[RUN] ${scenario.id} | ${scenario.title}`);

    try {
      await scenario.run(context);
      const elapsedMs = Date.now() - startedAt;
      console.log(`[PASS] ${scenario.id} (${elapsedMs} ms)`);
      results.push({ id: scenario.id, ok: true, elapsedMs });
    } catch (error) {
      const elapsedMs = Date.now() - startedAt;
      console.error(`[FAIL] ${scenario.id} (${elapsedMs} ms)`);
      console.error(`       ${error.message}`);

      const artifact = await buildFailureArtifact(context, scenario, error);
      await writeFailureArtifact(config.failureArtifactPath, artifact);
      results.push({ id: scenario.id, ok: false, elapsedMs, error: error.message });
    }
  }

  const passed = results.filter(item => item.ok).length;
  const failed = results.length - passed;
  console.log(`\nResumen: ${passed}/${results.length} escenario(s) ok, ${failed} fallo(s).`);

  if (failed > 0) {
    console.log(`Ultimo fallo guardado en ${config.failureArtifactPath}`);
    process.exitCode = 1;
  }
}

function parseArgs(argv) {
  const args = {
    configPath: null,
    list: false,
    probe: false,
    scenarios: [],
    tags: []
  };

  for (let index = 0; index < argv.length; index += 1) {
    const current = argv[index];

    if (current === '--list') {
      args.list = true;
      continue;
    }
    if (current === '--probe') {
      args.probe = true;
      continue;
    }
    if (current === '--config') {
      args.configPath = argv[index + 1] || null;
      index += 1;
      continue;
    }
    if (current === '--scenario') {
      if (argv[index + 1]) args.scenarios.push(argv[index + 1]);
      index += 1;
      continue;
    }
    if (current === '--tag') {
      if (argv[index + 1]) args.tags.push(argv[index + 1]);
      index += 1;
    }
  }

  return args;
}

async function loadConfig(configPath) {
  const target = await resolveTestsConfigPath_(configPath);
  let raw;

  try {
    raw = await fs.readFile(target, 'utf8');
  } catch (error) {
    throw new Error(
      `No pude leer ${path.relative(process.cwd(), target)}. ` +
      'Usa tests/qtas.test.config.json o tests/qtas.config.json.'
    );
  }

  const userConfig = JSON.parse(raw);
  const claspConfig = await loadClaspConfig(
    userConfig.claspProjectFile || await resolveDefaultClaspProjectFile_()
  );
  const merged = {
    ...DEFAULT_CONFIG,
    ...userConfig,
    scopes: unirValoresUnicos_([
      ...(DEFAULT_CONFIG.scopes || []),
      ...((userConfig && userConfig.scopes) || [])
    ])
  };

  merged.scriptProjectId = userConfig.scriptProjectId || userConfig.scriptId || claspConfig.scriptId || '';

  if (!merged.deploymentId) {
    if (merged.scriptId) {
      throw new Error(
        'Falta deploymentId en el archivo de configuracion de tests. Ojo: scripts.run usa el deployment ID del API executable, no el scriptId del proyecto.'
      );
    }
    throw new Error('Falta deploymentId en el archivo de configuracion de tests.');
  }

  merged.credentialsPath = resolveProjectPath(merged.credentialsPath);
  merged.failureArtifactPath = resolveProjectPath(merged.failureArtifactPath);
  return merged;
}

function selectScenarios(args) {
  let selected = SCENARIOS.slice();

  if (args.scenarios.length) {
    const wanted = new Set(args.scenarios);
    selected = selected.filter(item => wanted.has(item.id));
  }

  if (args.tags.length) {
    const wantedTags = new Set(args.tags);
    selected = selected.filter(item => (item.tags || []).some(tag => wantedTags.has(tag)));
  }

  if (!selected.length) {
    throw new Error('No quedaron escenarios seleccionados. Usa --list para ver IDs y tags.');
  }

  return selected;
}

function printScenarioList(scenarios) {
  console.log('Escenarios disponibles:\n');
  scenarios.forEach(item => {
    console.log(`- ${item.id}`);
    console.log(`  ${item.title}`);
    console.log(`  tags: ${(item.tags || []).join(', ') || 'sin tags'}`);
  });
}

function createScenarioContext(config, scriptClient) {
  const executionTargetId = resolveExecutionScriptId_(config);

  return {
    async call(functionName, ...parameters) {
      let response;
      try {
        response = await scriptClient.scripts.run({
          scriptId: executionTargetId,
          requestBody: {
            function: functionName,
            parameters,
            devMode: Boolean(config.devMode)
          }
        });
      } catch (error) {
        throw createAppsScriptTransportError(functionName, error, config);
      }

      const operation = response.data || {};
      if (operation.error) {
        throw createAppsScriptExecutionError(functionName, operation.error);
      }

      return operation.response ? operation.response.result : null;
    },

    async expectError(functionName, parameters, expectedText) {
      try {
        await this.call(functionName, ...(parameters || []));
      } catch (error) {
        const haystack = String(error.message || '').toLowerCase();
        const needle = String(expectedText || '').toLowerCase();
        if (!needle || haystack.includes(needle)) {
          return error;
        }

        throw new Error(
          `La funcion ${functionName} fallo, pero el mensaje no contiene "${expectedText}". Mensaje real: ${error.message}`
        );
      }

      throw new Error(`Se esperaba un error al ejecutar ${functionName}, pero la llamada fue exitosa.`);
    },

    async reset(options = {}) {
      return this.call('testResetEntornoQTAS', options);
    },

    async snapshot(options = {}) {
      return this.call('testSnapshotQTAS', options);
    },

    async probe() {
      const report = {
        ok: true,
        deploymentId: config.deploymentId,
        scriptProjectId: config.scriptProjectId || '',
        executionTargetId: executionTargetId,
        devMode: Boolean(config.devMode),
        checks: {}
      };

      if (config.scriptProjectId) {
        try {
          const project = await scriptClient.projects.get({
            scriptId: config.scriptProjectId
          });
          report.checks.project = {
            ok: true,
            title: project.data?.title || '',
            createTime: project.data?.createTime || '',
            updateTime: project.data?.updateTime || ''
          };
        } catch (error) {
          report.ok = false;
          report.checks.project = {
            ok: false,
            error: resumirErrorApiGoogle_(error)
          };
        }

        try {
          const deployment = await scriptClient.projects.deployments.get({
            scriptId: config.scriptProjectId,
            deploymentId: config.deploymentId
          });
          report.checks.deployment = {
            ok: true,
            deploymentId: deployment.data?.deploymentId || '',
            versionNumber: deployment.data?.deploymentConfig?.versionNumber || '',
            manifestFileName: deployment.data?.deploymentConfig?.manifestFileName || '',
            description: deployment.data?.deploymentConfig?.description || ''
          };
        } catch (error) {
          report.ok = false;
          report.checks.deployment = {
            ok: false,
            error: resumirErrorApiGoogle_(error)
          };
        }
      } else {
        report.checks.project = {
          ok: false,
          error: 'No pude inferir scriptProjectId. Revisa el archivo clasp configurado o agrega scriptProjectId en tu config de tests.'
        };
      }

      try {
        report.checks.ping = {
          ok: true,
          result: await this.call('testPingQTAS')
        };
      } catch (error) {
        report.ok = false;
        report.checks.ping = {
          ok: false,
          error: resumirErrorApiGoogle_(error)
        };
      }

      return report;
    },

    sheetRows(state, sheetName) {
      return Array.isArray(state?.sheets?.[sheetName]) ? state.sheets[sheetName] : [];
    },

    findRow(state, sheetName, predicate, message) {
      const row = this.sheetRows(state, sheetName).find(predicate);
      if (row) return row;
      throw new Error(message || `No se encontro fila en ${sheetName}.`);
    },

    assert(condition, message) {
      if (!condition) {
        throw new Error(message || 'Fallo una asercion.');
      }
    },

    equal(actual, expected, message) {
      if (actual !== expected) {
        throw new Error(`${message || 'Valores distintos.'} Esperado: ${expected}. Actual: ${actual}.`);
      }
    },

    approx(actual, expected, epsilon, message) {
      const tolerance = epsilon === undefined ? 0.01 : Math.max(0, Number(epsilon) || 0);
      if (Math.abs(this.num(actual) - this.num(expected)) > tolerance) {
        throw new Error(`${message || 'Valores fuera de rango.'} Esperado: ${expected}. Actual: ${actual}.`);
      }
    },

    num(value) {
      if (value === null || value === undefined || value === '') return 0;
      if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

      let text = String(value)
        .replace(/\s/g, '')
        .replace(/\$/g, '')
        .trim();

      if (!text) return 0;

      const hasDot = text.includes('.');
      const hasComma = text.includes(',');

      if (hasDot && hasComma) {
        if (text.lastIndexOf('.') > text.lastIndexOf(',')) {
          text = text.replace(/,/g, '');
        } else {
          text = text.replace(/\./g, '').replace(',', '.');
        }
      } else if (hasComma) {
        const parts = text.split(',');
        if (parts.length > 2 || (parts[1] && parts[1].length === 3 && parts[0].length > 0)) {
          text = text.replace(/,/g, '');
        } else {
          text = text.replace(',', '.');
        }
      } else if (hasDot) {
        const parts = text.split('.');
        if (parts.length > 2 || (parts[1] && parts[1].length === 3 && parts[0].length > 0)) {
          text = text.replace(/\./g, '');
        }
      }

      const parsed = Number(text);
      return Number.isNaN(parsed) ? 0 : parsed;
    },

    log(message) {
      console.log(`       ${message}`);
    }
  };
}

function createAppsScriptExecutionError(functionName, apiError) {
  const detail = Array.isArray(apiError.details) && apiError.details.length ? apiError.details[0] : {};
  const scriptMessage = detail.errorMessage || apiError.message || 'Error desconocido ejecutando Apps Script.';
  const stack = Array.isArray(detail.scriptStackTraceElements)
    ? detail.scriptStackTraceElements.map(item => `${item.function}:${item.lineNumber}`).join(' > ')
    : '';
  const error = new Error(stack ? `${functionName}: ${scriptMessage} | ${stack}` : `${functionName}: ${scriptMessage}`);
  error.apiError = apiError;
  return error;
}

function createAppsScriptTransportError(functionName, error, config) {
  const summary = resumirErrorApiGoogle_(error);
  const fragments = [
    `${functionName}: error llamando Apps Script API`,
    summary.status ? `HTTP ${summary.status}` : '',
    summary.message || '',
    config?.deploymentId ? `deploymentId=${config.deploymentId}` : '',
    config?.scriptProjectId ? `scriptProjectId=${config.scriptProjectId}` : ''
  ].filter(Boolean);
  const transportError = new Error(fragments.join(' | '));
  transportError.originalError = error;
  transportError.httpStatus = summary.status || null;
  transportError.responseData = summary.responseData || null;
  transportError.apiMessage = summary.message || '';
  return transportError;
}

function resumirErrorApiGoogle_(error) {
  const responseData = error?.response?.data || null;
  const payloadError = responseData?.error || {};
  return {
    status: error?.response?.status || error?.status || error?.code || null,
    message: payloadError.message || error?.message || 'Error desconocido',
    responseData
  };
}

async function buildFailureArtifact(context, scenario, error) {
  let snapshot = null;

  try {
    snapshot = await context.snapshot();
  } catch (snapshotError) {
    snapshot = {
      ok: false,
      error: snapshotError.message
    };
  }

  return {
    generatedAt: new Date().toISOString(),
    scenario: {
      id: scenario.id,
      title: scenario.title,
      tags: scenario.tags || []
    },
    error: {
      message: error.message,
      httpStatus: error.httpStatus || null,
      responseData: error.responseData || null
    },
    snapshot
  };
}

async function writeFailureArtifact(targetPath, payload) {
  const dir = path.dirname(targetPath);
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(targetPath, JSON.stringify(payload, null, 2), 'utf8');
}

function resolveProjectPath(targetPath) {
  return path.isAbsolute(targetPath)
    ? targetPath
    : path.resolve(process.cwd(), targetPath);
}

async function resolveTestsConfigPath_(configPath) {
  const candidates = configPath
    ? [configPath]
    : ['./tests/qtas.test.config.json', './tests/qtas.config.json'];

  for (const candidate of candidates) {
    const resolved = resolveProjectPath(candidate);
    try {
      await fs.access(resolved);
      return resolved;
    } catch (error) {
      // seguimos con el siguiente candidato
    }
  }

  throw new Error(
    'No encontre configuracion de tests. ' +
    'Crea tests/qtas.test.config.json a partir de tests/qtas.test.config.example.json ' +
    'o usa el archivo legado tests/qtas.config.json.'
  );
}

async function resolveDefaultClaspProjectFile_() {
  const candidates = ['./.clasp.test.json', './.clasp.json'];

  for (const candidate of candidates) {
    const resolved = resolveProjectPath(candidate);
    try {
      await fs.access(resolved);
      return candidate;
    } catch (error) {
      // seguimos con el siguiente candidato
    }
  }

  return './.clasp.json';
}

async function loadClaspConfig(targetPath) {
  const target = resolveProjectPath(targetPath || './.clasp.json');

  try {
    const raw = await fs.readFile(target, 'utf8');
    const parsed = JSON.parse(raw);
    return {
      scriptId: parsed.scriptId || ''
    };
  } catch (error) {
    return {
      scriptId: ''
    };
  }
}

function resolveExecutionScriptId_(config) {
  if (config && config.devMode && config.scriptProjectId) {
    return config.scriptProjectId;
  }
  return config ? config.deploymentId : '';
}

function unirValoresUnicos_(values) {
  return [...new Set((values || []).filter(Boolean))];
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
