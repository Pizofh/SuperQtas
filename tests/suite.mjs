/**
 * Entrada unica de pruebas QTAS. Mantiene casos unitarios y E2E modularizados,
 * pero la seleccion y los umbrales de cobertura viven en esta suite.
 *
 * Ejemplos:
 *   npm run test:suite -- --level unit
 *   npm run test:suite -- --level smoke --config tests/qtas.test.config.json
 *   npm run test:suite -- --level full --config tests/qtas.test.config.json
 *   npm run test:suite -- --level diagnose --scenario venta_con_deuda --config tests/qtas.test.config.json
 */
import { readdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import process from 'node:process';
import { SCENARIOS } from './scenarios.mjs';

const LEVELS = new Set(['unit', 'probe', 'smoke', 'full', 'diagnose', 'list']);
const MINIMUM = { smoke: 10, full: 24 };

function parse(argv) {
  const args = { level: 'unit', config: '', scenario: '', failFast: true };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--level' || a === '--config' || a === '--scenario') {
      const value = argv[++i];
      if (!value || value.startsWith('--')) throw new Error(`Falta valor para ${a}`);
      args[a.slice(2)] = value;
    } else if (a === '--no-fail-fast') {
      args.failFast = false;
    } else {
      throw new Error(`Argumento no soportado: ${a}`);
    }
  }
  if (!LEVELS.has(args.level)) throw new Error(`Nivel desconocido: ${args.level}`);
  if (args.level === 'diagnose' && !args.scenario) {
    throw new Error('Para diagnosticar indica --scenario con un ID de la misma suite.');
  }
  if (args.scenario && args.level !== 'diagnose') {
    throw new Error('--scenario solo se permite con --level diagnose.');
  }
  return args;
}

function assertCoverage(level) {
  if (!MINIMUM[level]) return;
  const ids = SCENARIOS.map(item => item.id);
  if (new Set(ids).size !== ids.length) throw new Error('IDs de escenario duplicados.');
  const selected = level === 'smoke'
    ? SCENARIOS.filter(item => (item.tags || []).includes('smoke'))
    : SCENARIOS;
  if (selected.length < MINIMUM[level]) {
    throw new Error(`Cobertura insuficiente: ${selected.length} escenarios ${level}; minimo ${MINIMUM[level]}.`);
  }
  console.log(`[COVERAGE] ${selected.length} escenarios ${level} (minimo ${MINIMUM[level]}).`);
}

function run(label, args) {
  console.log(`\n[SUITE] ${label} | ${new Date().toISOString()}`);
  const start = Date.now();
  const result = spawnSync(process.execPath, args, {
    stdio: 'inherit',
    env: process.env
  });
  if (result.error) throw result.error;
  console.log(`[SUITE] ${label}: ${Math.round((Date.now() - start) / 1000)}s | exit ${result.status}`);
  if (result.signal) throw new Error(`Interrumpido por ${result.signal}`);
  if (result.status !== 0) process.exitCode = result.status || 1;
}

try {
  const options = parse(process.argv.slice(2));
  if (options.level === 'unit') {
    const files = readdirSync(new URL('.', import.meta.url))
      .filter(file => file.endsWith('.test.mjs'))
      .sort().map(file => `tests/${file}`);
    if (!files.length) throw new Error('No se encontraron pruebas unitarias.');
    console.log(`[COVERAGE] ${files.length} modulos unitarios detectados automaticamente.`);
    run('Unitarias', ['--test', ...files]);
  } else {
    assertCoverage(options.level);
    const args = ['tests/run-qtas-tests.mjs'];
    if (options.config) args.push('--config', options.config);
    if (options.level === 'smoke') args.push('--tag', 'smoke');
    if (options.level === 'probe') args.push('--probe');
    if (options.level === 'list') args.push('--list');
    if (options.level === 'diagnose') args.push('--scenario', options.scenario);
    if (options.failFast && !['list', 'probe'].includes(options.level)) args.push('--fail-fast');
    run(`E2E ${options.level}`, args);
  }
} catch (error) {
  console.error('[SUITE ERROR]', error.message || error);
  process.exitCode = 1;
}
