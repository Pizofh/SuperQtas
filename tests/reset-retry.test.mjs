import assert from 'node:assert/strict';
import test from 'node:test';
import { isAppsScriptExecutionTimeout, retryQaReset } from './reset-retry.mjs';

test('reintenta una unica vez el reset de QA tras timeout del script', async () => {
  let calls = 0;
  let retried = 0;
  const result = await retryQaReset(async () => {
    calls += 1;
    if (calls === 1) throw new Error('testResetEntornoQTAS: Exceeded maximum execution time');
    return { ok: true, reset: true };
  }, { onRetry: () => { retried += 1; } });
  assert.deepEqual(result, { ok: true, reset: true });
  assert.equal(calls, 2);
  assert.equal(retried, 1);
});

test('falla tras dos timeouts del reset y no reintenta indefinidamente', async () => {
  let calls = 0;
  await assert.rejects(
    retryQaReset(async () => {
      calls += 1;
      throw new Error('testResetEntornoQTAS: Exceeded maximum execution time');
    }),
    /Exceeded maximum execution time/
  );
  assert.equal(calls, 2);
});

test('errores funcionales del reset no se reintentan', async () => {
  let calls = 0;
  await assert.rejects(
    retryQaReset(async () => {
      calls += 1;
      throw new Error('testResetEntornoQTAS: Hoja de QA invalida');
    }),
    /Hoja de QA invalida/
  );
  assert.equal(calls, 1);
});

test('timeout de compra no es elegible para reintento automatico', () => {
  assert.equal(isAppsScriptExecutionTimeout(
    new Error('registrarCompraQTAS: Exceeded maximum execution time')
  ), false);
});

test('timeout de venta no es elegible para reintento automatico', () => {
  assert.equal(isAppsScriptExecutionTimeout(
    new Error('registrarVentaQTAS: Exceeded maximum execution time')
  ), false);
});
