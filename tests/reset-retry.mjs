/**
 * Reintentos seguros de QA. Solo el reset completo de la hoja de pruebas
 * admite repetir una operacion luego de que Apps Script agote su ejecucion.
 * Nunca usar para ventas, compras o escrituras no idempotentes.
 */
export function isAppsScriptExecutionTimeout(error) {
  const message = String(error?.message || '');
  return /Exceeded maximum execution time/i.test(message) &&
    /testResetEntornoQTAS/i.test(message);
}

export async function retryQaReset(operation, { attempts = 2, onRetry = () => {} } = {}) {
  if (typeof operation !== 'function') throw new TypeError('Se requiere una operacion de reset.');
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      if (!isAppsScriptExecutionTimeout(error) || attempt >= attempts) throw error;
      onRetry(attempt, error);
    }
  }
}
