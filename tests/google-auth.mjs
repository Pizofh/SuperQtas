import fs from 'node:fs/promises';
import { authenticate } from '@google-cloud/local-auth';
import { google } from 'googleapis';

export const DEFAULT_GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.deployments.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/userinfo.email'
];

export async function createGoogleAuthClient(config = {}) {
  const scopes = unirValoresUnicos_([
    ...DEFAULT_GOOGLE_SCOPES,
    ...((config && config.scopes) || [])
  ]);
  const authMode = normalizarAuthMode_(config && config.authMode);

  const authorizedUserRaw = await loadJsonSource_(
    config && config.authorizedUserPath,
    'GOOGLE_AUTHORIZED_USER_JSON'
  );
  if (authMode === 'authorized_user') {
    if (!authorizedUserRaw) {
      throw new Error(
        'authMode=authorized_user requiere authorizedUserPath o GOOGLE_AUTHORIZED_USER_JSON.'
      );
    }
    return createHeadlessClient_(authorizedUserRaw, scopes);
  }
  if (authorizedUserRaw) {
    return createHeadlessClient_(authorizedUserRaw, scopes);
  }

  const serviceAccountRaw = await loadJsonSource_(
    config && config.serviceAccountPath,
    'GOOGLE_SERVICE_ACCOUNT_JSON'
  );
  if (authMode === 'service_account') {
    if (!serviceAccountRaw) {
      throw new Error(
        'authMode=service_account requiere serviceAccountPath o GOOGLE_SERVICE_ACCOUNT_JSON.'
      );
    }
    return createHeadlessClient_(serviceAccountRaw, scopes);
  }
  if (serviceAccountRaw) {
    return createHeadlessClient_(serviceAccountRaw, scopes);
  }

  if (authMode === 'interactive' || authMode === 'auto') {
    if (!(config && config.credentialsPath)) {
      throw new Error('Falta credentialsPath para autenticacion interactiva.');
    }
    return authenticate({
      scopes,
      keyfilePath: config.credentialsPath
    });
  }

  throw new Error(`Modo de autenticacion no soportado: ${authMode}`);
}

function normalizarAuthMode_(value) {
  const raw = String(value || 'auto').trim().toLowerCase();
  if (!raw) return 'auto';
  if (['auto', 'interactive', 'authorized_user', 'service_account'].includes(raw)) {
    return raw;
  }
  return raw;
}

async function loadJsonSource_(filePath, envName) {
  const envValue = String(process.env[envName] || '').trim();
  if (envValue) {
    return parseJsonOrThrow_(envValue, `variable ${envName}`);
  }
  if (!filePath) return null;

  try {
    const raw = await fs.readFile(filePath, 'utf8');
    return parseJsonOrThrow_(raw, filePath);
  } catch (error) {
    if (error && error.code === 'ENOENT') return null;
    throw error;
  }
}

function parseJsonOrThrow_(raw, label) {
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error(`No pude parsear JSON desde ${label}: ${error.message}`);
  }
}

async function createHeadlessClient_(credentials, scopes) {
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes
  });
  return auth.getClient();
}

function unirValoresUnicos_(values) {
  return [...new Set((values || []).filter(Boolean))];
}
