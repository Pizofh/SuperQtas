import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { authenticate } from '@google-cloud/local-auth';
import { DEFAULT_GOOGLE_SCOPES } from './google-auth.mjs';

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const credentialsPath = resolveProjectPath(
    args.credentialsPath || './tests/google-oauth.credentials.json'
  );
  const outputPath = resolveProjectPath(
    args.outputPath || './tests/google-oauth.authorized-user.json'
  );
  const scopes = unirValoresUnicos_([
    ...DEFAULT_GOOGLE_SCOPES,
    ...(args.scopes || [])
  ]);

  const auth = await authenticate({
    scopes,
    keyfilePath: credentialsPath
  });
  const oauthClientJson = JSON.parse(await fs.readFile(credentialsPath, 'utf8'));
  const clientInfo = oauthClientJson.installed || oauthClientJson.web || {};
  const refreshToken = String(auth?.credentials?.refresh_token || '').trim();

  if (!clientInfo.client_id || !clientInfo.client_secret) {
    throw new Error(
      'El archivo de credenciales OAuth no trae client_id/client_secret validos.'
    );
  }
  if (!refreshToken) {
    throw new Error(
      'No se obtuvo refresh_token. Repite el bootstrap con una credencial Desktop OAuth nueva o revocando el consentimiento anterior.'
    );
  }

  const authorizedUser = {
    type: 'authorized_user',
    client_id: clientInfo.client_id,
    client_secret: clientInfo.client_secret,
    refresh_token: refreshToken
  };

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, `${JSON.stringify(authorizedUser, null, 2)}\n`, 'utf8');

  console.log(JSON.stringify({
    ok: true,
    outputPath: path.relative(process.cwd(), outputPath),
    authMode: 'authorized_user',
    scopes
  }, null, 2));
}

function parseArgs(argv) {
  const args = {
    credentialsPath: '',
    outputPath: '',
    scopes: []
  };

  for (let index = 0; index < argv.length; index += 1) {
    const current = argv[index];

    if (current === '--credentials') {
      args.credentialsPath = argv[index + 1] || '';
      index += 1;
      continue;
    }
    if (current === '--out') {
      args.outputPath = argv[index + 1] || '';
      index += 1;
      continue;
    }
    if (current === '--scope') {
      if (argv[index + 1]) args.scopes.push(argv[index + 1]);
      index += 1;
    }
  }

  return args;
}

function resolveProjectPath(targetPath) {
  return path.resolve(process.cwd(), targetPath);
}

function unirValoresUnicos_(values) {
  return [...new Set((values || []).filter(Boolean))];
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
