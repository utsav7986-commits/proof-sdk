import express from 'express';
import compression from 'compression';
import { createServer } from 'http';
import { WebSocketServer } from 'ws';
import path from 'path';
import { fileURLToPath } from 'url';
import { apiRoutes } from './routes.js';
import { agentRoutes } from './agent-routes.js';
import { setupWebSocket } from './ws.js';
import { createBridgeMountRouter } from './bridge.js';
import { getCollabRuntime, startCollabRuntimeEmbedded } from './collab.js';
import { discoveryRoutes } from './discovery-routes.js';
import { shareWebRoutes } from './share-web-routes.js';
import {
  capabilitiesPayload,
  enforceApiClientCompatibility,
  enforceBridgeClientCompatibility,
} from './client-capabilities.js';
import { getBuildInfo } from './build-info.js';
import { listActiveDocuments } from './db.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = Number.parseInt(process.env.PORT || '4000', 10);
const DEFAULT_ALLOWED_CORS_ORIGINS = [
  'http://localhost:3000',
  'http://127.0.0.1:3000',
  'http://localhost:4000',
  'http://127.0.0.1:4000',
  'null',
];

function parseAllowedCorsOrigins(): Set<string> {
  const configured = (process.env.PROOF_CORS_ALLOW_ORIGINS || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean);
  return new Set(configured.length > 0 ? configured : DEFAULT_ALLOWED_CORS_ORIGINS);
}

async function main(): Promise<void> {
  const app = express();
  app.use(compression());
  const server = createServer(app);
  const wss = new WebSocketServer({ server, path: '/ws' });
  wss.on('error', (error) => {
    console.error('[server] WebSocketServer error (non-fatal):', error);
  });
  const allowedCorsOrigins = parseAllowedCorsOrigins();

  app.use(express.json({ limit: '10mb' }));
  app.use(express.static(path.join(__dirname, '..', 'public')));
  app.use('/assets', express.static(path.join(__dirname, '..', 'dist', 'assets')));

  app.use((req, res, next) => {
    const originHeader = req.header('origin');
    if (originHeader && allowedCorsOrigins.has(originHeader)) {
      res.setHeader('Access-Control-Allow-Origin', originHeader);
      res.setHeader('Vary', 'Origin');
    }
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,PATCH,DELETE,OPTIONS');
    res.setHeader(
      'Access-Control-Allow-Headers',
      [
        'Content-Type',
        'Authorization',
        'X-Proof-Client-Version',
        'X-Proof-Client-Build',
        'X-Proof-Client-Protocol',
        'x-share-token',
        'x-bridge-token',
        'x-auth-poll-token',
        'X-Agent-Id',
        'X-Window-Id',
        'X-Document-Id',
        'Idempotency-Key',
        'X-Idempotency-Key',
      ].join(', '),
    );
    if (req.method === 'OPTIONS') {
      res.status(204).end();
      return;
    }
    next();
  });

  app.get('/', (_req, res) => {
    let docs: Array<{ slug: string; title: string; updated_at: string }> = [];
    try {
      const rows = listActiveDocuments();
      docs = rows.map((r) => ({
        slug: r.slug,
        title: (r.title as string | null) || 'Untitled',
        updated_at: (r.updated_at as string | null) || '',
      })).sort((a, b) => b.updated_at.localeCompare(a.updated_at));
    } catch {}

    const docRows = docs.map((d) => `
      <tr>
        <td><a href="/d/${d.slug}">${d.title}</a></td>
        <td>${d.updated_at ? new Date(d.updated_at).toLocaleString() : '—'}</td>
        <td><a href="/d/${d.slug}">Open</a></td>
      </tr>`).join('');

    res.type('html').send(`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>My Docs</title>
    <style>
      body { font-family: ui-sans-serif, system-ui, sans-serif; margin: 0; padding: 48px 24px; color: #17261d; background: #f7faf5; }
      main { max-width: 860px; margin: 0 auto; }
      h1 { font-size: 2rem; margin: 0 0 1.5rem; }
      button { background: #266854; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 1rem; cursor: pointer; margin-bottom: 2rem; }
      button:hover { background: #1a4a3a; }
      table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
      th { text-align: left; padding: 12px 16px; background: #eaf2e6; font-size: 0.85rem; color: #4a6a5a; }
      td { padding: 12px 16px; border-top: 1px solid #eaf2e6; }
      a { color: #266854; text-decoration: none; }
      a:hover { text-decoration: underline; }
      .empty { text-align: center; padding: 48px; color: #8aaa9a; }
    </style>
  </head>
  <body>
    <main>
      <h1>My Docs</h1>
      <button onclick="createDoc()">+ New Document</button>
      <table>
        <thead><tr><th>Title</th><th>Last Edited</th><th></th></tr></thead>
        <tbody>${docRows || '<tr><td colspan="3" class="empty">No documents yet. Create your first one!</td></tr>'}</tbody>
      </table>
    </main>
    <script>
      async function createDoc() {
        const title = prompt('Document title:', 'Untitled');
        if (!title) return;
        const res = await fetch('/documents', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title, markdown: '# ' + title + '\\n\\n' })
        });
        const data = await res.json();
        if (data.url) window.location.href = data.tokenUrl || data.url;
      }
    </script>
  </body>
</html>`);
  });

  app.get('/health', (_req, res) => {
    const buildInfo = getBuildInfo();
    res.json({
      ok: true,
      buildInfo,
      collab: getCollabRuntime(),
    });
  });

  app.get('/api/capabilities', (_req, res) => {
    res.json(capabilitiesPayload());
  });

  app.use(discoveryRoutes);
  app.use('/api', enforceApiClientCompatibility, apiRoutes);
  app.use('/api/agent', agentRoutes);
  app.use(apiRoutes);
  app.use('/d', createBridgeMountRouter(enforceBridgeClientCompatibility));
  app.use('/documents', createBridgeMountRouter(enforceBridgeClientCompatibility));
  app.use('/documents', agentRoutes);
  app.use(shareWebRoutes);

  setupWebSocket(wss);
  await startCollabRuntimeEmbedded(PORT);

  server.listen(PORT, () => {
    console.log(`[proof-sdk] listening on http://127.0.0.1:${PORT}`);
  });
}

main().catch((error) => {
  console.error('[proof-sdk] failed to start server', error);
  process.exit(1);
});
