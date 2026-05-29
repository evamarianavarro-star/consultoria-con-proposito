const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

const PORT = process.env.PORT || 10000;
const PASSWORD = process.env.APP_PASSWORD || 'cambiame';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

// ─── DATABASE ────────────────────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function initDB() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS consultancies (
        id TEXT PRIMARY KEY,
        nombre TEXT NOT NULL,
        empresa TEXT,
        depto TEXT,
        fecha TEXT,
        consultora TEXT,
        data JSONB NOT NULL DEFAULT '{}',
        archivos JSONB NOT NULL DEFAULT '[]',
        report TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      )
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        consultancy_id TEXT NOT NULL REFERENCES consultancies(id) ON DELETE CASCADE,
        name TEXT NOT NULL,
        mime_type TEXT,
        size INTEGER,
        data BYTEA,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);
    console.log('✅ Base de datos lista');
  } catch (e) {
    console.error('❌ Error inicializando DB:', e.message);
  }
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function sendJSON(res, status, data) {
  res.writeHead(status, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(data));
}

function checkAuth(req) {
  const auth = req.headers['x-app-password'];
  return auth === PASSWORD;
}

function readBody(req, maxBytes = 50 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let body = '';
    let bytes = 0;
    req.on('data', chunk => {
      bytes += chunk.length;
      if (bytes > maxBytes) {
        reject(new Error('Payload too large'));
        req.destroy();
        return;
      }
      body += chunk;
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

// ─── ANTHROPIC PROXY ─────────────────────────────────────────────────────────
function callAnthropic(prompt, maxTokens = 4000) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: maxTokens,
      messages: [{ role: 'user', content: prompt }]
    });
    const opts = {
      hostname: 'api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Length': Buffer.byteLength(payload)
      }
    };
    const req = https.request(opts, r => {
      let body = '';
      r.on('data', c => body += c);
      r.on('end', () => {
        try {
          const data = JSON.parse(body);
          if (data.content && data.content[0]) {
            resolve(data.content[0].text);
          } else {
            console.error('Anthropic error:', body.substring(0, 300));
            reject(new Error(data.error?.message || 'Error desconocido'));
          }
        } catch (e) {
          console.error('Parse error:', body.substring(0, 300));
          reject(e);
        }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

// ─── ROUTES ──────────────────────────────────────────────────────────────────
async function handleAPI(req, res, url) {
  if (!checkAuth(req)) return sendJSON(res, 401, { error: 'No autorizado' });

  try {
    // LIST consultancies
    if (url === '/api/consultancies' && req.method === 'GET') {
      const r = await pool.query('SELECT id, nombre, empresa, depto, fecha, consultora, updated_at FROM consultancies ORDER BY updated_at DESC');
      return sendJSON(res, 200, r.rows);
    }

    // GET one consultancy
    const getMatch = url.match(/^\/api\/consultancies\/([^/]+)$/);
    if (getMatch && req.method === 'GET') {
      const r = await pool.query('SELECT * FROM consultancies WHERE id = $1', [getMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrada' });
      return sendJSON(res, 200, r.rows[0]);
    }

    // CREATE consultancy
    if (url === '/api/consultancies' && req.method === 'POST') {
      const body = await readBody(req);
      const c = JSON.parse(body);
      const id = c.id || 'c_' + Date.now();
      await pool.query(
        `INSERT INTO consultancies (id, nombre, empresa, depto, fecha, consultora, data, archivos)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
        [id, c.nombre, c.empresa, c.depto, c.fecha, c.consultora, JSON.stringify(c.data || {}), JSON.stringify(c.archivos || [])]
      );
      return sendJSON(res, 200, { id });
    }

    // UPDATE consultancy
    if (getMatch && req.method === 'PUT') {
      const body = await readBody(req);
      const c = JSON.parse(body);
      await pool.query(
        `UPDATE consultancies SET nombre=$2, empresa=$3, depto=$4, fecha=$5, consultora=$6, data=$7, archivos=$8, report=$9, updated_at=NOW()
         WHERE id=$1`,
        [getMatch[1], c.nombre, c.empresa, c.depto, c.fecha, c.consultora, JSON.stringify(c.data || {}), JSON.stringify(c.archivos || []), c.report || null]
      );
      return sendJSON(res, 200, { ok: true });
    }

    // DELETE consultancy
    if (getMatch && req.method === 'DELETE') {
      await pool.query('DELETE FROM consultancies WHERE id = $1', [getMatch[1]]);
      return sendJSON(res, 200, { ok: true });
    }

    // UPLOAD file
    const uploadMatch = url.match(/^\/api\/consultancies\/([^/]+)\/files$/);
    if (uploadMatch && req.method === 'POST') {
      const body = await readBody(req);
      const { id, name, mimeType, size, base64 } = JSON.parse(body);
      const buf = Buffer.from(base64, 'base64');
      await pool.query(
        'INSERT INTO files (id, consultancy_id, name, mime_type, size, data) VALUES ($1,$2,$3,$4,$5,$6)',
        [id, uploadMatch[1], name, mimeType, size, buf]
      );
      return sendJSON(res, 200, { id });
    }

    // LIST files
    const listFilesMatch = url.match(/^\/api\/consultancies\/([^/]+)\/files$/);
    if (listFilesMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, name, mime_type, size, created_at FROM files WHERE consultancy_id = $1 ORDER BY created_at DESC', [listFilesMatch[1]]);
      return sendJSON(res, 200, r.rows);
    }

    // DOWNLOAD file
    const dlMatch = url.match(/^\/api\/files\/([^/]+)\/download$/);
    if (dlMatch && req.method === 'GET') {
      const r = await pool.query('SELECT name, mime_type, data FROM files WHERE id = $1', [dlMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrado' });
      const f = r.rows[0];
      res.writeHead(200, {
        'Content-Type': f.mime_type || 'application/octet-stream',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(f.name)}"`
      });
      res.end(f.data);
      return;
    }

    // DELETE file
    const delFileMatch = url.match(/^\/api\/files\/([^/]+)$/);
    if (delFileMatch && req.method === 'DELETE') {
      await pool.query('DELETE FROM files WHERE id = $1', [delFileMatch[1]]);
      return sendJSON(res, 200, { ok: true });
    }

    // GENERATE report
    if (url === '/api/generate-report' && req.method === 'POST') {
      const body = await readBody(req);
      const { prompt } = JSON.parse(body);
      const result = await callAnthropic(prompt, 4000);
      return sendJSON(res, 200, { report: result });
    }

    sendJSON(res, 404, { error: 'Ruta no encontrada' });
  } catch (err) {
    console.error('API error:', err.message);
    sendJSON(res, 500, { error: err.message });
  }
}

// ─── SERVER ──────────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const url = req.url.split('?')[0];

  if (url.startsWith('/api/')) return handleAPI(req, res, url);

  // Check auth endpoint
  if (url === '/check-auth') {
    return sendJSON(res, 200, { ok: checkAuth(req) });
  }

  // Serve index.html
  if (url === '/' || url === '/index.html') {
    fs.readFile(path.join(__dirname, 'index.html'), (err, data) => {
      if (err) {
        res.writeHead(500);
        res.end('Error');
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(data);
    });
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

initDB().then(() => {
  server.listen(PORT, () => {
    console.log(`🌱 Servidor corriendo en puerto ${PORT}`);
  });
});
