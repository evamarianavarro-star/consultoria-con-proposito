const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

let pdfParse, mammoth;
try { pdfParse = require('pdf-parse'); } catch(e) { console.log('pdf-parse no disponible'); }
try { mammoth = require('mammoth'); } catch(e) { console.log('mammoth no disponible'); }

const PORT = process.env.PORT || 10000;
const PASSWORD = process.env.APP_PASSWORD || 'cambiame';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

// ─── DATABASE ───────────────────────────────────────────────────────────────
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
        extracted_text TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS reports (
        id TEXT PRIMARY KEY,
        consultancy_id TEXT NOT NULL REFERENCES consultancies(id) ON DELETE CASCADE,
        title TEXT NOT NULL,
        report_type TEXT NOT NULL,
        content TEXT NOT NULL,
        files_used JSONB DEFAULT '[]',
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);
    // Migration: add extracted_text column if missing
    await pool.query(`ALTER TABLE files ADD COLUMN IF NOT EXISTS extracted_text TEXT`);
    console.log('✅ Base de datos lista');
  } catch (e) {
    console.error('❌ Error inicializando DB:', e.message);
  }
}

// ─── HELPERS ────────────────────────────────────────────────────────────────
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

// ─── EXTRACT TEXT ───────────────────────────────────────────────────────────
async function extractText(buffer, mimeType, fileName) {
  try {
    const name = (fileName || '').toLowerCase();
    if (mimeType === 'application/pdf' || name.endsWith('.pdf')) {
      if (!pdfParse) return null;
      const data = await pdfParse(buffer);
      return (data.text || '').trim().substring(0, 50000);
    }
    if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || name.endsWith('.docx')) {
      if (!mammoth) return null;
      const result = await mammoth.extractRawText({ buffer });
      return (result.value || '').trim().substring(0, 50000);
    }
    if (mimeType?.startsWith('text/') || name.endsWith('.txt') || name.endsWith('.md') || name.endsWith('.csv')) {
      return buffer.toString('utf-8').trim().substring(0, 50000);
    }
    return null;
  } catch (e) {
    console.error('Error extrayendo texto:', e.message);
    return null;
  }
}

// ─── ANTHROPIC CALL ─────────────────────────────────────────────────────────
function callAnthropic(messages, maxTokens = 4000) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: maxTokens,
      messages: messages
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
            console.error('Anthropic error:', body.substring(0, 500));
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

// ─── ROUTES ─────────────────────────────────────────────────────────────────
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
      const extracted = await extractText(buf, mimeType, name);
      await pool.query(
        'INSERT INTO files (id, consultancy_id, name, mime_type, size, data, extracted_text) VALUES ($1,$2,$3,$4,$5,$6,$7)',
        [id, uploadMatch[1], name, mimeType, size, buf, extracted]
      );
      return sendJSON(res, 200, { id, extracted: !!extracted, textLength: extracted ? extracted.length : 0 });
    }

    // LIST files
    if (uploadMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, name, mime_type, size, extracted_text IS NOT NULL AS has_text, created_at FROM files WHERE consultancy_id = $1 ORDER BY created_at DESC', [uploadMatch[1]]);
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
        'Content-Disposition': `inline; filename="${encodeURIComponent(f.name)}"`
      });
      res.end(f.data);
      return;
    }

    // GET file metadata (for image base64 in reports)
    const fileMetaMatch = url.match(/^\/api\/files\/([^/]+)\/info$/);
    if (fileMetaMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, name, mime_type, size, extracted_text FROM files WHERE id = $1', [fileMetaMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrado' });
      return sendJSON(res, 200, r.rows[0]);
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
      const { consultancyId, prompt, fileIds, reportType, title } = JSON.parse(body);

      // Build messages with file content
      const contentBlocks = [{ type: 'text', text: prompt }];

      if (Array.isArray(fileIds) && fileIds.length > 0) {
        const filesResult = await pool.query(
          'SELECT id, name, mime_type, data, extracted_text FROM files WHERE id = ANY($1::text[])',
          [fileIds]
        );
        let textFilesContent = '\n\n=== CONTENIDO DE ARCHIVOS ADJUNTOS ===\n';
        let hasTextFiles = false;
        for (const f of filesResult.rows) {
          // Imágenes: añadir como bloque image
          if (f.mime_type && f.mime_type.startsWith('image/')) {
            const supported = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
            if (supported.includes(f.mime_type)) {
              contentBlocks.push({
                type: 'image',
                source: { type: 'base64', media_type: f.mime_type, data: f.data.toString('base64') }
              });
              contentBlocks.push({ type: 'text', text: `[Imagen anterior: ${f.name}]` });
            }
          } else if (f.extracted_text) {
            hasTextFiles = true;
            textFilesContent += `\n--- ARCHIVO: ${f.name} ---\n${f.extracted_text}\n`;
          }
        }
        if (hasTextFiles) {
          contentBlocks.push({ type: 'text', text: textFilesContent });
        }
      }

      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 4000);

      // Clean code fences
      let clean = (result || '').replace(/^\s*```(?:html)?\s*\n?/i, '').replace(/\n?```\s*$/i, '').trim();

      // Save to history
      const reportId = 'r_' + Date.now();
      const finalTitle = title || (reportType + ' — ' + new Date().toLocaleString('es-ES'));
      await pool.query(
        'INSERT INTO reports (id, consultancy_id, title, report_type, content, files_used) VALUES ($1,$2,$3,$4,$5,$6)',
        [reportId, consultancyId, finalTitle, reportType || 'diagnostico_completo', clean, JSON.stringify(fileIds || [])]
      );

      return sendJSON(res, 200, { report: clean, reportId, title: finalTitle });
    }

    // LIST reports for a consultancy
    const reportsListMatch = url.match(/^\/api\/consultancies\/([^/]+)\/reports$/);
    if (reportsListMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, title, report_type, files_used, created_at FROM reports WHERE consultancy_id = $1 ORDER BY created_at DESC', [reportsListMatch[1]]);
      return sendJSON(res, 200, r.rows);
    }

    // GET one report
    const reportGetMatch = url.match(/^\/api\/reports\/([^/]+)$/);
    if (reportGetMatch && req.method === 'GET') {
      const r = await pool.query('SELECT * FROM reports WHERE id = $1', [reportGetMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrado' });
      return sendJSON(res, 200, r.rows[0]);
    }

    // DELETE report
    if (reportGetMatch && req.method === 'DELETE') {
      await pool.query('DELETE FROM reports WHERE id = $1', [reportGetMatch[1]]);
      return sendJSON(res, 200, { ok: true });
    }

    sendJSON(res, 404, { error: 'Ruta no encontrada' });
  } catch (err) {
    console.error('API error:', err.message);
    sendJSON(res, 500, { error: err.message });
  }
}

// ─── SERVER ─────────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const url = req.url.split('?')[0];

  if (url.startsWith('/api/')) return handleAPI(req, res, url);

  if (url === '/check-auth') {
    return sendJSON(res, 200, { ok: checkAuth(req) });
  }

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
