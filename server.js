const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

let pdfParse, mammoth, HTMLtoDOCX;
try { pdfParse = require('pdf-parse'); } catch(e) { console.log('pdf-parse no disponible'); }
try { mammoth = require('mammoth'); } catch(e) { console.log('mammoth no disponible'); }
try { HTMLtoDOCX = require('html-to-docx'); } catch(e) { console.log('html-to-docx no disponible'); }

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
    await pool.query(`ALTER TABLE files ADD COLUMN IF NOT EXISTS description TEXT`);
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

function escapeHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Aplica estilos inline a las etiquetas del informe para que Word los respete bien
function transformReportHtml(html) {
  if (!html) return '';
  let out = html;

  // 1) Limpieza radical: eliminar contenido problemático

  // Caracteres de control inválidos en XML (excepto tab, LF y CR)
  out = out.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');

  // Comentarios y CDATA
  out = out.replace(/<!--[\s\S]*?-->/g, '');
  out = out.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '');

  // Etiquetas no soportadas
  out = out.replace(/<(style|script|link|meta|head|title)[\s\S]*?<\/\1>/gi, '');
  out = out.replace(/<(style|script|link|meta)[^>]*\/?>/gi, '');

  // 2) ELIMINAR TODOS LOS ATRIBUTOS de las etiquetas permitidas
  // Solo mantenemos las etiquetas estructurales, sin atributos
  const allowedTags = ['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'strong', 'b', 'em', 'i', 'br', 'span', 'table', 'tr', 'td', 'th', 'tbody', 'thead'];

  // Reemplaza cada etiqueta de apertura para quitarle TODOS los atributos
  out = out.replace(/<([a-zA-Z][a-zA-Z0-9]*)\s[^>]*>/g, function(match, tagName) {
    const tag = tagName.toLowerCase();
    if (allowedTags.indexOf(tag) === -1) return '';
    return '<' + tag + '>';
  });

  // Elimina etiquetas de cierre desconocidas
  out = out.replace(/<\/([a-zA-Z][a-zA-Z0-9]*)>/g, function(match, tagName) {
    const tag = tagName.toLowerCase();
    if (allowedTags.indexOf(tag) === -1) return '';
    return '</' + tag + '>';
  });

  // 3) Ahora añadimos estilos inline solo donde queremos

  out = out.replace(/<h1>/g, '<h1 style="font-family: Calibri, sans-serif; font-size:20pt; font-weight:600; color:#1A1814; margin:28pt 0 12pt 0;">');
  out = out.replace(/<h2>/g, '<h2 style="font-family: Calibri, sans-serif; font-size:16pt; font-weight:600; color:#1A1814; margin:24pt 0 10pt 0;">');
  out = out.replace(/<h3>/g, '<h3 style="font-family: Calibri, sans-serif; font-size:9pt; font-weight:600; color:#6B6560; text-transform:uppercase; letter-spacing:1.5pt; margin:16pt 0 6pt 0;">');
  out = out.replace(/<h4>/g, '<h4 style="font-family: Calibri, sans-serif; font-size:11pt; font-weight:600; color:#1A1814; margin:14pt 0 6pt 0;">');
  out = out.replace(/<p>/g, '<p style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.6; margin:0 0 10pt 0;">');
  out = out.replace(/<ul>/g, '<ul style="margin:6pt 0 12pt 0; padding-left:20pt;">');
  out = out.replace(/<ol>/g, '<ol style="margin:6pt 0 12pt 0; padding-left:20pt;">');
  out = out.replace(/<li>/g, '<li style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.55; margin-bottom:4pt;">');
  out = out.replace(/<strong>/g, '<strong style="font-weight:600;">');
  out = out.replace(/<b>/g, '<b style="font-weight:600;">');

  return out;
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
      const { id, name, mimeType, size, base64, description } = JSON.parse(body);
      const buf = Buffer.from(base64, 'base64');
      const extracted = await extractText(buf, mimeType, name);
      await pool.query(
        'INSERT INTO files (id, consultancy_id, name, mime_type, size, data, extracted_text, description) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)',
        [id, uploadMatch[1], name, mimeType, size, buf, extracted, description || null]
      );
      return sendJSON(res, 200, { id, extracted: !!extracted, textLength: extracted ? extracted.length : 0 });
    }

    // LIST files
    if (uploadMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, name, mime_type, size, description, extracted_text IS NOT NULL AS has_text, created_at FROM files WHERE consultancy_id = $1 ORDER BY created_at DESC', [uploadMatch[1]]);
      return sendJSON(res, 200, r.rows);
    }

    // UPDATE file description
    const fileUpdateMatch = url.match(/^\/api\/files\/([^/]+)\/description$/);
    if (fileUpdateMatch && req.method === 'PUT') {
      const body = await readBody(req);
      const { description } = JSON.parse(body);
      await pool.query('UPDATE files SET description = $1 WHERE id = $2', [description || null, fileUpdateMatch[1]]);
      return sendJSON(res, 200, { ok: true });
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
          'SELECT id, name, mime_type, data, extracted_text, description FROM files WHERE id = ANY($1::text[])',
          [fileIds]
        );
        let textFilesContent = '\n\n=== CONTENIDO DE ARCHIVOS ADJUNTOS ===\n';
        let hasTextFiles = false;
        for (const f of filesResult.rows) {
          const descLine = f.description ? `(Descripción: ${f.description})` : '';
          // Imágenes: añadir como bloque image
          if (f.mime_type && f.mime_type.startsWith('image/')) {
            const supported = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
            if (supported.includes(f.mime_type)) {
              contentBlocks.push({
                type: 'image',
                source: { type: 'base64', media_type: f.mime_type, data: f.data.toString('base64') }
              });
              contentBlocks.push({ type: 'text', text: `[Imagen anterior: ${f.name}] ${descLine}` });
            }
          } else if (f.extracted_text) {
            hasTextFiles = true;
            textFilesContent += `\n--- ARCHIVO: ${f.name} ${descLine} ---\n${f.extracted_text}\n`;
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

    // ADMIN — stats
    if (url === '/api/admin/stats' && req.method === 'GET') {
      const stats = {};
      const c = await pool.query('SELECT COUNT(*) AS n FROM consultancies');
      stats.consultancies = parseInt(c.rows[0].n);
      const f = await pool.query('SELECT COUNT(*) AS n, COALESCE(SUM(size),0) AS total_size FROM files');
      stats.files = parseInt(f.rows[0].n);
      stats.totalSize = parseInt(f.rows[0].total_size);
      const r = await pool.query('SELECT COUNT(*) AS n FROM reports');
      stats.reports = parseInt(r.rows[0].n);
      const fe = await pool.query('SELECT COUNT(*) AS n FROM files WHERE extracted_text IS NOT NULL');
      stats.filesWithText = parseInt(fe.rows[0].n);
      // DB info
      try {
        const sz = await pool.query("SELECT pg_size_pretty(pg_database_size(current_database())) AS size");
        stats.dbSize = sz.rows[0].size;
      } catch(e) { stats.dbSize = 'n/d'; }
      return sendJSON(res, 200, stats);
    }

    // ADMIN — all files
    if (url === '/api/admin/files' && req.method === 'GET') {
      const r = await pool.query(`
        SELECT f.id, f.name, f.mime_type, f.size, f.extracted_text IS NOT NULL AS has_text,
               f.created_at, f.consultancy_id, c.nombre AS consultancy_name
        FROM files f LEFT JOIN consultancies c ON c.id = f.consultancy_id
        ORDER BY f.created_at DESC
      `);
      return sendJSON(res, 200, r.rows);
    }

    // ADMIN — all reports
    if (url === '/api/admin/reports' && req.method === 'GET') {
      const r = await pool.query(`
        SELECT r.id, r.title, r.report_type, r.created_at, r.consultancy_id,
               c.nombre AS consultancy_name, LENGTH(r.content) AS content_size
        FROM reports r LEFT JOIN consultancies c ON c.id = r.consultancy_id
        ORDER BY r.created_at DESC
      `);
      return sendJSON(res, 200, r.rows);
    }

    // ADMIN — health
    if (url === '/api/admin/health' && req.method === 'GET') {
      const health = { database: 'ok', anthropic: 'unknown', extraction: { pdf: !!pdfParse, docx: !!mammoth } };
      try {
        await pool.query('SELECT 1');
      } catch (e) { health.database = 'error: ' + e.message; }
      health.anthropic = ANTHROPIC_API_KEY ? 'configured' : 'missing';
      return sendJSON(res, 200, health);
    }

    // DOWNLOAD report as Word
    const reportDocxMatch = url.match(/^\/api\/reports\/([^/]+)\/docx$/);
    if (reportDocxMatch && req.method === 'GET') {
      if (!HTMLtoDOCX) return sendJSON(res, 500, { error: 'html-to-docx no disponible' });
      const r = await pool.query(`
        SELECT r.title, r.report_type, r.content, r.created_at,
               c.nombre AS consultancy_name, c.empresa, c.depto, c.fecha, c.consultora
        FROM reports r LEFT JOIN consultancies c ON c.id = r.consultancy_id
        WHERE r.id = $1
      `, [reportDocxMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrado' });
      const rep = r.rows[0];

      const typeLabels = {
        diagnostico_completo: 'Diagnóstico completo',
        resumen_ejecutivo: 'Resumen ejecutivo',
        plan_accion: 'Plan de acción',
        devolucion_equipo: 'Devolución al equipo',
        tecnico_procesos: 'Análisis de procesos',
        analisis_herramientas: 'Análisis de herramientas'
      };
      const typeLabel = typeLabels[rep.report_type] || 'Informe';

      // HTML simple y robusto para evitar problemas de XML
      const metaItems = [];
      metaItems.push(`<strong>Cliente:</strong> ${escapeHtml(rep.empresa || rep.consultancy_name || '—')}`);
      if (rep.depto) metaItems.push(`<strong>Departamento:</strong> ${escapeHtml(rep.depto)}`);
      metaItems.push(`<strong>Fecha:</strong> ${escapeHtml(rep.fecha || new Date(rep.created_at).toLocaleDateString('es-ES'))}`);
      if (rep.consultora) metaItems.push(`<strong>Consultora:</strong> ${escapeHtml(rep.consultora)}`);

      const fullHtml = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body>
<p style="font-family: Calibri, sans-serif; font-size:9pt; letter-spacing:2pt; color:#6B6560; margin:0 0 16pt 0;">${escapeHtml(typeLabel.toUpperCase())}</p>
<h1 style="font-family: Calibri, sans-serif; font-size:24pt; font-weight:300; color:#1A1814; margin:0 0 20pt 0;">${escapeHtml(rep.title)}</h1>
<p style="font-family: Calibri, sans-serif; font-size:10pt; color:#6B6560; margin:0 0 28pt 0; line-height:1.8;">${metaItems.join(' &nbsp;·&nbsp; ')}</p>
${transformReportHtml(rep.content)}
</body></html>`;

      let buffer;
      try {
        buffer = await HTMLtoDOCX(fullHtml, null, {
          orientation: 'portrait',
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          font: 'Calibri',
          fontSize: 22,
          title: rep.title,
          creator: rep.consultora || 'Consultoría con Propósito'
        });
      } catch (genErr) {
        console.error('Error generando docx:', genErr.message);
        return sendJSON(res, 500, { error: 'No se pudo generar el Word: ' + genErr.message + '. Prueba a regenerar el informe.' });
      }

      const fileName = (rep.title || 'informe').replace(/[^a-z0-9áéíóúñ\s\-_]/gi, '').replace(/\s+/g, '_') + '.docx';
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(fileName)}"`,
        'Content-Length': buffer.length
      });
      res.end(buffer);
      return;
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
