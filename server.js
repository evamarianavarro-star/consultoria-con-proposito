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
    // Migration: add report metadata for regeneration
    await pool.query(`ALTER TABLE reports ADD COLUMN IF NOT EXISTS custom_instructions TEXT`);
    await pool.query(`ALTER TABLE reports ADD COLUMN IF NOT EXISTS target_persona_id TEXT`);
    await pool.query(`ALTER TABLE reports ADD COLUMN IF NOT EXISTS previous_report_id TEXT`);
    await pool.query(`ALTER TABLE reports ADD COLUMN IF NOT EXISTS reference_report_ids JSONB DEFAULT '[]'`);

    // Migration: post-consultancy phase (5A)
    await pool.query(`ALTER TABLE consultancies ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'active'`);
    await pool.query(`ALTER TABLE consultancies ADD COLUMN IF NOT EXISTS closed_at TIMESTAMPTZ`);
    await pool.query(`ALTER TABLE consultancies ADD COLUMN IF NOT EXISTS official_report_ids JSONB DEFAULT '[]'`);

    await pool.query(`CREATE TABLE IF NOT EXISTS projects (
      id TEXT PRIMARY KEY,
      consultancy_id TEXT REFERENCES consultancies(id) ON DELETE CASCADE,
      title TEXT NOT NULL,
      description TEXT,
      status TEXT DEFAULT 'pending',
      priority TEXT DEFAULT 'medium',
      responsible_persona_id TEXT,
      target_date DATE,
      progress INT DEFAULT 0,
      notes TEXT,
      sort_order INT DEFAULT 0,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_projects_consultancy ON projects(consultancy_id)`);
    // Migration 5B: tasks + interactions + progress mode
    await pool.query(`ALTER TABLE projects ADD COLUMN IF NOT EXISTS tasks JSONB DEFAULT '[]'`);
    await pool.query(`ALTER TABLE projects ADD COLUMN IF NOT EXISTS interactions JSONB DEFAULT '[]'`);
    await pool.query(`ALTER TABLE projects ADD COLUMN IF NOT EXISTS progress_manual BOOLEAN DEFAULT false`);
    // Migration 5B+: link files to project/task
    await pool.query(`ALTER TABLE files ADD COLUMN IF NOT EXISTS linked_project_id TEXT`);
    await pool.query(`ALTER TABLE files ADD COLUMN IF NOT EXISTS linked_task_id TEXT`);
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
function callAnthropic(messages, maxTokens = 16000) {
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
      const { id, name, mimeType, size, base64, description, linkedProjectId, linkedTaskId } = JSON.parse(body);
      const buf = Buffer.from(base64, 'base64');
      const extracted = await extractText(buf, mimeType, name);
      await pool.query(
        'INSERT INTO files (id, consultancy_id, name, mime_type, size, data, extracted_text, description, linked_project_id, linked_task_id) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
        [id, uploadMatch[1], name, mimeType, size, buf, extracted, description || null, linkedProjectId || null, linkedTaskId || null]
      );
      return sendJSON(res, 200, { id, extracted: !!extracted, textLength: extracted ? extracted.length : 0 });
    }

    // LIST files
    if (uploadMatch && req.method === 'GET') {
      const r = await pool.query('SELECT id, name, mime_type, size, description, extracted_text IS NOT NULL AS has_text, linked_project_id, linked_task_id, created_at FROM files WHERE consultancy_id = $1 ORDER BY created_at DESC', [uploadMatch[1]]);
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
      const parsed = JSON.parse(body);
      const { consultancyId, prompt, fileIds, reportType, title, previousReportId, customInstructions, targetPersonaId } = parsed;

      // Build messages with file content
      const contentBlocks = [{ type: 'text', text: prompt }];

      // Si hay informe previo, lo añadimos como contexto
      if (previousReportId) {
        try {
          const prevR = await pool.query('SELECT title, content, created_at FROM reports WHERE id = $1', [previousReportId]);
          if (prevR.rows.length > 0) {
            const prev = prevR.rows[0];
            const prevDate = new Date(prev.created_at).toLocaleString('es-ES');
            contentBlocks.push({
              type: 'text',
              text: '\n\n=== INFORME ANTERIOR (versión previa generada el ' + prevDate + ') ===\n' +
                'Este es el informe que generaste en una sesión anterior con menos datos. Tu tarea ahora es ACTUALIZARLO y MATIZARLO con la nueva información disponible (datos del formulario y archivos), NO rehacerlo desde cero. Mantén las conclusiones que sigan siendo válidas, añade lo nuevo que hayas aprendido, ajusta lo que ahora veas diferente, y al inicio del informe incluye un pequeño párrafo destacando los cambios principales respecto a la versión anterior.\n\nTÍTULO ANTERIOR: ' + prev.title + '\n\nCONTENIDO ANTERIOR (HTML):\n' + prev.content
            });
          }
        } catch (e) { console.error('No se pudo cargar informe previo:', e.message); }
      }

      // Si hay informes de referencia, los añadimos como contexto adicional
      const referenceReportIds = parsed.referenceReportIds || [];
      if (Array.isArray(referenceReportIds) && referenceReportIds.length > 0) {
        try {
          const refR = await pool.query(
            'SELECT id, title, report_type, content, created_at FROM reports WHERE id = ANY($1::text[]) ORDER BY created_at ASC',
            [referenceReportIds]
          );
          if (refR.rows.length > 0) {
            let refText = '\n\n=== INFORMES DE REFERENCIA (otros informes ya generados en esta consultoría que debes tener en cuenta) ===\n' +
              'A continuación tienes ' + refR.rows.length + ' informe(s) generados previamente en esta misma consultoría. Úsalos como contexto y referencia para construir el nuevo informe: puedes basarte en sus conclusiones, ampliarlas, contrastarlas, o referenciarlas explícitamente cuando sea útil. NO los repitas literalmente — son insumos, no plantillas a copiar.\n';
            for (const ref of refR.rows) {
              const refDate = new Date(ref.created_at).toLocaleString('es-ES');
              refText += '\n--- INFORME DE REFERENCIA: "' + ref.title + '" (tipo: ' + ref.report_type + ', generado el ' + refDate + ') ---\n' + ref.content + '\n';
            }
            contentBlocks.push({ type: 'text', text: refText });
          }
        } catch (e) { console.error('No se pudieron cargar informes de referencia:', e.message); }
      }

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

      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 16000);

      // Clean code fences
      let clean = (result || '').replace(/^\s*```(?:html)?\s*\n?/i, '').replace(/\n?```\s*$/i, '').trim();

      // Save to history
      const reportId = 'r_' + Date.now();
      const finalTitle = title || (reportType + ' — ' + new Date().toLocaleString('es-ES'));
      await pool.query(
        'INSERT INTO reports (id, consultancy_id, title, report_type, content, files_used, custom_instructions, target_persona_id, previous_report_id, reference_report_ids) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
        [reportId, consultancyId, finalTitle, reportType || 'diagnostico_completo', clean, JSON.stringify(fileIds || []), customInstructions || null, targetPersonaId || null, previousReportId || null, JSON.stringify(referenceReportIds || [])]
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

    // DOWNLOAD notes as Word (raw export of all filled sections)
    const notesDocxMatch = url.match(/^\/api\/consultancies\/([^/]+)\/notes-docx$/);
    if (notesDocxMatch && req.method === 'GET') {
      if (!HTMLtoDOCX) return sendJSON(res, 500, { error: 'html-to-docx no disponible' });
      const r = await pool.query('SELECT * FROM consultancies WHERE id = $1', [notesDocxMatch[1]]);
      if (r.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrada' });
      const c = r.rows[0];
      const d = c.data || {};

      // Labels para cada campo
      const labels = {
        contexto: {
          _title: 'Contexto previo',
          empresa: 'Empresa / organización',
          depto: 'Departamento',
          personas: 'Número de personas en el equipo',
          fecha_visita: 'Fecha de visita',
          problema: '¿Por qué te han llamado? Problema conocido',
          objetivo: '¿Qué espera la dirección?',
          software: 'Software conocido',
          docs: 'Documentación previa disponible'
        },
        estructura: {
          _title: 'Reunión de apertura — Estructura',
          organigrama: '¿Existe organigrama actualizado? ¿Refleja la realidad?',
          roles: 'Roles y responsabilidades: ¿están claros?',
          reporte: '¿A quién reporta el departamento? ¿Con quién se coordina?',
          expectativas: '¿Qué espera el responsable de esta consultoría?',
          clima: 'Primera impresión del clima del equipo (1-5)',
          obs: 'Observaciones de apertura'
        },
        procesos: {
          _title: 'Análisis de procesos',
          procesos_principales: 'Procesos principales del departamento',
          cuellos: 'Cuellos de botella',
          duplicidades: 'Tareas duplicadas o redundantes',
          manual: 'Tareas manuales automatizables',
          tipos_disfuncion: 'Tipos de disfunción',
          documentado: 'Grado de documentación (1-5)',
          flujo_real: 'Flujo real de un proceso clave'
        },
        herramientas: {
          _title: 'Herramientas y datos',
          software_real: 'Herramientas que usan realmente',
          duplicidad: '¿Hay información duplicada en varios sistemas?',
          calidad_datos: 'Calidad de los datos (1-5)',
          procedimientos: '¿Existen procedimientos escritos? ¿Actualizados?',
          formacion: '¿Tienen formación suficiente?',
          problemas: 'Problemas detectados'
        },
        sintesis: {
          _title: 'Síntesis y cierre',
          impresion: 'Impresión general (en caliente)',
          fortalezas: 'Fortalezas detectadas',
          urgente: '¿Qué es lo más urgente de resolver?',
          resistencias: 'Resistencias al cambio detectadas',
          nivel_disfuncion: 'Nivel de disfunción global (1-5)',
          areas_mejora: 'Áreas prioritarias de mejora',
          quick_wins: 'Quick wins',
          cambios: 'Cambios estructurales necesarios'
        }
      };

      const entrevistaLabels = {
        trabajo: '¿En qué consiste tu trabajo día a día?',
        info_llega: '¿Cómo te llega la información? ¿Se pierde algo importante?',
        frustracion: '¿Qué es lo que más te dificulta o ralentiza?',
        informal: '¿Cómo funcionáis realmente? (vs el procedimiento)',
        mejora: '¿Qué cambiarías si pudieras?',
        actitud: 'Actitud ante el cambio (1-5)',
        senales: 'Señales detectadas',
        obs: 'Observaciones libres'
      };

      function hasValue(v) {
        if (v === null || v === undefined) return false;
        if (Array.isArray(v)) return v.length > 0;
        if (typeof v === 'string') return v.trim() !== '';
        return true;
      }

      function fmtValue(v) {
        if (Array.isArray(v)) return v.join(', ');
        return String(v);
      }

      let body = '';

      // Encabezado
      body += `<p style="font-family: Calibri, sans-serif; font-size:9pt; letter-spacing:2pt; color:#6B6560; margin:0 0 16pt 0;">NOTAS DE CAMPO</p>`;
      body += `<h1 style="font-family: Calibri, sans-serif; font-size:24pt; font-weight:300; color:#1A1814; margin:0 0 20pt 0;">${escapeHtml(c.nombre || 'Consultoría')}</h1>`;
      const metaParts = [];
      if (c.empresa) metaParts.push(`<strong>Empresa:</strong> ${escapeHtml(c.empresa)}`);
      if (c.depto) metaParts.push(`<strong>Departamento:</strong> ${escapeHtml(c.depto)}`);
      if (c.fecha) metaParts.push(`<strong>Fecha:</strong> ${escapeHtml(c.fecha)}`);
      if (c.consultora) metaParts.push(`<strong>Consultora:</strong> ${escapeHtml(c.consultora)}`);
      if (metaParts.length) body += `<p style="font-family: Calibri, sans-serif; font-size:10pt; color:#6B6560; margin:0 0 28pt 0; line-height:1.8;">${metaParts.join(' &nbsp;·&nbsp; ')}</p>`;

      // Secciones estándar
      for (const sec of ['contexto', 'estructura', 'procesos', 'herramientas', 'sintesis']) {
        const data = d[sec] || {};
        const section = labels[sec];
        const filled = Object.keys(data).filter(k => k !== '_title' && hasValue(data[k]) && section[k]);
        if (filled.length === 0) continue;

        body += `<h2 style="font-family: Calibri, sans-serif; font-size:16pt; font-weight:600; color:#1A1814; margin:28pt 0 12pt 0;">${escapeHtml(section._title)}</h2>`;
        for (const k of Object.keys(section)) {
          if (k === '_title') continue;
          if (!hasValue(data[k])) continue;
          body += `<p style="font-family: Calibri, sans-serif; font-size:9pt; font-weight:600; color:#6B6560; text-transform:uppercase; letter-spacing:1.5pt; margin:14pt 0 4pt 0;">${escapeHtml(section[k])}</p>`;
          body += `<p style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.6; margin:0 0 10pt 0; white-space:pre-wrap;">${escapeHtml(fmtValue(data[k]))}</p>`;
        }
      }

      // Entrevistas
      const entrevistas = Array.isArray(d.entrevistas) ? d.entrevistas : [];
      const entrevistasConDatos = entrevistas.filter(p => {
        const hasResp = p.resp && Object.keys(p.resp).some(k => hasValue(p.resp[k]));
        const hasSesiones = Array.isArray(p.sesiones) && p.sesiones.some(s => hasValue(s.notas) || hasValue(s.titulo));
        return hasResp || hasSesiones;
      });
      if (entrevistasConDatos.length > 0) {
        body += `<h2 style="font-family: Calibri, sans-serif; font-size:16pt; font-weight:600; color:#1A1814; margin:28pt 0 12pt 0;">Entrevistas individuales</h2>`;
        for (const p of entrevistasConDatos) {
          const cargo = p.cargo ? ` — ${p.cargo}` : '';
          const antig = p.antiguedad ? ` (antigüedad: ${p.antiguedad})` : '';
          body += `<h3 style="font-family: Calibri, sans-serif; font-size:12pt; font-weight:600; color:#1A1814; margin:20pt 0 8pt 0;">${escapeHtml(p.nombre + cargo + antig)}</h3>`;

          // Sesiones cronológicas
          if (Array.isArray(p.sesiones) && p.sesiones.length > 0) {
            const ordenadas = p.sesiones.slice().sort((a,b) => (a.fecha||'').localeCompare(b.fecha||''));
            const sesionesConDatos = ordenadas.filter(s => hasValue(s.notas) || hasValue(s.titulo));
            if (sesionesConDatos.length > 0) {
              body += `<p style="font-family: Calibri, sans-serif; font-size:10pt; font-weight:600; color:#1A1814; margin:14pt 0 6pt 0;">Sesiones de entrevista</p>`;
              for (const s of sesionesConDatos) {
                const fecha = s.fecha || 'sin fecha';
                const titulo = s.titulo ? ` — ${s.titulo}` : '';
                body += `<p style="font-family: Calibri, sans-serif; font-size:9pt; font-weight:600; color:#6B6560; margin:10pt 0 3pt 0;">📅 ${escapeHtml(fecha + titulo)}</p>`;
                if (s.notas) {
                  body += `<p style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.6; margin:0 0 8pt 0; white-space:pre-wrap; padding-left:12pt; border-left:2pt solid #E2DDD5;">${escapeHtml(s.notas)}</p>`;
                }
              }
            }
          }

          // Resumen vivo (campos estructurados)
          const r = p.resp || {};
          const hasResumen = Object.keys(entrevistaLabels).some(k => hasValue(r[k]));
          if (hasResumen) {
            body += `<p style="font-family: Calibri, sans-serif; font-size:10pt; font-weight:600; color:#1A1814; margin:16pt 0 6pt 0;">Resumen consolidado</p>`;
            for (const k of Object.keys(entrevistaLabels)) {
              if (!hasValue(r[k])) continue;
              body += `<p style="font-family: Calibri, sans-serif; font-size:9pt; font-weight:600; color:#6B6560; text-transform:uppercase; letter-spacing:1.5pt; margin:10pt 0 4pt 0;">${escapeHtml(entrevistaLabels[k])}</p>`;
              body += `<p style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.6; margin:0 0 10pt 0; white-space:pre-wrap;">${escapeHtml(fmtValue(r[k]))}</p>`;
            }
          }
        }
      }

      // Archivos adjuntos
      const archivos = await pool.query('SELECT name, description, mime_type, size FROM files WHERE consultancy_id = $1 ORDER BY created_at', [c.id]);
      if (archivos.rows.length > 0) {
        body += `<h2 style="font-family: Calibri, sans-serif; font-size:16pt; font-weight:600; color:#1A1814; margin:28pt 0 12pt 0;">Archivos adjuntos</h2>`;
        body += '<ul style="margin:6pt 0 12pt 0; padding-left:20pt;">';
        for (const f of archivos.rows) {
          const desc = f.description ? ` — ${escapeHtml(f.description)}` : '';
          body += `<li style="font-family: Calibri, sans-serif; font-size:11pt; line-height:1.55; margin-bottom:4pt;"><strong>${escapeHtml(f.name)}</strong>${desc}</li>`;
        }
        body += '</ul>';
      }

      const fullHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>${body}</body></html>`;

      let buffer;
      try {
        buffer = await HTMLtoDOCX(fullHtml, null, {
          orientation: 'portrait',
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          font: 'Calibri',
          fontSize: 22,
          title: 'Notas — ' + (c.nombre || 'Consultoría'),
          creator: c.consultora || 'Consultoría con Propósito'
        });
      } catch (genErr) {
        console.error('Error generando notas docx:', genErr.message);
        return sendJSON(res, 500, { error: 'No se pudo generar el Word: ' + genErr.message });
      }

      const fileName = 'Notas_' + (c.nombre || 'consultoria').replace(/[^a-z0-9áéíóúñ\s\-_]/gi, '').replace(/\s+/g, '_') + '.docx';
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(fileName)}"`,
        'Content-Length': buffer.length
      });
      res.end(buffer);
      return;
    }

    // UPDATE file link (project/task)
    const fileLinkMatch = url.match(/^\/api\/files\/([^/]+)\/link$/);
    if (fileLinkMatch && req.method === 'PUT') {
      const body = await readBody(req);
      const { linkedProjectId, linkedTaskId } = JSON.parse(body);
      await pool.query(
        'UPDATE files SET linked_project_id = $1, linked_task_id = $2 WHERE id = $3',
        [linkedProjectId || null, linkedTaskId || null, fileLinkMatch[1]]
      );
      return sendJSON(res, 200, { ok: true });
    }

    // SEARCH consultancy (global search)
    const searchMatch = url.match(/^\/api\/consultancies\/([^/]+)\/search$/);
    if (searchMatch && req.method === 'POST') {
      const body = await readBody(req);
      const { query } = JSON.parse(body || '{}');
      const q = (query || '').trim();
      if (!q || q.length < 2) return sendJSON(res, 200, { results: [] });
      const cId = searchMatch[1];
      const qLower = q.toLowerCase();
      const results = [];

      // 1) Consultancy fields (data JSONB)
      const cRes = await pool.query('SELECT nombre, empresa, depto, data FROM consultancies WHERE id = $1', [cId]);
      if (cRes.rows.length === 0) return sendJSON(res, 404, { error: 'No encontrada' });
      const c = cRes.rows[0];
      const d = c.data || {};

      function makeSnippet(text, query, ctx = 80) {
        const idx = text.toLowerCase().indexOf(query.toLowerCase());
        if (idx === -1) return text.slice(0, 200);
        const start = Math.max(0, idx - ctx);
        const end = Math.min(text.length, idx + query.length + ctx);
        return (start > 0 ? '…' : '') + text.slice(start, end) + (end < text.length ? '…' : '');
      }

      // Search in section fields
      const sectionLabels = {
        contexto: 'Contexto', estructura: 'Estructura', procesos: 'Procesos',
        herramientas: 'Herramientas', sintesis: 'Síntesis'
      };
      for (const sec of Object.keys(sectionLabels)) {
        const data = d[sec] || {};
        for (const k of Object.keys(data)) {
          const v = data[k];
          if (typeof v !== 'string' || !v) continue;
          if (v.toLowerCase().includes(qLower)) {
            results.push({
              type: 'section', category: 'Secciones',
              title: sectionLabels[sec] + ' — ' + k,
              snippet: makeSnippet(v, q),
              nav: sec
            });
          }
        }
      }

      // Interviews
      (d.entrevistas || []).forEach(p => {
        // Person name/cargo
        if ((p.nombre || '').toLowerCase().includes(qLower) || (p.cargo || '').toLowerCase().includes(qLower)) {
          results.push({
            type: 'person', category: 'Personas',
            title: p.nombre + (p.cargo ? ' — ' + p.cargo : ''),
            snippet: 'Antigüedad: ' + (p.antiguedad || 'no indicada'),
            nav: 'entrevistas', personaId: p.id
          });
        }
        // Sessions
        (p.sesiones || []).forEach(s => {
          const combined = (s.titulo || '') + '\n' + (s.notas || '');
          if (combined.toLowerCase().includes(qLower)) {
            results.push({
              type: 'session', category: 'Sesiones de entrevista',
              title: p.nombre + (s.fecha ? ' · ' + s.fecha : '') + (s.titulo ? ' — ' + s.titulo : ''),
              snippet: makeSnippet(s.notas || s.titulo || '', q),
              nav: 'entrevistas', personaId: p.id
            });
          }
        });
        // Resp fields
        const r = p.resp || {};
        Object.keys(r).forEach(k => {
          const v = r[k];
          if (typeof v !== 'string' || !v) return;
          if (v.toLowerCase().includes(qLower)) {
            results.push({
              type: 'field', category: 'Campos de entrevista',
              title: p.nombre + ' — ' + k,
              snippet: makeSnippet(v, q),
              nav: 'entrevistas', personaId: p.id
            });
          }
        });
      });

      // Reports
      const rRes = await pool.query(
        `SELECT id, title, report_type, content, created_at FROM reports WHERE consultancy_id = $1 AND (LOWER(title) LIKE $2 OR LOWER(content) LIKE $2)`,
        [cId, '%' + qLower + '%']
      );
      rRes.rows.forEach(r => {
        // Strip HTML for snippet
        const textOnly = (r.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
        results.push({
          type: 'report', category: 'Informes',
          title: r.title,
          snippet: makeSnippet(textOnly, q),
          nav: 'informe', reportId: r.id
        });
      });

      // Files (name, description, extracted_text)
      const fRes = await pool.query(
        `SELECT id, name, description, extracted_text, linked_project_id FROM files WHERE consultancy_id = $1 AND (LOWER(name) LIKE $2 OR LOWER(COALESCE(description,'')) LIKE $2 OR LOWER(COALESCE(extracted_text,'')) LIKE $2)`,
        [cId, '%' + qLower + '%']
      );
      fRes.rows.forEach(f => {
        let snippet = '';
        if ((f.description || '').toLowerCase().includes(qLower)) {
          snippet = makeSnippet(f.description, q);
        } else if ((f.extracted_text || '').toLowerCase().includes(qLower)) {
          snippet = makeSnippet(f.extracted_text, q);
        } else {
          snippet = f.description || 'Nombre de archivo coincide';
        }
        results.push({
          type: 'file', category: 'Archivos',
          title: f.name,
          snippet: snippet,
          nav: 'archivos', fileId: f.id
        });
      });

      // Projects and tasks/interactions
      const pRes = await pool.query('SELECT * FROM projects WHERE consultancy_id = $1', [cId]);
      pRes.rows.forEach(pr => {
        const combined = (pr.title || '') + '\n' + (pr.description || '') + '\n' + (pr.notes || '');
        if (combined.toLowerCase().includes(qLower)) {
          results.push({
            type: 'project', category: 'Proyectos',
            title: pr.title,
            snippet: makeSnippet(pr.description || pr.notes || pr.title, q),
            nav: 'proyectos', projectId: pr.id
          });
        }
        // Tasks
        const tasks = Array.isArray(pr.tasks) ? pr.tasks : [];
        tasks.forEach(t => {
          const tc = (t.title || '') + '\n' + (t.notes || '');
          if (tc.toLowerCase().includes(qLower)) {
            results.push({
              type: 'task', category: 'Tareas',
              title: pr.title + ' → ' + (t.title || 'Sin título'),
              snippet: makeSnippet(t.notes || t.title || '', q),
              nav: 'proyectos', projectId: pr.id, taskId: t.id
            });
          }
        });
        // Interactions
        const interactions = Array.isArray(pr.interactions) ? pr.interactions : [];
        interactions.forEach(it => {
          const ic = (it.title || '') + '\n' + (it.notes || '');
          if (ic.toLowerCase().includes(qLower)) {
            results.push({
              type: 'interaction', category: 'Interacciones',
              title: pr.title + ' → ' + (it.title || it.type || 'Interacción') + (it.date ? ' (' + it.date + ')' : ''),
              snippet: makeSnippet(it.notes || '', q),
              nav: 'proyectos', projectId: pr.id, interactionId: it.id
            });
          }
        });
      });

      return sendJSON(res, 200, { results, total: results.length });
    }

    // ── POST-CONSULTANCY: close & projects ──

    // Close consultancy → seguimiento
    const closeMatch = url.match(/^\/api\/consultancies\/([^/]+)\/close$/);
    if (closeMatch && req.method === 'POST') {
      const body = await readBody(req);
      const { officialReportIds } = JSON.parse(body || '{}');
      await pool.query(
        `UPDATE consultancies SET status = 'follow_up', closed_at = NOW(), official_report_ids = $1, updated_at = NOW() WHERE id = $2`,
        [JSON.stringify(officialReportIds || []), closeMatch[1]]
      );
      return sendJSON(res, 200, { ok: true });
    }

    // Reopen consultancy (undo close)
    const reopenMatch = url.match(/^\/api\/consultancies\/([^/]+)\/reopen$/);
    if (reopenMatch && req.method === 'POST') {
      await pool.query(
        `UPDATE consultancies SET status = 'active', closed_at = NULL, updated_at = NOW() WHERE id = $1`,
        [reopenMatch[1]]
      );
      return sendJSON(res, 200, { ok: true });
    }

    // List projects of a consultancy
    const projListMatch = url.match(/^\/api\/consultancies\/([^/]+)\/projects$/);
    if (projListMatch && req.method === 'GET') {
      const r = await pool.query(
        `SELECT * FROM projects WHERE consultancy_id = $1 ORDER BY sort_order ASC, created_at ASC`,
        [projListMatch[1]]
      );
      return sendJSON(res, 200, r.rows);
    }

    // Create project
    if (projListMatch && req.method === 'POST') {
      const body = await readBody(req);
      const p = JSON.parse(body);
      const id = 'proj_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
      await pool.query(
        `INSERT INTO projects (id, consultancy_id, title, description, status, priority, responsible_persona_id, target_date, progress, notes, sort_order, tasks, interactions, progress_manual)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`,
        [id, projListMatch[1], p.title || 'Sin título', p.description || null, p.status || 'pending',
         p.priority || 'medium', p.responsible_persona_id || null, p.target_date || null,
         p.progress || 0, p.notes || null, p.sort_order || 0,
         JSON.stringify(p.tasks || []), JSON.stringify(p.interactions || []), p.progress_manual === true]
      );
      return sendJSON(res, 200, { id, ok: true });
    }

    // Update project
    const projItemMatch = url.match(/^\/api\/projects\/([^/]+)$/);
    if (projItemMatch && req.method === 'PUT') {
      const body = await readBody(req);
      const p = JSON.parse(body);
      await pool.query(
        `UPDATE projects SET title=$1, description=$2, status=$3, priority=$4, responsible_persona_id=$5,
                target_date=$6, progress=$7, notes=$8, sort_order=$9, tasks=$10, interactions=$11, progress_manual=$12, updated_at=NOW()
         WHERE id=$13`,
        [p.title, p.description || null, p.status, p.priority, p.responsible_persona_id || null,
         p.target_date || null, p.progress || 0, p.notes || null, p.sort_order || 0,
         JSON.stringify(p.tasks || []), JSON.stringify(p.interactions || []), p.progress_manual === true,
         projItemMatch[1]]
      );
      return sendJSON(res, 200, { ok: true });
    }

    // Delete project
    if (projItemMatch && req.method === 'DELETE') {
      await pool.query('DELETE FROM projects WHERE id = $1', [projItemMatch[1]]);
      return sendJSON(res, 200, { ok: true });
    }

    // ── ASK AI (chat contextual: consultancy / project / task) ──
    if (url === '/api/ask-ai' && req.method === 'POST') {
      const body = await readBody(req);
      const { scope, entityId, question, presetKey, parentProjectId } = JSON.parse(body || '{}');
      if (!scope || !entityId) return sendJSON(res, 400, { error: 'Faltan datos (scope y entityId)' });
      const userQuestion = (question || '').trim();
      if (!userQuestion && !presetKey) return sendJSON(res, 400, { error: 'Falta la pregunta' });

      // Presets universales (adaptables al scope)
      const PRESETS = {
        // Consultoría
        estado_general: 'Dame una lectura honesta y breve del estado actual del acompañamiento. ¿Cómo vamos? ¿Dónde va bien y dónde no? Sin adornos.',
        priorizar_foco: 'Sugiere en qué proyectos o áreas debería centrar mi foco las próximas 2-3 semanas, y explícame por qué. Sé concreto: nombra los proyectos.',
        mensaje_cliente: 'Redacta un mensaje profesional para enviar al cliente informando del estado global. Tono cercano pero cuidado. Que se pueda copiar y enviar tal cual (ajustable).',
        cruce_patrones: '¿Qué patrones detectas cruzando la información de todos los proyectos? Personas sobrecargadas, bloqueos que se repiten, áreas paradas, ritmos. Solo lo que veas de verdad.',
        // Proyecto
        email_responsable: 'Redacta un email de seguimiento profesional para el responsable de este proyecto. Ajustado al estado actual. Que pueda copiarlo y enviarlo tal cual (adaptándolo si quiero).',
        proximos_pasos: 'Sugiere los 3 próximos pasos más aterrizados y prácticos para hacer avanzar este proyecto en las próximas 2-3 semanas. Cada paso: acción concreta, quién, cuándo.',
        desbloquear: 'Este proyecto necesita moverse. Dame ideas concretas y prácticas para desbloquearlo: qué conversaciones tener, qué decisiones forzar, qué cambios probar. Sin abstracciones.',
        preparar_reunion: 'Ayúdame a preparar la próxima reunión sobre este proyecto: propuesta de agenda, preguntas clave para el cliente, temas que conviene sacar. Adáptalo al estado actual.',
        // Tarea
        recordatorio_tarea: 'Redacta un recordatorio profesional y cercano para el responsable de esta tarea. Ajustado al contexto (fecha, retraso, contenido). Directo pero amable.',
        descomponer_tarea: 'Descompón esta tarea en 3-5 subpasos concretos y accionables. Cada uno: qué hacer, ideas de cómo, en qué orden.',
        desbloquear_tarea: 'Esta tarea está atascada. Dame ideas concretas para desbloquearla: qué probar, con quién hablar, qué alternativa considerar.'
      };

      const questionToSend = userQuestion || PRESETS[presetKey] || 'Ayúdame con este proyecto.';

      // Construir contexto según scope
      let contextBlock = '';
      let consultancyId = null;
      let personas = [];

      if (scope === 'consultancy') {
        consultancyId = entityId;
        const cR = await pool.query('SELECT * FROM consultancies WHERE id = $1', [entityId]);
        if (cR.rows.length === 0) return sendJSON(res, 404, { error: 'Consultoría no encontrada' });
        const c = cR.rows[0];
        personas = (c.data && c.data.entrevistas) || [];
        const pR = await pool.query('SELECT id, title, description, status, priority, responsible_persona_id, target_date, progress, notes, tasks, interactions FROM projects WHERE consultancy_id = $1 ORDER BY sort_order ASC, created_at ASC', [entityId]);
        const projects = pR.rows;
        const rR = await pool.query('SELECT id, title, report_type, created_at FROM reports WHERE consultancy_id = $1 ORDER BY created_at DESC LIMIT 20', [entityId]);
        const reports = rR.rows;
        const fR = await pool.query('SELECT name, description, LEFT(COALESCE(extracted_text, \'\'), 800) AS preview FROM files WHERE consultancy_id = $1', [entityId]);
        const files = fR.rows;

        const projectsTxt = projects.map(p => {
          const tasks = Array.isArray(p.tasks) ? p.tasks : [];
          const interactions = Array.isArray(p.interactions) ? p.interactions : [];
          const done = tasks.filter(t => t.done).length;
          const resp = personas.find(x => x.id === p.responsible_persona_id);
          const respName = p.responsible_persona_id === '__me__' ? 'Yo (consultora)' : (resp ? resp.nombre : 'sin asignar');
          return `- [${p.status}] "${p.title}" (${done}/${tasks.length} tareas, ${interactions.length} interacciones, resp: ${respName}, avance ${p.progress||0}%)${p.description?': '+p.description.substring(0,120):''}`;
        }).join('\n');

        const personasTxt = personas.map(p => `- ${p.nombre}${p.cargo?' ('+p.cargo+')':''}`).join('\n');

        contextBlock = `=== CONSULTORÍA: ${c.nombre} ===
Cliente: ${c.empresa || c.nombre}${c.depto?' - Depto: '+c.depto:''}
Estado: ${c.status || 'active'}${c.closed_at?' (cerrada '+new Date(c.closed_at).toLocaleDateString('es-ES')+')':''}

--- EQUIPO DEL CLIENTE (${personas.length} personas) ---
${personasTxt || '(sin personas registradas)'}

--- PROYECTOS DE MEJORA (${projects.length}) ---
${projectsTxt || '(sin proyectos)'}

--- INFORMES GENERADOS (${reports.length} más recientes) ---
${reports.map(r => `- ${r.title} (${r.report_type}, ${new Date(r.created_at).toLocaleDateString('es-ES')})`).join('\n') || '(sin informes)'}

--- ARCHIVOS ADJUNTOS (${files.length}) ---
${files.map(f => `- ${f.name}${f.description?': '+f.description:''}`).join('\n') || '(sin archivos)'}
`;

      } else if (scope === 'project') {
        const pR = await pool.query('SELECT p.*, c.nombre AS c_nombre, c.empresa, c.depto FROM projects p LEFT JOIN consultancies c ON c.id = p.consultancy_id WHERE p.id = $1', [entityId]);
        if (pR.rows.length === 0) return sendJSON(res, 404, { error: 'Proyecto no encontrado' });
        const p = pR.rows[0];
        consultancyId = p.consultancy_id;
        const personasR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [p.consultancy_id]);
        personas = (personasR.rows[0] && personasR.rows[0].data && personasR.rows[0].data.entrevistas) || [];

        const tasks = Array.isArray(p.tasks) ? p.tasks : [];
        const interactions = Array.isArray(p.interactions) ? p.interactions : [];
        const responsibleP = personas.find(x => x.id === p.responsible_persona_id);
        const respName = p.responsible_persona_id === '__me__' ? 'Yo (consultora)' : (responsibleP ? responsibleP.nombre + (responsibleP.cargo?' ('+responsibleP.cargo+')':'') : 'sin asignar');

        const tasksTxt = tasks.map(t => {
          const tResp = personas.find(x => x.id === t.responsible_persona_id);
          const tRespName = t.responsible_persona_id === '__me__' ? 'Yo' : (tResp ? tResp.nombre : 'sin asignar');
          const state = t.done ? '[HECHA]' : (t.status === 'blocked' ? '[BLOQUEADA]' : (t.status === 'in_progress' ? '[EN CURSO]' : '[PENDIENTE]'));
          return `- ${state} ${t.title} (resp: ${tRespName}${t.target_date?', fecha: '+t.target_date:''})${t.notes?' - '+t.notes:''}`;
        }).join('\n');

        const interSorted = interactions.slice().sort((a,b) => (b.date||'').localeCompare(a.date||''));
        const interTxt = interSorted.slice(0, 10).map(i => `- ${i.date || 'sin fecha'} · ${i.type || 'interacción'}${i.title?' - '+i.title:''}${i.notes?': '+i.notes.substring(0,200):''}`).join('\n');

        // Archivos vinculados al proyecto
        const fR = await pool.query('SELECT name, description, LEFT(COALESCE(extracted_text, \'\'), 500) AS preview FROM files WHERE linked_project_id = $1', [entityId]);
        const filesTxt = fR.rows.map(f => `- ${f.name}${f.description?': '+f.description:''}`).join('\n');

        const statusLabels = {pending:'Pendiente', in_progress:'En curso', blocked:'Bloqueado', completed:'Completado', discarded:'Descartado'};

        contextBlock = `=== PROYECTO: ${p.title} ===
Cliente: ${p.empresa || p.c_nombre}${p.depto?' - '+p.depto:''}
Descripción: ${p.description || '(sin descripción)'}
Estado: ${statusLabels[p.status] || p.status}
Prioridad: ${p.priority}
Responsable en el cliente: ${respName}
Fecha objetivo: ${p.target_date || 'no definida'}
Progreso: ${p.progress || 0}%
Notas del proyecto: ${p.notes || '(sin notas)'}

--- TAREAS (${tasks.length}) ---
${tasksTxt || '(sin tareas)'}

--- INTERACCIONES RECIENTES (${interactions.length} totales, últimas 10) ---
${interTxt || '(sin interacciones)'}

--- ARCHIVOS VINCULADOS AL PROYECTO (${fR.rows.length}) ---
${filesTxt || '(sin archivos vinculados)'}

--- EQUIPO DEL CLIENTE (${personas.length} personas) ---
${personas.map(p => `- ${p.nombre}${p.cargo?' ('+p.cargo+')':''}`).join('\n') || '(sin personas)'}
`;

      } else if (scope === 'task') {
        // parentProjectId es obligatorio para localizar la tarea
        if (!parentProjectId) return sendJSON(res, 400, { error: 'Falta parentProjectId para el scope task' });
        const pR = await pool.query('SELECT p.*, c.nombre AS c_nombre, c.empresa FROM projects p LEFT JOIN consultancies c ON c.id = p.consultancy_id WHERE p.id = $1', [parentProjectId]);
        if (pR.rows.length === 0) return sendJSON(res, 404, { error: 'Proyecto padre no encontrado' });
        const p = pR.rows[0];
        consultancyId = p.consultancy_id;
        const personasR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [p.consultancy_id]);
        personas = (personasR.rows[0] && personasR.rows[0].data && personasR.rows[0].data.entrevistas) || [];
        const tasks = Array.isArray(p.tasks) ? p.tasks : [];
        const task = tasks.find(t => t.id === entityId);
        if (!task) return sendJSON(res, 404, { error: 'Tarea no encontrada' });

        const tResp = personas.find(x => x.id === task.responsible_persona_id);
        const tRespName = task.responsible_persona_id === '__me__' ? 'Yo (consultora)' : (tResp ? tResp.nombre + (tResp.cargo?' ('+tResp.cargo+')':'') : 'sin asignar');
        const state = task.done ? 'HECHA' : (task.status === 'blocked' ? 'BLOQUEADA' : (task.status === 'in_progress' ? 'EN CURSO' : 'PENDIENTE'));

        contextBlock = `=== TAREA: ${task.title} ===
Estado: ${state}
Responsable: ${tRespName}
Fecha objetivo: ${task.target_date || 'no definida'}
Notas: ${task.notes || '(sin notas)'}

--- PROYECTO PADRE ---
Título: ${p.title}
Descripción: ${p.description || '(sin descripción)'}
Cliente: ${p.empresa || p.c_nombre}

--- OTRAS TAREAS DEL MISMO PROYECTO (contexto) ---
${tasks.filter(t => t.id !== entityId).map(t => {
  const state2 = t.done ? '[HECHA]' : (t.status === 'blocked' ? '[BLOQUEADA]' : '[PENDIENTE]');
  return `- ${state2} ${t.title}`;
}).join('\n') || '(sin otras tareas)'}
`;
      } else {
        return sendJSON(res, 400, { error: 'Scope inválido' });
      }

      // Prompt final
      const prompt = `Eres una asistente de trabajo de una consultora de procesos organizacionales. Ella te consulta sobre un caso concreto que tiene entre manos, y tú le ayudas con criterio: sugerencias prácticas, borradores de mensajes, análisis, próximos pasos aterrizados.

Estilo de respuesta:
- Cercana, profesional, directa. Nada de rodeos ni disclaimers largos.
- Si te piden un email o mensaje, escríbelo listo para copiar y enviar.
- Si te piden análisis o sugerencias, sé concreta: nombra proyectos, personas y acciones específicas. Nada de generalidades.
- Cuando sugieras próximos pasos, que sean implementables desde mañana.
- Usa el contexto que te dan a fondo. Si algo importante no está en el contexto, dilo.
- Longitud: la que hace falta. Ni una palabra de más. Prefiere claridad a extensión.

--- CONTEXTO DEL CASO ---
${contextBlock}

--- PREGUNTA DE LA CONSULTORA ---
${questionToSend}

--- FORMATO DE SALIDA ---
HTML puro, sencillo. Usa <p>, <strong>, <ul>/<li>, <h3> cuando aporten claridad. Nada de markdown, ni bloques \`\`\`. Empieza directamente con el contenido.`;

      const contentBlocks = [{ type: 'text', text: prompt }];
      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 4000);
      let clean = result.replace(/```html\s*/gi,'').replace(/```\s*$/g,'').trim();
      return sendJSON(res, 200, { answer: clean });
    }

    // Guardar respuesta de IA como nota (según scope)
    if (url === '/api/ai/save-as-note' && req.method === 'POST') {
      const body = await readBody(req);
      const { scope, entityId, parentProjectId, answerHtml, question } = JSON.parse(body || '{}');
      if (!scope || !entityId || !answerHtml) return sendJSON(res, 400, { error: 'Faltan datos' });
      // Convertir HTML a texto plano razonable
      const plain = answerHtml
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|li|h[1-6])>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/\n{3,}/g, '\n\n').trim();
      const stamp = new Date().toLocaleDateString('es-ES', {day:'numeric', month:'short', year:'numeric'}) + ' ' + new Date().toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'});
      const header = '💬 Consulta a IA — ' + stamp + (question ? '\nPregunta: ' + question : '') + '\n\n';
      const noteToAppend = header + plain + '\n---\n';

      if (scope === 'project') {
        await pool.query('UPDATE projects SET notes = COALESCE(notes, \'\') || $1, updated_at = NOW() WHERE id = $2', ['\n\n' + noteToAppend, entityId]);
        return sendJSON(res, 200, { ok: true });
      }
      if (scope === 'task') {
        if (!parentProjectId) return sendJSON(res, 400, { error: 'Falta parentProjectId' });
        const pR = await pool.query('SELECT tasks FROM projects WHERE id = $1', [parentProjectId]);
        if (pR.rows.length === 0) return sendJSON(res, 404, { error: 'Proyecto padre no encontrado' });
        const tasks = Array.isArray(pR.rows[0].tasks) ? pR.rows[0].tasks : [];
        const updated = tasks.map(t => t.id === entityId ? Object.assign({}, t, {notes: (t.notes || '') + '\n\n' + noteToAppend, updated_at: new Date().toISOString()}) : t);
        await pool.query('UPDATE projects SET tasks = $1, updated_at = NOW() WHERE id = $2', [JSON.stringify(updated), parentProjectId]);
        return sendJSON(res, 200, { ok: true });
      }
      if (scope === 'consultancy') {
        // Se guarda en data.sintesis.notas_ia
        const cR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [entityId]);
        if (cR.rows.length === 0) return sendJSON(res, 404, { error: 'Consultoría no encontrada' });
        const data = cR.rows[0].data || {};
        if (!data.sintesis) data.sintesis = {};
        data.sintesis.notas_ia = (data.sintesis.notas_ia || '') + '\n\n' + noteToAppend;
        await pool.query('UPDATE consultancies SET data = $1, updated_at = NOW() WHERE id = $2', [JSON.stringify(data), entityId]);
        return sendJSON(res, 200, { ok: true });
      }
      return sendJSON(res, 400, { error: 'Scope inválido' });
    }

    // Informe global de seguimiento (todos los proyectos de la consultoría)
    const globalReportMatch = url.match(/^\/api\/consultancies\/([^/]+)\/global-follow-up-report$/);
    if (globalReportMatch && req.method === 'POST') {
      const cId = globalReportMatch[1];
      const cR = await pool.query('SELECT nombre, empresa, depto, consultora, fecha, closed_at, data FROM consultancies WHERE id = $1', [cId]);
      if (cR.rows.length === 0) return sendJSON(res, 404, { error: 'Consultoría no encontrada' });
      const c = cR.rows[0];
      const personas = (c.data && c.data.entrevistas) || [];

      const pR = await pool.query('SELECT * FROM projects WHERE consultancy_id = $1 ORDER BY sort_order ASC, created_at ASC', [cId]);
      const projects = pR.rows;
      if (projects.length === 0) {
        return sendJSON(res, 400, { error: 'Aún no hay proyectos creados. Crea o extrae alguno antes de generar el informe global.' });
      }

      const statusLabels = {pending:'Pendiente', in_progress:'En curso', blocked:'Bloqueado', completed:'Completado', discarded:'Descartado'};
      const priorityLabels = {high:'Alta', medium:'Media', low:'Baja'};

      // Contexto agregado por proyecto
      const projectsTxt = projects.map((p, idx) => {
        const tasks = Array.isArray(p.tasks) ? p.tasks : [];
        const interactions = Array.isArray(p.interactions) ? p.interactions : [];
        const doneTasks = tasks.filter(t => t.done);
        const pendingTasks = tasks.filter(t => !t.done && t.status !== 'discarded');
        const blockedTasks = tasks.filter(t => !t.done && t.status === 'blocked');
        const resp = personas.find(x => x.id === p.responsible_persona_id);
        const respName = p.responsible_persona_id === '__me__' ? 'Yo (consultora)' : (resp ? resp.nombre + (resp.cargo?' ('+resp.cargo+')':'') : 'Sin asignar');
        // Cálculo de progreso
        const activeTasks = tasks.filter(t => t.status !== 'discarded');
        let progress = p.progress || 0;
        if (!p.progress_manual && activeTasks.length > 0) {
          progress = Math.round((doneTasks.length / activeTasks.length) * 100);
        }
        // Última interacción
        const sortedInter = interactions.slice().sort((a,b) => (b.date||'').localeCompare(a.date||''));
        const lastInter = sortedInter[0];
        // Tarea más urgente pendiente
        const urgentTask = pendingTasks.slice().sort((a,b) => (a.target_date||'9999').localeCompare(b.target_date||'9999'))[0];

        // Detalle de tareas
        const tasksDetail = tasks.map(t => {
          const tResp = personas.find(x => x.id === t.responsible_persona_id);
          const tRespName = t.responsible_persona_id === '__me__' ? 'Yo' : (tResp ? tResp.nombre : 'sin asignar');
          const state = t.done ? '[HECHA]' : (t.status === 'blocked' ? '[BLOQUEADA]' : (t.status === 'in_progress' ? '[EN CURSO]' : '[PENDIENTE]'));
          return `    · ${state} ${t.title} (resp: ${tRespName}${t.target_date?', fecha: '+t.target_date:''})`;
        }).join('\n');

        // Detalle de interacciones
        const interDetail = sortedInter.slice(0, 5).map(i => `    · ${i.date || 'sin fecha'} · ${i.type || 'interacción'}${i.title?' — '+i.title:''}${i.notes?': '+i.notes.substring(0,150):''}`).join('\n');

        return `--- PROYECTO ${idx+1}: ${p.title} ---
Estado: ${statusLabels[p.status] || p.status}
Prioridad: ${priorityLabels[p.priority] || p.priority}
Responsable en el cliente: ${respName}
Fecha objetivo: ${p.target_date || 'sin definir'}
Progreso: ${progress}%${p.progress_manual?' (fijado manualmente)':''}
Descripción: ${p.description || '(sin descripción)'}
Notas generales: ${p.notes || '(ninguna)'}

Tareas (${tasks.length} — ${doneTasks.length} hechas, ${pendingTasks.length} pendientes, ${blockedTasks.length} bloqueadas):
${tasksDetail || '    (sin tareas)'}

Interacciones (${interactions.length}${lastInter?', última: '+lastInter.date:''}):
${interDetail || '    (sin interacciones)'}
`;
      }).join('\n');

      // Datos agregados globales
      const today = new Date(); today.setHours(0,0,0,0);
      const monthAgo = new Date(today); monthAgo.setDate(monthAgo.getDate() - 30);
      const allTasks = projects.flatMap(p => (p.tasks || []).map(t => ({...t, _project: p.title})));
      const allInter = projects.flatMap(p => (p.interactions || []).map(i => ({...i, _project: p.title})));
      const recentInter = allInter.filter(i => i.date && new Date(i.date+'T00:00:00') >= monthAgo);
      const byStatus = { pending: 0, in_progress: 0, blocked: 0, completed: 0, discarded: 0 };
      projects.forEach(p => { byStatus[p.status] = (byStatus[p.status] || 0) + 1; });

      // Reparto de responsables (para detectar sobrecargas)
      const respMap = {};
      projects.forEach(p => {
        if (!p.responsible_persona_id) return;
        respMap[p.responsible_persona_id] = (respMap[p.responsible_persona_id] || 0) + 1;
      });
      const respTxt = Object.keys(respMap).map(rid => {
        const name = rid === '__me__' ? 'Yo (consultora)' : (personas.find(x => x.id === rid)||{}).nombre || 'Desconocido';
        return `- ${name}: responsable en ${respMap[rid]} proyecto(s)`;
      }).join('\n');

      const closedDate = c.closed_at ? new Date(c.closed_at).toLocaleDateString('es-ES', {day:'numeric', month:'long', year:'numeric'}) : 'no cerrada aún';

      const prompt = `Eres una consultora de procesos organizacionales redactando un INFORME GLOBAL DE SEGUIMIENTO sobre TODOS los proyectos de mejora derivados de una consultoría entregada a un cliente. El objetivo es dar una foto panorámica y honesta del estado del acompañamiento a fecha de hoy.

--- DATOS DE LA CONSULTORÍA ---
Cliente: ${c.empresa || c.nombre}${c.depto?' — Departamento: '+c.depto:''}
Consultora: ${c.consultora || 'no indicada'}
Fecha de la consultoría original: ${c.fecha || 'no indicada'}
Fecha de cierre y paso a seguimiento: ${closedDate}
Fecha del presente informe: ${new Date().toLocaleDateString('es-ES', {day:'numeric', month:'long', year:'numeric'})}

--- CIFRAS AGREGADAS ---
Total de proyectos: ${projects.length}
Distribución por estado: ${byStatus.pending} pendientes · ${byStatus.in_progress} en curso · ${byStatus.blocked} bloqueados · ${byStatus.completed} completados · ${byStatus.discarded} descartados
Total de tareas registradas: ${allTasks.length} (${allTasks.filter(t=>t.done).length} completadas)
Total de interacciones registradas: ${allInter.length}
Interacciones en los últimos 30 días: ${recentInter.length}

--- REPARTO DE RESPONSABILIDAD (cliente) ---
${respTxt || '(sin responsables asignados en el cliente)'}

--- DETALLE DE CADA PROYECTO ---
${projectsTxt}

--- ESTRUCTURA DEL INFORME (obligatoria, en este orden) ---
1. Contexto — quién es el cliente, cuándo se cerró la consultoría original, cuánto tiempo llevamos en fase de seguimiento, cuántos proyectos activos hay.
2. Foto actual en 30 segundos — resumen ejecutivo con las cifras clave (proyectos por estado, % medio de avance real que calcules a partir de los datos, interacciones del último mes). Un párrafo denso que se pueda leer de un vistazo.
3. Estado por proyecto — para cada uno de los ${projects.length} proyectos: título, estado, % avance, tarea pendiente más urgente y última interacción registrada. Usa <h3> para el nombre de cada proyecto y párrafo/lista debajo.
4. Palancas y bloqueos transversales — aquí es donde se demuestra el análisis: cruza información entre proyectos y detecta patrones. Ejemplos posibles (solo si los datos los sustentan):
   - Proyectos que llevan mucho tiempo sin movimiento
   - Personas que aparecen como responsables en varios proyectos (posible sobrecarga)
   - Bloqueos que se repiten en varios proyectos (esperando la misma decisión, por ejemplo)
   - Áreas del departamento donde hay más movimiento vs. donde no
5. Ritmo del acompañamiento — cómo está siendo el pulso de las interacciones con el cliente. ¿Hay proyectos huérfanos que no reciben atención? ¿Estás siendo tú la que impulsa o el cliente? Sé honesta.
6. Próximos pasos priorizados — 3-5 acciones concretas para las próximas semanas, muy aterrizadas. Cruza urgencia + impacto. Cada una debe ser algo que se pueda empezar mañana.
7. Reflexión final — una lectura de conjunto, personal y honesta, sobre cómo va el acompañamiento. Qué observas del proceso, dónde ves la palanca principal, qué te preocupa y qué te esperanza. Breve pero significativa.

--- REGLAS ---
- Datos reales, cero invenciones. Si algo no se puede saber con los datos, dilo con honestidad.
- No repitas literalmente lo que ya está en el detalle: analiza, cruza, sintetiza.
- Los próximos pasos deben ser aterrizados y realmente implementables por la consultora.
- La reflexión final es personal, no corporativa. Puedes usar la primera persona.
- Tono profesional pero cercano. Nada de lenguaje vacío.

--- FORMATO DE SALIDA (obligatorio) ---
HTML puro. Cada sección principal envuelta en <div class="report-section"> con <h2>. Subsecciones con <h3>. Nada de markdown. Sin bloques \`\`\`, sin DOCTYPE. Empieza directamente con <div class="report-section"><h2>Contexto</h2>...`;

      const contentBlocks = [{ type: 'text', text: prompt }];
      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 14000);
      let clean = result.replace(/```html\s*/gi,'').replace(/```\s*$/g,'').trim();
      const dateStr = new Date().toLocaleDateString('es-ES', {day:'numeric', month:'short', year:'numeric'});
      const globalTitle = 'Seguimiento global — ' + (c.empresa || c.nombre) + ' · ' + dateStr;
      const globalReportId = 'r_' + Date.now();
      try {
        await pool.query(
          'INSERT INTO reports (id, consultancy_id, title, report_type, content, files_used, custom_instructions, target_persona_id, previous_report_id, reference_report_ids) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
          [globalReportId, cId, globalTitle, 'seguimiento_global', clean, JSON.stringify([]), null, null, null, JSON.stringify([])]
        );
      } catch (e) { console.error('No se pudo guardar informe global:', e.message); }
      return sendJSON(res, 200, {
        report: clean,
        title: globalTitle,
        reportId: globalReportId
      });
    }

    // Generar Word ad-hoc desde HTML (para actas y seguimientos sintéticos)
    if (url === '/api/docx/adhoc' && req.method === 'POST') {
      if (!HTMLtoDOCX) return sendJSON(res, 500, { error: 'html-to-docx no disponible' });
      const body = await readBody(req);
      const { title, content, cliente, depto, consultora, tipo } = JSON.parse(body || '{}');
      if (!title || !content) return sendJSON(res, 400, { error: 'Faltan datos' });

      const metaItems = [];
      if (cliente) metaItems.push(`<strong>Cliente:</strong> ${escapeHtml(cliente)}`);
      if (depto) metaItems.push(`<strong>Departamento:</strong> ${escapeHtml(depto)}`);
      metaItems.push(`<strong>Fecha:</strong> ${new Date().toLocaleDateString('es-ES')}`);
      if (consultora) metaItems.push(`<strong>Consultora:</strong> ${escapeHtml(consultora)}`);

      const fullHtml = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body>
<p style="font-family: Calibri, sans-serif; font-size:9pt; letter-spacing:2pt; color:#6B6560; margin:0 0 16pt 0;">${escapeHtml((tipo || 'INFORME').toUpperCase())}</p>
<h1 style="font-family: Calibri, sans-serif; font-size:24pt; font-weight:300; color:#1A1814; margin:0 0 20pt 0;">${escapeHtml(title)}</h1>
<p style="font-family: Calibri, sans-serif; font-size:10pt; color:#6B6560; margin:0 0 28pt 0; line-height:1.8;">${metaItems.join(' &nbsp;·&nbsp; ')}</p>
${transformReportHtml(content)}
</body></html>`;

      let buffer;
      try {
        buffer = await HTMLtoDOCX(fullHtml, null, {
          orientation: 'portrait',
          margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          font: 'Calibri',
          fontSize: 22,
          title: title,
          creator: consultora || 'Consultoría con Propósito'
        });
      } catch (genErr) {
        return sendJSON(res, 500, { error: 'No se pudo generar Word: ' + genErr.message });
      }

      const fileName = (title || 'informe').replace(/[^a-z0-9áéíóúñ\s\-_]/gi,'').replace(/\s+/g,'_') + '.docx';
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(fileName)}"`,
        'Content-Length': buffer.length
      });
      res.end(buffer);
      return;
    }

    // ── AI FOR PROJECTS (5C) ──

    // Generar acta de interacción
    const actaMatch = url.match(/^\/api\/projects\/([^/]+)\/interactions\/([^/]+)\/acta$/);
    if (actaMatch && req.method === 'POST') {
      const projectId = actaMatch[1];
      const interactionId = actaMatch[2];
      const projR = await pool.query('SELECT p.*, c.nombre AS c_nombre, c.empresa, c.depto, c.consultora FROM projects p LEFT JOIN consultancies c ON c.id = p.consultancy_id WHERE p.id = $1', [projectId]);
      if (projR.rows.length === 0) return sendJSON(res, 404, { error: 'Proyecto no encontrado' });
      const proj = projR.rows[0];
      const interactions = Array.isArray(proj.interactions) ? proj.interactions : [];
      const inter = interactions.find(i => i.id === interactionId);
      if (!inter) return sendJSON(res, 404, { error: 'Interacción no encontrada' });

      const personasR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [proj.consultancy_id]);
      const personas = (personasR.rows[0] && personasR.rows[0].data && personasR.rows[0].data.entrevistas) || [];
      const personasStr = personas.map(p => `- ${p.nombre}${p.cargo?' ('+p.cargo+')':''}`).join('\n');

      const typeLabels = {meeting:'Reunión', call:'Llamada telefónica', email:'Email', visit:'Visita', note:'Nota'};
      const typeLabel = typeLabels[inter.type] || 'Interacción';

      const prompt = `Eres una consultora que redacta un ACTA formal de una interacción con el cliente. Genera un acta profesional en HTML puro con las secciones siguientes.

--- CONTEXTO ---
Cliente: ${proj.empresa || proj.c_nombre || 'Cliente'}${proj.depto?' — '+proj.depto:''}
Proyecto: ${proj.title}
Descripción del proyecto: ${proj.description || 'No especificada'}

--- TIPO DE INTERACCIÓN ---
${typeLabel}${inter.title?' — '+inter.title:''}
Fecha: ${inter.date || 'No especificada'}

--- NOTAS DE LA INTERACCIÓN ---
${inter.notes || '(sin notas)'}

--- PERSONAS DEL EQUIPO DEL CLIENTE (para resolver @menciones) ---
${personasStr}

--- ESTRUCTURA DEL ACTA (obligatoria) ---
1. Encabezado con datos: tipo de interacción, fecha, proyecto, participantes deducidos de las notas y las @menciones
2. Contexto breve (1 párrafo) sobre por qué se produjo esta interacción y qué se buscaba
3. Puntos tratados: lista clara de los temas discutidos, con contenido específico de cada uno
4. Acuerdos alcanzados: lista de decisiones o compromisos concretos que se derivaron
5. Tareas derivadas: lista de acciones a realizar, con responsable y fecha si se puede deducir. Si no hay tareas claras, dilo honestamente.
6. Próximos pasos o siguiente encuentro previsto (si aplica)

--- REGLAS ---
- NO inventes información que no esté en las notas
- Si algo no está claro, dilo con honestidad («no se especifica en las notas»)
- Traduce @Nombre a los nombres reales
- Tono profesional pero cercano, como una consultora sería con su cliente

--- FORMATO DE SALIDA (obligatorio) ---
Solo HTML puro, sin markdown. Cada sección en <div class="report-section"> con <h2>. Sin DOCTYPE, sin <html>, sin bloques \`\`\`. Empieza directamente con <div class="report-section"><h2>Acta...`;

      const contentBlocks = [{ type: 'text', text: prompt }];
      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 8000);
      let clean = result.replace(/```html\s*/gi,'').replace(/```\s*$/g,'').trim();
      const actaTitle = 'Acta — ' + (inter.title || typeLabel) + ' (' + (inter.date || '') + ')';
      const actaReportId = 'r_' + Date.now();
      try {
        await pool.query(
          'INSERT INTO reports (id, consultancy_id, title, report_type, content, files_used, custom_instructions, target_persona_id, previous_report_id, reference_report_ids) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
          [actaReportId, proj.consultancy_id, actaTitle, 'acta_interaccion', clean, JSON.stringify([]), null, null, null, JSON.stringify([])]
        );
      } catch (e) { console.error('No se pudo guardar acta:', e.message); }
      return sendJSON(res, 200, { report: clean, title: actaTitle, reportId: actaReportId });
    }

    // Generar informe de seguimiento del proyecto
    const projReportMatch = url.match(/^\/api\/projects\/([^/]+)\/follow-up-report$/);
    if (projReportMatch && req.method === 'POST') {
      const projectId = projReportMatch[1];
      const projR = await pool.query('SELECT p.*, c.nombre AS c_nombre, c.empresa, c.depto FROM projects p LEFT JOIN consultancies c ON c.id = p.consultancy_id WHERE p.id = $1', [projectId]);
      if (projR.rows.length === 0) return sendJSON(res, 404, { error: 'Proyecto no encontrado' });
      const proj = projR.rows[0];

      const personasR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [proj.consultancy_id]);
      const personas = (personasR.rows[0] && personasR.rows[0].data && personasR.rows[0].data.entrevistas) || [];

      const tasks = Array.isArray(proj.tasks) ? proj.tasks : [];
      const interactions = Array.isArray(proj.interactions) ? proj.interactions : [];

      const tasksTxt = tasks.map(t => {
        const resp = personas.find(p => p.id === t.responsible_persona_id);
        const respName = t.responsible_persona_id === '__me__' ? 'Yo (consultora)' : (resp ? resp.nombre : 'Sin asignar');
        const state = t.done ? '[HECHA]' : (t.status === 'blocked' ? '[BLOQUEADA]' : (t.status === 'in_progress' ? '[EN CURSO]' : '[PENDIENTE]'));
        return `- ${state} ${t.title} (responsable: ${respName}${t.target_date?', fecha: '+t.target_date:''})${t.notes?' — '+t.notes:''}`;
      }).join('\n');

      const interTxt = interactions.slice().sort((a,b)=>(a.date||'').localeCompare(b.date||'')).map(i => {
        return `- ${i.date || 'sin fecha'} · ${i.type || 'interacción'}${i.title?' — '+i.title:''}${i.notes?': '+i.notes:''}`;
      }).join('\n');

      const statusLabels = {pending:'Pendiente', in_progress:'En curso', blocked:'Bloqueado', completed:'Completado', discarded:'Descartado'};

      const prompt = `Eres una consultora redactando un INFORME DE SEGUIMIENTO de un proyecto de mejora que se derivó de una consultoría. El objetivo es dar una foto clara del estado del proyecto a fecha de hoy, para uso propio o para compartir con el cliente.

--- DATOS DEL PROYECTO ---
Cliente: ${proj.empresa || proj.c_nombre}${proj.depto?' — '+proj.depto:''}
Proyecto: ${proj.title}
Descripción: ${proj.description || '(sin descripción)'}
Estado actual: ${statusLabels[proj.status] || proj.status}
Prioridad: ${proj.priority}
Fecha objetivo: ${proj.target_date || 'no definida'}
Progreso: ${proj.progress || 0}%
Notas generales: ${proj.notes || '(sin notas)'}

--- TAREAS (${tasks.length}) ---
${tasksTxt || '(sin tareas)'}

--- INTERACCIONES CRONOLÓGICAS (${interactions.length}) ---
${interTxt || '(sin interacciones registradas)'}

--- ESTRUCTURA DEL INFORME (obligatoria) ---
1. Objetivo del proyecto (qué se quería conseguir)
2. Estado actual (dónde estamos hoy, con datos: % avance, tareas hechas/pendientes, bloqueos)
3. Avances logrados desde el inicio (qué se ha hecho, con hitos concretos)
4. Situación inicial vs. actual (comparativa honesta de dónde partíamos y dónde estamos)
5. Bloqueos y riesgos (si los hay, sino dilo)
6. Próximos pasos (lo más importante — qué hay que hacer ya, con responsables y fechas si se pueden deducir de las tareas)
7. Reflexión de la consultora (breve, pero personal — qué observas del proceso, qué recomendarías, dónde está la palanca)

--- REGLAS ---
- Datos reales, no inventes
- Honestidad sobre lo que no se sabe
- Tono profesional cercano
- Los próximos pasos deben ser aterrizados y realmente implementables

--- FORMATO DE SALIDA ---
HTML puro. Cada sección en <div class="report-section"> con <h2>. Sin markdown, sin bloques \`\`\`, sin DOCTYPE. Empieza directamente con <div class="report-section">.`;

      const contentBlocks = [{ type: 'text', text: prompt }];
      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 12000);
      let clean = result.replace(/```html\s*/gi,'').replace(/```\s*$/g,'').trim();
      const seguimTitle = 'Seguimiento — ' + proj.title + ' · ' + new Date().toLocaleDateString('es-ES', {day:'numeric', month:'short', year:'numeric'});
      const seguimReportId = 'r_' + Date.now();
      try {
        await pool.query(
          'INSERT INTO reports (id, consultancy_id, title, report_type, content, files_used, custom_instructions, target_persona_id, previous_report_id, reference_report_ids) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)',
          [seguimReportId, proj.consultancy_id, seguimTitle, 'seguimiento_proyecto', clean, JSON.stringify([]), null, null, null, JSON.stringify([])]
        );
      } catch (e) { console.error('No se pudo guardar seguimiento:', e.message); }
      return sendJSON(res, 200, { report: clean, title: seguimTitle, reportId: seguimReportId });
    }

    // Extraer proyectos automáticamente de los informes oficiales
    const extractMatch = url.match(/^\/api\/consultancies\/([^/]+)\/extract-projects$/);
    if (extractMatch && req.method === 'POST') {
      const cId = extractMatch[1];
      const body = await readBody(req);
      const { reportIds } = JSON.parse(body || '{}');
      if (!Array.isArray(reportIds) || reportIds.length === 0) {
        return sendJSON(res, 400, { error: 'Selecciona al menos un informe' });
      }
      const rR = await pool.query(
        'SELECT title, report_type, content FROM reports WHERE id = ANY($1::text[]) AND consultancy_id = $2',
        [reportIds, cId]
      );
      if (rR.rows.length === 0) return sendJSON(res, 404, { error: 'Informes no encontrados' });

      const personasR = await pool.query('SELECT data FROM consultancies WHERE id = $1', [cId]);
      const personas = (personasR.rows[0] && personasR.rows[0].data && personasR.rows[0].data.entrevistas) || [];
      const personasList = personas.map(p => `${p.nombre}${p.cargo?' ('+p.cargo+')':''} [id:${p.id}]`).join('\n');

      const reportsTxt = rR.rows.map(r => `=== INFORME: ${r.title} (tipo: ${r.report_type}) ===\n${r.content}\n`).join('\n\n');

      const prompt = `Eres una consultora que ha entregado unos informes a un cliente y ahora tiene que extraer las PROPUESTAS DE MEJORA concretas para convertirlas en fichas de proyecto de seguimiento.

--- INFORMES OFICIALES ENTREGADOS ---
${reportsTxt}

--- PERSONAS DEL EQUIPO DEL CLIENTE ---
${personasList || 'No hay personas registradas.'}

--- TAREA ---
Analiza los informes e identifica las PROPUESTAS DE MEJORA concretas y accionables. Para cada una, propon una ficha de proyecto con esta estructura:

- title: título breve y claro (máx. 80 caracteres, sin comillas dobles dentro)
- description: descripción concisa (2-3 frases) de qué se quiere conseguir (sin comillas dobles dentro; usa comillas simples si necesitas)
- priority: exactamente uno de: high | medium | low
- responsible_persona_id: si un responsable natural aparece claro en el informe, pon su id (de la lista arriba). Si no, deja "".
- rationale: breve explicación (1 frase) de por qué has propuesto este proyecto

--- REGLAS ---
- NO inventes propuestas que no estén en los informes
- Agrupa cosas relacionadas en un mismo proyecto
- Prioriza calidad sobre cantidad: mejor 4-8 proyectos bien definidos que 20 vagos
- Los títulos deben ser accionables (empieza con verbo cuando sea posible)

--- FORMATO DE SALIDA (CRÍTICO Y OBLIGATORIO) ---
Tu respuesta debe ser ÚNICAMENTE un array JSON válido. NO escribas NADA antes ni después del array. NI explicaciones, NI saludos, NI títulos, NI markdown, NI \`\`\`json.

Empieza tu respuesta directamente con el carácter [ y termínala con ].

Ejemplo exacto de formato de respuesta:
[{"title":"Digitalizar altas de empleados","description":"Implementar un flujo digital para reducir tiempos de proceso.","priority":"high","responsible_persona_id":"","rationale":"Detectado como cuello de botella principal en la sección de procesos."},{"title":"Formar al equipo en la nueva herramienta","description":"Sesiones prácticas para asegurar el uso correcto de la plataforma.","priority":"medium","responsible_persona_id":"","rationale":"Varios entrevistados mencionan falta de formación."}]

Recuerda: SOLO el JSON, nada más. Escapa correctamente cualquier carácter especial dentro de los strings.`;

      const contentBlocks = [{ type: 'text', text: prompt }];
      const result = await callAnthropic([{ role: 'user', content: contentBlocks }], 6000);
      let raw = result || '';

      // Estrategia robusta de parseo
      let projects = null;
      let parseError = null;

      // Intento 1: parseo directo
      try {
        projects = JSON.parse(raw.trim());
      } catch (e1) {
        parseError = e1.message;
        // Intento 2: quitar bloques ```
        let clean = raw.replace(/```json\s*/gi,'').replace(/```\s*/g,'').trim();
        try {
          projects = JSON.parse(clean);
        } catch (e2) {
          // Intento 3: extraer solo desde el primer [ hasta el último ] balanceado
          const startIdx = clean.indexOf('[');
          const endIdx = clean.lastIndexOf(']');
          if (startIdx >= 0 && endIdx > startIdx) {
            const extracted = clean.substring(startIdx, endIdx + 1);
            try {
              projects = JSON.parse(extracted);
            } catch (e3) {
              // Intento 4: extraer array elemento a elemento con regex tolerante
              try {
                // Buscar objetos { ... } dentro del texto y armar array
                const objects = [];
                let depth = 0;
                let start = -1;
                for (let i = 0; i < clean.length; i++) {
                  const ch = clean[i];
                  if (ch === '{') {
                    if (depth === 0) start = i;
                    depth++;
                  } else if (ch === '}') {
                    depth--;
                    if (depth === 0 && start >= 0) {
                      try {
                        const obj = JSON.parse(clean.substring(start, i + 1));
                        if (obj && obj.title) objects.push(obj);
                      } catch (_) {}
                      start = -1;
                    }
                  }
                }
                if (objects.length > 0) projects = objects;
              } catch (e4) { /* ignore */ }
            }
          }
        }
      }

      if (!Array.isArray(projects) || projects.length === 0) {
        console.error('Extract projects: JSON parse failed. Raw response preview:', raw.substring(0, 500));
        return sendJSON(res, 500, {
          error: 'La IA no devolvió un JSON válido. Suele pasar cuando los informes son muy largos o densos. Prueba a marcar menos informes, o vuelve a intentarlo.'
        });
      }

      // Validar/limpiar cada proyecto
      projects = projects.filter(p => p && typeof p === 'object' && p.title).map(p => ({
        title: String(p.title || '').substring(0, 200),
        description: String(p.description || ''),
        priority: ['high','medium','low'].indexOf(p.priority) >= 0 ? p.priority : 'medium',
        responsible_persona_id: String(p.responsible_persona_id || ''),
        rationale: String(p.rationale || '')
      }));

      return sendJSON(res, 200, { projects });
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
