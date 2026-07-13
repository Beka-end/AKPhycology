'use strict';
/**
 * Almaty Clinic — сервер психологического скрининга (HADS → Бек).
 * Express + libSQL/Turso + выгрузка в Excel (SheetJS).
 *
 * Работает и как обычный сервер (node server.js), и как serverless-функция
 * на Vercel (см. api/index.js и vercel.json).
 */
const express = require('express');
const XLSX = require('xlsx');
const crypto = require('crypto');
const path = require('path');
try { require('dotenv').config(); } catch (_) {}
const { client, ensureSchema } = require('./db');

const PUBLIC_DIR = path.join(__dirname, 'public');

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'almaty-clinic';
// Токен стабилен между запусками: берём из env либо детерминированно из пароля
const ADMIN_TOKEN = process.env.ADMIN_TOKEN
  || crypto.createHash('sha256').update('almaty-clinic:' + ADMIN_PASSWORD).digest('hex').slice(0, 40);

const LABEL = {
  hads_norm: 'норма', hads_sub: 'субклинически выраженная', hads_clin: 'клинически выраженная',
  b_none: 'отсутствие депрессии / ремиссия', b_border: 'пограничное состояние',
  b_mild: 'лёгкая депрессия', b_mod: 'депрессия средней тяжести', b_sev: 'тяжёлая депрессия',
};
const L = k => LABEL[k] || (k || '');

const app = express();
app.use(express.json({ limit: '256kb' }));

// Перед API-запросами убеждаемся, что таблица существует (идемпотентно).
// На отдачу самой страницы это не влияет — фронтенд откроется в любом случае.
app.use('/api', async (_req, res, next) => {
  try { await ensureSchema(); next(); }
  catch (e) { console.error('DB init error:', e); res.status(500).json({ error: 'db_init' }); }
});

function auth(req, res, next) {
  const h = req.headers.authorization || '';
  const q = req.query.token || '';
  if (h === 'Bearer ' + ADMIN_TOKEN || q === ADMIN_TOKEN) return next();
  res.status(401).json({ error: 'unauthorized' });
}

// Вход психолога
app.post('/api/login', (req, res) => {
  const pass = (req.body && req.body.password) || '';
  if (pass === ADMIN_PASSWORD) return res.json({ token: ADMIN_TOKEN });
  res.status(401).json({ error: 'wrong_password' });
});

// Сохранение результата (студент)
app.post('/api/records', async (req, res) => {
  try {
    const b = req.body || {};
    if (!/^\d{12}$/.test(String(b.iin || ''))) return res.status(400).json({ error: 'bad_iin' });
    if (!b.fio || !b.vuz) return res.status(400).json({ error: 'missing_fields' });
    const args = [
      new Date().toISOString(), b.ts || null,
      String(b.vuz).slice(0, 200), String(b.fio).slice(0, 200), String(b.iin),
      b.lang === 'kz' ? 'kz' : 'ru', int(b.hadsAnx), int(b.hadsDep),
      b.hadsAnxI || null, b.hadsDepI || null, b.stage2 ? 1 : 0,
      b.stage2 ? int(b.beck) : null, b.stage2 ? (b.beckI || null) : null,
      int(b.beckItem9), b.flag ? 1 : 0, JSON.stringify(b).slice(0, 20000),
    ];
    const r = await client.execute({
      sql: `INSERT INTO records
        (created_at,client_ts,vuz,fio,iin,lang,hads_anx,hads_dep,hads_anx_i,hads_dep_i,
         stage2,beck,beck_i,beck_item9,flag,raw)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      args,
    });
    res.json({ ok: true, id: Number(r.lastInsertRowid) });
  } catch (e) { console.error('save error:', e); res.status(500).json({ error: 'save_failed' }); }
});

// Список записей (психолог)
app.get('/api/records', auth, async (_req, res) => {
  try {
    const r = await client.execute('SELECT * FROM records ORDER BY id DESC');
    res.json(r.rows.map(toClient));
  } catch (e) { console.error('list error:', e); res.status(500).json({ error: 'list_failed' }); }
});

// Статистика
app.get('/api/records/stats', auth, async (_req, res) => {
  try {
    const total = (await client.execute('SELECT COUNT(*) c FROM records')).rows[0].c;
    const stage2 = (await client.execute('SELECT COUNT(*) c FROM records WHERE stage2=1')).rows[0].c;
    const flagged = (await client.execute('SELECT COUNT(*) c FROM records WHERE flag=1')).rows[0].c;
    res.json({ total, stage2, flagged });
  } catch (e) { console.error('stats error:', e); res.status(500).json({ error: 'stats_failed' }); }
});

// Очистка
app.delete('/api/records', auth, async (_req, res) => {
  try { await client.execute('DELETE FROM records'); res.json({ ok: true }); }
  catch (e) { console.error('clear error:', e); res.status(500).json({ error: 'clear_failed' }); }
});

// Выгрузка в Excel
app.get('/api/export.xlsx', auth, async (_req, res) => {
  try {
    const r = await client.execute('SELECT * FROM records ORDER BY id ASC');
    const data = r.rows.map((row, i) => ({
      '№': i + 1,
      'Дата и время': new Date(row.created_at).toLocaleString('ru-RU'),
      'ВУЗ': row.vuz, 'ФИО': row.fio, 'ИИН': row.iin,
      'Язык': row.lang === 'kz' ? 'каз' : 'рус',
      'HADS Тревога (балл)': row.hads_anx, 'HADS Тревога': L(row.hads_anx_i),
      'HADS Депрессия (балл)': row.hads_dep, 'HADS Депрессия': L(row.hads_dep_i),
      'Направлен на 2 этап': row.stage2 ? 'да' : 'нет',
      'Бек (балл)': row.stage2 ? row.beck : '', 'Бек интерпретация': row.stage2 ? L(row.beck_i) : '',
      'Отметка риска (п.9 Бек)': row.beck_item9 > 0 ? 'да' : '',
      'Требует внимания': row.flag ? 'да' : '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:4},{wch:18},{wch:24},{wch:26},{wch:14},{wch:6},{wch:14},{wch:22},
                   {wch:16},{wch:22},{wch:16},{wch:9},{wch:24},{wch:16},{wch:14}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Результаты');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const stamp = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="almaty_clinic_skrining_${stamp}.xlsx"`);
    res.send(buf);
  } catch (e) { console.error('export error:', e); res.status(500).json({ error: 'export_failed' }); }
});

// Раздача фронтенда через это же приложение — и локально, и на Vercel.
// (index.html и статика лежат в public/; на Vercel папка включается в бандл
//  функции через includeFiles в vercel.json.)
app.use(express.static(PUBLIC_DIR));
app.get(/^(?!\/api\/).*/, (_req, res) => res.sendFile(path.join(PUBLIC_DIR, 'index.html')));

// Локальный запуск: слушаем порт. На Vercel этот блок не выполняется.
if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  ensureSchema()
    .then(() => app.listen(PORT, () => console.log('Almaty Clinic → http://localhost:' + PORT)))
    .catch(e => { console.error('DB init failed:', e); process.exit(1); });
}

module.exports = app;

/* ---------- утилиты ---------- */
function int(v) { const n = parseInt(v, 10); return Number.isFinite(n) ? n : 0; }
function toClient(r) {
  return {
    id: r.id, ts: r.created_at, vuz: r.vuz, fio: r.fio, iin: r.iin, lang: r.lang,
    hadsAnx: r.hads_anx, hadsDep: r.hads_dep, hadsAnxI: r.hads_anx_i, hadsDepI: r.hads_dep_i,
    stage2: !!r.stage2, beck: r.stage2 ? r.beck : null, beckI: r.stage2 ? r.beck_i : null,
    beckItem9: r.beck_item9, flag: !!r.flag,
  };
}
