/**
 * Almaty Clinic — сервер психологического скрининга (HADS → Бек)
 * Express + SQLite (better-sqlite3) + выгрузка в Excel (SheetJS).
 *
 * Запуск:
 *   npm install
 *   cp .env.example .env   # и задать пароли
 *   npm start
 *
 * Эндпоинты:
 *   POST /api/login          { password }            → { token }
 *   POST /api/records        { ...результат теста }   → { ok, id }     (студент)
 *   GET  /api/records        [Bearer]                → [ записи ]      (психолог)
 *   GET  /api/records/stats  [Bearer]                → { total, ... }
 *   GET  /api/export.xlsx    [Bearer]                → файл .xlsx
 *   DELETE /api/records      [Bearer]                → { ok }
 */
'use strict';
const express = require('express');
const Database = require('better-sqlite3');
const XLSX = require('xlsx');
const path = require('path');
const crypto = require('crypto');
try { require('dotenv').config(); } catch (_) {}

const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'almaty-clinic';
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || crypto.randomBytes(24).toString('hex');
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'screening.db');

/* ---------- База данных ---------- */
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.exec(`
CREATE TABLE IF NOT EXISTS records (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at  TEXT NOT NULL,
  client_ts   TEXT,
  vuz         TEXT,
  fio         TEXT,
  iin         TEXT,
  lang        TEXT,
  hads_anx    INTEGER,
  hads_dep    INTEGER,
  hads_anx_i  TEXT,
  hads_dep_i  TEXT,
  stage2      INTEGER,
  beck        INTEGER,
  beck_i      TEXT,
  beck_item9  INTEGER,
  flag        INTEGER,
  raw         TEXT
);
CREATE INDEX IF NOT EXISTS idx_records_created ON records(created_at);
`);

/* Словарь интерпретаций для Excel (русский) */
const LABEL = {
  hads_norm:'норма', hads_sub:'субклинически выраженная', hads_clin:'клинически выраженная',
  b_none:'отсутствие депрессии / ремиссия', b_border:'пограничное состояние',
  b_mild:'лёгкая депрессия', b_mod:'депрессия средней тяжести', b_sev:'тяжёлая депрессия',
};
const L = k => LABEL[k] || (k || '');

/* ---------- Приложение ---------- */
const app = express();
app.use(express.json({ limit: '256kb' }));

// Простая защита раздела психолога токеном
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
const insert = db.prepare(`INSERT INTO records
  (created_at, client_ts, vuz, fio, iin, lang, hads_anx, hads_dep, hads_anx_i, hads_dep_i,
   stage2, beck, beck_i, beck_item9, flag, raw)
  VALUES (@created_at,@client_ts,@vuz,@fio,@iin,@lang,@hads_anx,@hads_dep,@hads_anx_i,@hads_dep_i,
   @stage2,@beck,@beck_i,@beck_item9,@flag,@raw)`);

app.post('/api/records', (req, res) => {
  const b = req.body || {};
  if (!/^\d{12}$/.test(String(b.iin || ''))) return res.status(400).json({ error: 'bad_iin' });
  if (!b.fio || !b.vuz) return res.status(400).json({ error: 'missing_fields' });
  const row = {
    created_at: new Date().toISOString(),
    client_ts: b.ts || null,
    vuz: String(b.vuz).slice(0, 200),
    fio: String(b.fio).slice(0, 200),
    iin: String(b.iin),
    lang: b.lang === 'kz' ? 'kz' : 'ru',
    hads_anx: int(b.hadsAnx), hads_dep: int(b.hadsDep),
    hads_anx_i: b.hadsAnxI || null, hads_dep_i: b.hadsDepI || null,
    stage2: b.stage2 ? 1 : 0,
    beck: b.stage2 ? int(b.beck) : null,
    beck_i: b.stage2 ? (b.beckI || null) : null,
    beck_item9: int(b.beckItem9),
    flag: b.flag ? 1 : 0,
    raw: JSON.stringify(b).slice(0, 20000),
  };
  const info = insert.run(row);
  res.json({ ok: true, id: info.lastInsertRowid });
});

// Список записей (психолог) — в camelCase, как ждёт фронтенд
app.get('/api/records', auth, (_req, res) => {
  const rows = db.prepare('SELECT * FROM records ORDER BY id DESC').all();
  res.json(rows.map(toClient));
});

// Статистика
app.get('/api/records/stats', auth, (_req, res) => {
  const total = db.prepare('SELECT COUNT(*) c FROM records').get().c;
  const stage2 = db.prepare('SELECT COUNT(*) c FROM records WHERE stage2=1').get().c;
  const flagged = db.prepare('SELECT COUNT(*) c FROM records WHERE flag=1').get().c;
  res.json({ total, stage2, flagged });
});

// Очистка
app.delete('/api/records', auth, (_req, res) => {
  db.prepare('DELETE FROM records').run();
  res.json({ ok: true });
});

// Выгрузка в Excel
app.get('/api/export.xlsx', auth, (_req, res) => {
  const rows = db.prepare('SELECT * FROM records ORDER BY id ASC').all();
  const data = rows.map((r, i) => ({
    '№': i + 1,
    'Дата и время': new Date(r.created_at).toLocaleString('ru-RU'),
    'ВУЗ': r.vuz, 'ФИО': r.fio, 'ИИН': r.iin,
    'Язык': r.lang === 'kz' ? 'каз' : 'рус',
    'HADS Тревога (балл)': r.hads_anx, 'HADS Тревога': L(r.hads_anx_i),
    'HADS Депрессия (балл)': r.hads_dep, 'HADS Депрессия': L(r.hads_dep_i),
    'Направлен на 2 этап': r.stage2 ? 'да' : 'нет',
    'Бек (балл)': r.stage2 ? r.beck : '', 'Бек интерпретация': r.stage2 ? L(r.beck_i) : '',
    'Отметка риска (п.9 Бек)': r.beck_item9 > 0 ? 'да' : '',
    'Требует внимания': r.flag ? 'да' : '',
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
});

/* ---------- Статика (фронтенд) ---------- */
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => {
  console.log(`Almaty Clinic screening → http://localhost:${PORT}`);
  if (!process.env.ADMIN_TOKEN) console.log('ADMIN_TOKEN (сгенерирован): ' + ADMIN_TOKEN);
});

/* ---------- Утилиты ---------- */
function int(v) { const n = parseInt(v, 10); return Number.isFinite(n) ? n : 0; }
function toClient(r) {
  return {
    id: r.id, ts: r.created_at, vuz: r.vuz, fio: r.fio, iin: r.iin, lang: r.lang,
    hadsAnx: r.hads_anx, hadsDep: r.hads_dep, hadsAnxI: r.hads_anx_i, hadsDepI: r.hads_dep_i,
    stage2: !!r.stage2, beck: r.stage2 ? r.beck : null, beckI: r.stage2 ? r.beck_i : null,
    beckItem9: r.beck_item9, flag: !!r.flag,
  };
}
