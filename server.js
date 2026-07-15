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
try { require('dotenv').config(); } catch (_) {}
const { client, ensureSchema } = require('./db');
const FRONTEND_HTML = require('./frontend'); // страница, встроенная в бандл функции

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'almaty-clinic';
// Секрет для подписи сессионных токенов (стабилен между запусками serverless)
const SECRET = process.env.SESSION_SECRET || process.env.ADMIN_TOKEN
  || crypto.createHash('sha256').update('almaty-clinic-secret:' + ADMIN_PASSWORD).digest('hex');
const TOKEN_TTL_MS = 1000 * 60 * 60 * 12; // 12 часов

const LABEL = {
  hads_norm: 'норма', hads_sub: 'субклинически выраженная', hads_clin: 'клинически выраженная',
  b_none: 'отсутствие депрессии / ремиссия', b_border: 'пограничное состояние',
  b_mild: 'лёгкая депрессия', b_mod: 'депрессия средней тяжести', b_sev: 'тяжёлая депрессия',
};
const L = k => LABEL[k] || (k || '');

/* ---------- Пароли и токены ---------- */
function hashPassword(pw, salt) {
  salt = salt || crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(String(pw), salt, 32).toString('hex');
  return { hash, salt };
}
function verifyPassword(pw, hash, salt) {
  try {
    const h = crypto.scryptSync(String(pw), salt, 32).toString('hex');
    return crypto.timingSafeEqual(Buffer.from(h), Buffer.from(hash));
  } catch (e) { return false; }
}
function b64u(s) { return Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
function unb64u(s) { return Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString(); }
function signToken(payload) {
  const body = Object.assign({}, payload, { exp: Date.now() + TOKEN_TTL_MS });
  const data = b64u(JSON.stringify(body));
  const sig = crypto.createHmac('sha256', SECRET).update(data).digest('hex');
  return data + '.' + sig;
}
function verifyToken(token) {
  if (!token || token.indexOf('.') < 0) return null;
  const [data, sig] = token.split('.');
  const expect = crypto.createHmac('sha256', SECRET).update(data).digest('hex');
  if (sig.length !== expect.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expect))) return null;
  try { const body = JSON.parse(unb64u(data)); if (body.exp && Date.now() > body.exp) return null; return body; }
  catch (e) { return null; }
}

const app = express();
app.use(express.json({ limit: '256kb' }));

// Создаём учётку admin, если её ещё нет (пароль = ADMIN_PASSWORD из окружения)
let adminReady = null;
function ensureAdmin() {
  if (!adminReady) {
    adminReady = (async () => {
      const r = await client.execute({ sql: 'SELECT id FROM users WHERE login = ?', args: ['admin'] });
      if (!r.rows.length) {
        const { hash, salt } = hashPassword(ADMIN_PASSWORD);
        await client.execute({
          sql: 'INSERT INTO users (login,pass_hash,pass_salt,role,created_at) VALUES (?,?,?,?,?)',
          args: ['admin', hash, salt, 'admin', new Date().toISOString()],
        });
      }
      return true;
    })();
  }
  return adminReady;
}

// Автозагрузка стандартных вопросов в базу, если таблица вопросов пуста
let questionsReady = null;
function ensureQuestions() {
  if (!questionsReady) {
    questionsReady = (async () => {
      const cnt = (await client.execute('SELECT COUNT(*) c FROM questions')).rows[0].c;
      if (cnt > 0) return true;
      let data = {};
      try { data = require('./default-questions.json'); } catch (e) { console.error('no default-questions.json'); return true; }
      for (const scale of ['hads_dep', 'hads_anx', 'beck']) {
        const items = data[scale] || [];
        for (let i = 0; i < items.length; i++) await insertQuestion(scale, i, items[i]);
      }
      console.log('Стандартные вопросы загружены в базу');
      return true;
    })();
  }
  return questionsReady;
}

// Перед API-запросами гарантируем схему, учётку admin и наличие вопросов.
app.use('/api', async (_req, res, next) => {
  try { await ensureSchema(); await ensureAdmin(); await ensureQuestions(); next(); }
  catch (e) { console.error('DB init error:', e); res.status(500).json({ error: 'db_init' }); }
});

function auth(req, res, next) {
  const h = req.headers.authorization || '';
  const t = h.startsWith('Bearer ') ? h.slice(7) : (req.query.token || '');
  const u = verifyToken(t);
  if (!u) return res.status(401).json({ error: 'unauthorized' });
  req.user = u; next();
}
function adminOnly(req, res, next) {
  auth(req, res, () => {
    if (req.user.role !== 'admin') return res.status(403).json({ error: 'forbidden' });
    next();
  });
}

// Вход: логин + пароль → сессионный токен
app.post('/api/login', async (req, res) => {
  const login = ((req.body && req.body.login) || '').trim().toLowerCase();
  const password = (req.body && req.body.password) || '';
  if (!login || !password) return res.status(400).json({ error: 'missing' });
  try {
    const r = await client.execute({ sql: 'SELECT * FROM users WHERE login = ?', args: [login] });
    const u = r.rows[0];
    if (!u || !verifyPassword(password, u.pass_hash, u.pass_salt)) return res.status(401).json({ error: 'bad_credentials' });
    res.json({ token: signToken({ login: u.login, role: u.role }), login: u.login, role: u.role });
  } catch (e) { console.error('login:', e); res.status(500).json({ error: 'login_failed' }); }
});

// Смена собственного пароля
app.post('/api/account/password', auth, async (req, res) => {
  const oldP = (req.body && req.body.oldPassword) || '';
  const newP = (req.body && req.body.newPassword) || '';
  if (String(newP).length < 4) return res.status(400).json({ error: 'weak' });
  try {
    const r = await client.execute({ sql: 'SELECT * FROM users WHERE login = ?', args: [req.user.login] });
    const u = r.rows[0];
    if (!u || !verifyPassword(oldP, u.pass_hash, u.pass_salt)) return res.status(401).json({ error: 'bad_old' });
    const { hash, salt } = hashPassword(newP);
    await client.execute({ sql: 'UPDATE users SET pass_hash=?, pass_salt=? WHERE id=?', args: [hash, salt, u.id] });
    res.json({ ok: true });
  } catch (e) { console.error('pwd:', e); res.status(500).json({ error: 'pwd_failed' }); }
});

// Управление учётками (только admin)
app.get('/api/users', adminOnly, async (_req, res) => {
  try { const r = await client.execute('SELECT id, login, role, created_at FROM users ORDER BY id'); res.json(r.rows); }
  catch (e) { console.error('users list:', e); res.status(500).json({ error: 'users_list' }); }
});
app.post('/api/users', adminOnly, async (req, res) => {
  const login = ((req.body && req.body.login) || '').trim().toLowerCase();
  const password = (req.body && req.body.password) || '';
  const role = (req.body && req.body.role) === 'admin' ? 'admin' : 'psy';
  if (!/^[a-z0-9_.-]{3,32}$/.test(login)) return res.status(400).json({ error: 'bad_login' });
  if (String(password).length < 4) return res.status(400).json({ error: 'weak' });
  try {
    const { hash, salt } = hashPassword(password);
    await client.execute({
      sql: 'INSERT INTO users (login,pass_hash,pass_salt,role,created_at) VALUES (?,?,?,?,?)',
      args: [login, hash, salt, role, new Date().toISOString()],
    });
    res.json({ ok: true });
  } catch (e) {
    if (String(e).includes('UNIQUE')) return res.status(409).json({ error: 'exists' });
    console.error('user add:', e); res.status(500).json({ error: 'user_add' });
  }
});
app.delete('/api/users', adminOnly, async (req, res) => {
  const id = req.body && req.body.id;
  if (!id) return res.status(400).json({ error: 'no_id' });
  try {
    const r = await client.execute({ sql: 'SELECT login FROM users WHERE id = ?', args: [id] });
    const u = r.rows[0];
    if (!u) return res.status(404).json({ error: 'not_found' });
    if (u.login === 'admin') return res.status(400).json({ error: 'cannot_delete_admin' });
    if (u.login === req.user.login) return res.status(400).json({ error: 'cannot_delete_self' });
    await client.execute({ sql: 'DELETE FROM users WHERE id = ?', args: [id] });
    res.json({ ok: true });
  } catch (e) { console.error('user del:', e); res.status(500).json({ error: 'user_del' }); }
});

// Сохранение результата (пациент)
app.post('/api/records', async (req, res) => {
  try {
    const b = req.body || {};
    if (!/^\d{12}$/.test(String(b.iin || ''))) return res.status(400).json({ error: 'bad_iin' });
    if (!b.surname || !b.name) return res.status(400).json({ error: 'missing_fields' });
    const surname = String(b.surname).slice(0, 100);
    const name = String(b.name).slice(0, 100);
    const patronymic = b.patronymic ? String(b.patronymic).slice(0, 100) : '';
    const fio = [surname, name, patronymic].filter(Boolean).join(' ');
    const args = [
      new Date().toISOString(), b.ts || null,
      surname, name, patronymic, fio,
      b.sex === 'f' || b.sex === 'm' ? b.sex : null,
      b.age != null ? int(b.age) : null,
      b.phone ? String(b.phone).slice(0, 40) : null,
      b.vuz ? String(b.vuz).slice(0, 200) : null,
      String(b.iin),
      b.lang === 'kz' ? 'kz' : 'ru', int(b.hadsAnx), int(b.hadsDep),
      b.hadsAnxI || null, b.hadsDepI || null, b.stage2 ? 1 : 0,
      b.stage2 ? int(b.beck) : null, b.stage2 ? (b.beckI || null) : null,
      int(b.beckItem9), b.flag ? 1 : 0, JSON.stringify(b).slice(0, 20000),
    ];
    const r = await client.execute({
      sql: `INSERT INTO records
        (created_at,client_ts,surname,name,patronymic,fio,sex,age,phone,vuz,iin,lang,
         hads_anx,hads_dep,hads_anx_i,hads_dep_i,stage2,beck,beck_i,beck_item9,flag,raw)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      args,
    });
    res.json({ ok: true, id: Number(r.lastInsertRowid) });
  } catch (e) { console.error('save error:', e); res.status(500).json({ error: 'save_failed' }); }
});

// Справочник ВУЗов: список (публичный — для формы), добавление и удаление (психолог)
app.get('/api/institutions', async (_req, res) => {
  try {
    const r = await client.execute('SELECT id, name FROM institutions ORDER BY name COLLATE NOCASE');
    res.json(r.rows);
  } catch (e) { console.error('inst list:', e); res.status(500).json({ error: 'inst_list' }); }
});
app.post('/api/institutions', auth, async (req, res) => {
  const name = ((req.body && req.body.name) || '').trim();
  if (!name) return res.status(400).json({ error: 'empty' });
  try {
    await client.execute({ sql: 'INSERT OR IGNORE INTO institutions (name) VALUES (?)', args: [name.slice(0, 200)] });
    res.json({ ok: true });
  } catch (e) { console.error('inst add:', e); res.status(500).json({ error: 'inst_add' }); }
});
app.delete('/api/institutions', auth, async (req, res) => {
  const id = req.body && req.body.id;
  if (!id) return res.status(400).json({ error: 'no_id' });
  try {
    await client.execute({ sql: 'DELETE FROM institutions WHERE id = ?', args: [id] });
    res.json({ ok: true });
  } catch (e) { console.error('inst del:', e); res.status(500).json({ error: 'inst_del' }); }
});

/* ---------- Вопросы методик (просмотр — публично, правка — admin) ---------- */
async function insertQuestion(scale, ord, item) {
  const r = await client.execute({
    sql: 'INSERT INTO questions (scale,ord,text_ru,text_kz) VALUES (?,?,?,?)',
    args: [scale, ord, (item.text && item.text.ru) || '', (item.text && item.text.kz) || ''],
  });
  const qid = Number(r.lastInsertRowid);
  const opts = item.options || [];
  for (let j = 0; j < opts.length; j++) {
    const o = opts[j];
    await client.execute({
      sql: 'INSERT INTO options (question_id,ord,text_ru,text_kz,score) VALUES (?,?,?,?,?)',
      args: [qid, j, (o.text && o.text.ru) || '', (o.text && o.text.kz) || '', int(o.score)],
    });
  }
  return qid;
}

app.get('/api/questions', async (_req, res) => {
  try {
    const qs = (await client.execute('SELECT id,scale,ord,text_ru,text_kz FROM questions ORDER BY scale,ord,id')).rows;
    const ops = (await client.execute('SELECT id,question_id,ord,text_ru,text_kz,score FROM options ORDER BY question_id,ord,id')).rows;
    const byQ = {};
    ops.forEach(o => { (byQ[o.question_id] = byQ[o.question_id] || []).push({ text: { ru: o.text_ru || '', kz: o.text_kz || '' }, score: o.score }); });
    const out = { hads_dep: [], hads_anx: [], beck: [] };
    qs.forEach(q => { if (!out[q.scale]) out[q.scale] = []; out[q.scale].push({ id: q.id, text: { ru: q.text_ru || '', kz: q.text_kz || '' }, options: byQ[q.id] || [] }); });
    res.json(out);
  } catch (e) { console.error('q list:', e); res.status(500).json({ error: 'q_list' }); }
});

// Первичная загрузка стандартных вопросов (только если база пуста)
app.post('/api/questions/seed', adminOnly, async (req, res) => {
  try {
    const cnt = (await client.execute('SELECT COUNT(*) c FROM questions')).rows[0].c;
    if (cnt > 0) return res.json({ ok: true, already: true });
    const data = (req.body && req.body.data) || {};
    let n = 0;
    for (const scale of ['hads_dep', 'hads_anx', 'beck']) {
      const items = data[scale] || [];
      for (let i = 0; i < items.length; i++) { await insertQuestion(scale, i, items[i]); n++; }
    }
    res.json({ ok: true, inserted: n });
  } catch (e) { console.error('q seed:', e); res.status(500).json({ error: 'q_seed' }); }
});

app.post('/api/questions', adminOnly, async (req, res) => {
  const b = req.body || {};
  if (!['hads_dep', 'hads_anx', 'beck'].includes(b.scale)) return res.status(400).json({ error: 'bad_scale' });
  try {
    const m = (await client.execute({ sql: 'SELECT COALESCE(MAX(ord),-1) m FROM questions WHERE scale=?', args: [b.scale] })).rows[0].m;
    const id = await insertQuestion(b.scale, m + 1, { text: b.text || {}, options: b.options || [] });
    res.json({ ok: true, id });
  } catch (e) { console.error('q add:', e); res.status(500).json({ error: 'q_add' }); }
});

app.put('/api/questions', adminOnly, async (req, res) => {
  const b = req.body || {};
  if (!b.id) return res.status(400).json({ error: 'no_id' });
  try {
    await client.execute({ sql: 'UPDATE questions SET text_ru=?, text_kz=? WHERE id=?', args: [(b.text && b.text.ru) || '', (b.text && b.text.kz) || '', b.id] });
    await client.execute({ sql: 'DELETE FROM options WHERE question_id=?', args: [b.id] });
    const opts = b.options || [];
    for (let j = 0; j < opts.length; j++) {
      const o = opts[j];
      await client.execute({ sql: 'INSERT INTO options (question_id,ord,text_ru,text_kz,score) VALUES (?,?,?,?,?)', args: [b.id, j, (o.text && o.text.ru) || '', (o.text && o.text.kz) || '', int(o.score)] });
    }
    res.json({ ok: true });
  } catch (e) { console.error('q upd:', e); res.status(500).json({ error: 'q_upd' }); }
});

app.delete('/api/questions', adminOnly, async (req, res) => {
  const id = req.body && req.body.id;
  if (!id) return res.status(400).json({ error: 'no_id' });
  try {
    await client.execute({ sql: 'DELETE FROM options WHERE question_id=?', args: [id] });
    await client.execute({ sql: 'DELETE FROM questions WHERE id=?', args: [id] });
    res.json({ ok: true });
  } catch (e) { console.error('q del:', e); res.status(500).json({ error: 'q_del' }); }
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

// Очистка (только admin)
app.delete('/api/records', adminOnly, async (_req, res) => {
  try { await client.execute('DELETE FROM records'); res.json({ ok: true }); }
  catch (e) { console.error('clear error:', e); res.status(500).json({ error: 'clear_failed' }); }
});

// Выгрузка в Excel: лист «Результаты» (сводка) + лист «Ответы» (по каждому вопросу)
app.get('/api/export.xlsx', auth, async (_req, res) => {
  try {
    const r = await client.execute('SELECT * FROM records ORDER BY id ASC');
    const data = r.rows.map((row, i) => ({
      '№': i + 1,
      'Дата и время': new Date(row.created_at).toLocaleString('ru-RU'),
      'Фамилия': row.surname || '', 'Имя': row.name || '', 'Отчество': row.patronymic || '',
      'Пол': row.sex === 'f' ? 'жен' : (row.sex === 'm' ? 'муж' : ''),
      'Возраст': row.age != null ? row.age : '',
      'Телефон': row.phone || '', 'ИИН': row.iin, 'ВУЗ': row.vuz || '',
      'Язык': row.lang === 'kz' ? 'каз' : 'рус',
      'HADS Тревога (балл)': row.hads_anx, 'HADS Тревога': L(row.hads_anx_i),
      'HADS Депрессия (балл)': row.hads_dep, 'HADS Депрессия': L(row.hads_dep_i),
      'Направлен на 2 этап': row.stage2 ? 'да' : 'нет',
      'Бек (балл)': row.stage2 ? row.beck : '', 'Бек интерпретация': row.stage2 ? L(row.beck_i) : '',
      'Отметка риска (п.9 Бек)': row.beck_item9 > 0 ? 'да' : '',
      'Требует внимания': row.flag ? 'да' : '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:4},{wch:18},{wch:16},{wch:14},{wch:16},{wch:6},{wch:8},{wch:16},{wch:14},{wch:24},{wch:6},
                   {wch:14},{wch:22},{wch:16},{wch:22},{wch:16},{wch:9},{wch:24},{wch:16},{wch:14}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Результаты');

    // Лист «Ответы»: баллы по каждому пункту (из сохранённого raw)
    const ans = r.rows.map((row, i) => {
      let b = {};
      try { b = JSON.parse(row.raw || '{}'); } catch (e) {}
      const o = { '№': i + 1, 'ФИО': row.fio || '' };
      const dep = b.ansHadsDep || [], anx = b.ansHadsAnx || [], beck = b.ansBeck || [];
      for (let k = 0; k < 7; k++) o['HADS-Д' + (k + 1)] = dep[k] != null ? dep[k] : '';
      for (let k = 0; k < 7; k++) o['HADS-Т' + (k + 1)] = anx[k] != null ? anx[k] : '';
      for (let k = 0; k < 21; k++) o['Бек-' + (k + 1)] = row.stage2 && beck[k] != null ? beck[k] : '';
      return o;
    });
    const ws2 = XLSX.utils.json_to_sheet(ans);
    ws2['!cols'] = [{wch:4},{wch:24}].concat(Array(35).fill({wch:8}));
    XLSX.utils.book_append_sheet(wb, ws2, 'Ответы');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const stamp = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="almaty_clinic_skrining_${stamp}.xlsx"`);
    res.send(buf);
  } catch (e) { console.error('export error:', e); res.status(500).json({ error: 'export_failed' }); }
});

// Раздача фронтенда: страница встроена в код (frontend.js), поэтому работает
// и локально, и на Vercel без обращения к файловой системе.
app.get(/^(?!\/api\/).*/, (_req, res) => res.type('html').send(FRONTEND_HTML));

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
    id: r.id, ts: r.created_at,
    surname: r.surname, name: r.name, patronymic: r.patronymic, fio: r.fio,
    sex: r.sex, age: r.age, phone: r.phone,
    vuz: r.vuz, iin: r.iin, lang: r.lang,
    hadsAnx: r.hads_anx, hadsDep: r.hads_dep, hadsAnxI: r.hads_anx_i, hadsDepI: r.hads_dep_i,
    stage2: !!r.stage2, beck: r.stage2 ? r.beck : null, beckI: r.stage2 ? r.beck_i : null,
    beckItem9: r.beck_item9, flag: !!r.flag,
  };
}
