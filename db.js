'use strict';
/**
 * Слой базы данных на libSQL / Turso — через HTTP-клиент (@libsql/client/web).
 *
 * HTTP-клиент не использует нативных модулей и предназначен для serverless
 * (Vercel, Edge и т.п.), поэтому функция не падает при холодном старте.
 *
 * Требуются переменные окружения:
 *   TURSO_DATABASE_URL = libsql://<...>.turso.io   (или https://<...>.turso.io)
 *   TURSO_AUTH_TOKEN   = <токен>
 *
 * Если переменные не заданы — используется заглушка: страница откроется,
 * а запросы к /api вернут понятную ошибку (а не крах функции).
 */
const url = process.env.TURSO_DATABASE_URL;
const authToken = process.env.TURSO_AUTH_TOKEN;

let client;
if (url) {
  const { createClient } = require('@libsql/client/web');
  client = createClient({ url, authToken, intMode: 'number' });
} else {
  client = { execute: async () => { throw new Error('TURSO_DATABASE_URL не задан'); } };
}

let ready = null;
function ensureSchema() {
  if (!ready) {
    ready = client.execute(`
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
      )
    `).then(() => true);
  }
  return ready;
}

module.exports = { client, ensureSchema };
