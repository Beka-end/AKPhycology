'use strict';
/**
 * Слой базы данных на libSQL / Turso — через HTTP-клиент (@libsql/client/web).
 * Без нативных модулей, подходит для serverless (Vercel).
 *
 * Требуются переменные окружения:
 *   TURSO_DATABASE_URL = libsql://<...>.turso.io
 *   TURSO_AUTH_TOKEN   = <токен>
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
    ready = (async () => {
      await client.execute(`
        CREATE TABLE IF NOT EXISTS records (
          id          INTEGER PRIMARY KEY AUTOINCREMENT,
          created_at  TEXT NOT NULL,
          client_ts   TEXT,
          surname     TEXT,
          name        TEXT,
          patronymic  TEXT,
          fio         TEXT,
          sex         TEXT,
          age         INTEGER,
          phone       TEXT,
          vuz         TEXT,
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
      `);
      await client.execute(`
        CREATE TABLE IF NOT EXISTS institutions (
          id   INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT UNIQUE NOT NULL
        )
      `);
      await client.execute(`
        CREATE TABLE IF NOT EXISTS users (
          id         INTEGER PRIMARY KEY AUTOINCREMENT,
          login      TEXT UNIQUE NOT NULL,
          pass_hash  TEXT NOT NULL,
          pass_salt  TEXT NOT NULL,
          role       TEXT NOT NULL DEFAULT 'psy',
          created_at TEXT
        )
      `);
      await client.execute(`
        CREATE TABLE IF NOT EXISTS questions (
          id      INTEGER PRIMARY KEY AUTOINCREMENT,
          scale   TEXT NOT NULL,
          ord     INTEGER NOT NULL DEFAULT 0,
          text_ru TEXT,
          text_kz TEXT
        )
      `);
      await client.execute(`
        CREATE TABLE IF NOT EXISTS options (
          id          INTEGER PRIMARY KEY AUTOINCREMENT,
          question_id INTEGER NOT NULL,
          ord         INTEGER NOT NULL DEFAULT 0,
          text_ru     TEXT,
          text_kz     TEXT,
          score       INTEGER NOT NULL DEFAULT 0
        )
      `);
      // Миграции для ранее созданных таблиц (безопасно игнорируем, если колонка есть)
      const cols = ['surname TEXT','name TEXT','patronymic TEXT','fio TEXT',
                    'sex TEXT','age INTEGER','phone TEXT'];
      for (const c of cols) { try { await client.execute('ALTER TABLE records ADD COLUMN ' + c); } catch (e) {} }
      return true;
    })();
  }
  return ready;
}

module.exports = { client, ensureSchema };
