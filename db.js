'use strict';
/**
 * Слой базы данных на libSQL / Turso.
 *
 *  - На Vercel (и любом serverless) задайте переменные окружения:
 *      TURSO_DATABASE_URL = libsql://<...>.turso.io
 *      TURSO_AUTH_TOKEN   = <токен>
 *    Тогда данные хранятся в облаке Turso и не теряются между запросами.
 *
 *  - Локально, если переменные не заданы, используется файл ./screening.db
 *    (для разработки и запуска на обычном сервере с постоянным диском).
 */
const { createClient } = require('@libsql/client');

const url = process.env.TURSO_DATABASE_URL || 'file:screening.db';
const authToken = process.env.TURSO_AUTH_TOKEN;

const client = createClient(
  authToken ? { url, authToken, intMode: 'number' } : { url, intMode: 'number' }
);

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
