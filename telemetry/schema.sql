-- Har bir kompyuter uchun oxirgi holat (bitta qator) + tarix uchun ping vaqti.
CREATE TABLE IF NOT EXISTS clients (
  host        TEXT PRIMARY KEY,
  version     TEXT NOT NULL,
  first_seen  TEXT NOT NULL,
  last_seen   TEXT NOT NULL,
  ping_count  INTEGER NOT NULL DEFAULT 1
);
