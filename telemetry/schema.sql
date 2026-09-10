-- Har bir kompyuter uchun oxirgi holat (bitta qator) + tarix uchun ping vaqti.
CREATE TABLE IF NOT EXISTS clients (
  host        TEXT PRIMARY KEY,
  version     TEXT NOT NULL,
  first_seen  TEXT NOT NULL,
  last_seen   TEXT NOT NULL,
  ping_count  INTEGER NOT NULL DEFAULT 1
);

-- AI so'rovlarining kunlik sanog'i. Faqat SON saqlanadi — so'rov matnida
-- kontragent nomlari va to'lov izohlari bo'lgani uchun ular hech qayerga
-- yozilmaydi.
CREATE TABLE IF NOT EXISTS ai_usage (
  kun   TEXT NOT NULL,
  host  TEXT NOT NULL,
  soni  INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (kun, host)
);
