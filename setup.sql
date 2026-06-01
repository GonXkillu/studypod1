-- ============================================================
-- StudyPod Scheduler — Complete Database Setup
-- Run this file in any PostgreSQL database to set up the app
-- Usage: psql -U your_user -d your_database -f setup.sql
-- ============================================================

-- Drop tables if re-running (safe reset)
DROP TABLE IF EXISTS reservations;
DROP TABLE IF EXISTS rooms;
DROP TABLE IF EXISTS users;
DROP TABLE IF EXISTS session;

-- ── Users ────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS users (
  id            SERIAL PRIMARY KEY,
  name          TEXT    NOT NULL,
  email         TEXT    NOT NULL UNIQUE,
  password_hash TEXT    NOT NULL,
  role          TEXT    NOT NULL DEFAULT 'student',
  created_at    TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- ── Rooms ────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS rooms (
  id          SERIAL PRIMARY KEY,
  name        TEXT    NOT NULL,
  capacity    INTEGER NOT NULL,
  description TEXT    NOT NULL DEFAULT '',
  is_active   BOOLEAN NOT NULL DEFAULT true,
  created_at  TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- ── Reservations ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS reservations (
  id         SERIAL PRIMARY KEY,
  user_id    INTEGER NOT NULL REFERENCES users(id),
  room_id    INTEGER NOT NULL REFERENCES rooms(id),
  date       TEXT    NOT NULL,
  start_time TEXT    NOT NULL,
  end_time   TEXT    NOT NULL,
  status     TEXT    NOT NULL DEFAULT 'active',
  notes      TEXT             DEFAULT '',
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- ── Session store (required for Vercel / production) ─────────
CREATE TABLE IF NOT EXISTS session (
  sid    VARCHAR    NOT NULL COLLATE "default",
  sess   JSON       NOT NULL,
  expire TIMESTAMP(6) NOT NULL,
  CONSTRAINT session_pkey PRIMARY KEY (sid) NOT DEFERRABLE INITIALLY IMMEDIATE
);
CREATE INDEX IF NOT EXISTS IDX_session_expire ON session (expire);

-- ── Seed: Admin user (password = admin123!) ───────────────────
-- Hash below is bcrypt cost-12 of "admin123!"
INSERT INTO users (name, email, password_hash, role) VALUES
  ('Admin User', 'admin@studypod.com',
   '$2b$12$WVWfCTCHASPwVLvSRQWmZ.wTOOcQVxjf7.hsXUxFdrsHDS6JnAxuC',
   'admin')
ON CONFLICT (email) DO NOTHING;

-- ── Seed: 6 Study Rooms ───────────────────────────────────────
INSERT INTO rooms (name, capacity, description, is_active) VALUES
  ('Study Pod A',    2,  'Quiet individual/pair study space',         true),
  ('Study Pod B',    2,  'Quiet individual/pair study space',         true),
  ('Group Room 1',   6,  'Collaborative group room with whiteboard',  true),
  ('Group Room 2',   8,  'Large group room with projector',           true),
  ('Focus Room',     1,  'Private silent study room',                 true),
  ('Conference Room',12, 'Large meeting and presentation room',       true)
ON CONFLICT DO NOTHING;

-- ── Done ──────────────────────────────────────────────────────
SELECT 'Setup complete!' AS status,
       (SELECT COUNT(*) FROM users)        AS users_count,
       (SELECT COUNT(*) FROM rooms)        AS rooms_count,
       (SELECT COUNT(*) FROM reservations) AS reservations_count;
