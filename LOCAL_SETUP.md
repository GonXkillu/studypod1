# StudyPod Scheduler — Local Setup & Vercel Deployment Guide

## Part 1: Run Locally

### Step 1 — Install PostgreSQL on your computer

**Windows:**
1. Download from https://www.postgresql.org/download/windows/
2. Run the installer. Remember the password you set for the `postgres` user.
3. PostgreSQL will be available at `localhost:5432`

**Mac:**
```bash
brew install postgresql@16
brew services start postgresql@16
```

**Already have PostgreSQL?** Skip to Step 2.

---

### Step 2 — Create the database

Open pgAdmin (installed with PostgreSQL on Windows) or your terminal:

**Using terminal / psql:**
```bash
psql -U postgres
```
Then type:
```sql
CREATE DATABASE studypod;
\q
```

**Using pgAdmin:**
1. Open pgAdmin → right-click "Databases" → Create → Database
2. Name it `studypod` → Save

---

### Step 3 — Run the setup SQL

This creates all tables and inserts the sample data (admin user + 6 rooms).

```bash
psql -U postgres -d studypod -f setup.sql
```

Or in pgAdmin: open the `studypod` database → Tools → Query Tool → open `setup.sql` → Run (▶)

You should see:
```
status          | users_count | rooms_count | reservations_count
Setup complete! |      1      |      6      |         0
```

---

### Step 4 — Install Node.js and pnpm

1. Download Node.js 20+ from https://nodejs.org
2. Install pnpm:
```bash
npm install -g pnpm
```

---

### Step 5 — Install project dependencies

From inside the project folder (where `pnpm-workspace.yaml` is):
```bash
pnpm install
```

---

### Step 6 — Create your .env file

Copy `.env.example` to `.env`:
```bash
cp .env.example .env
```

Edit `.env` and set:
```
DATABASE_URL=postgresql://postgres:YOURPASSWORD@localhost:5432/studypod
SESSION_SECRET=any-long-random-string-here-make-it-32-plus-chars
PORT=8080
NODE_ENV=development
```

Replace `YOURPASSWORD` with the postgres password you set during installation.

---

### Step 7 — Start the app

```bash
pnpm --filter @workspace/api-server run dev
```

Open http://localhost:8080 in your browser.

**Login:** `admin@studypod.com` / `admin123!`

---

## Part 2: Deploy to Vercel

Vercel needs a cloud PostgreSQL database (it can't connect to your laptop's local PostgreSQL). Use **Neon** — it's free and works perfectly with Vercel.

### Step 1 — Create a free Neon database

1. Go to https://neon.tech and sign up (free)
2. Create a new project → name it `studypod`
3. Copy the **Connection String** — it looks like:
   `postgresql://user:pass@ep-xxx.us-east-2.aws.neon.tech/neondb?sslmode=require`

### Step 2 — Run setup.sql on Neon

In the Neon dashboard → SQL Editor → paste the entire contents of `setup.sql` → Run.

This creates tables and seeds the admin user + rooms in your cloud database.

### Step 3 — Push code to GitHub

```bash
git init
git add .
git commit -m "StudyPod Scheduler"
git remote add origin https://github.com/YOUR_USERNAME/studypod-scheduler.git
git push -u origin main
```

### Step 4 — Deploy on Vercel

1. Go to https://vercel.com → New Project → Import your GitHub repo
2. Vercel will detect the `vercel.json` automatically
3. Before clicking Deploy, go to **Environment Variables** and add:
   - `DATABASE_URL` → your Neon connection string
   - `SESSION_SECRET` → a random 32+ character string
   - `NODE_ENV` → `production`
4. Click **Deploy**

Your app will be live at `https://your-app-name.vercel.app`

---

## Credentials

| Role | Email | Password |
|------|-------|----------|
| Admin | admin@studypod.com | admin123! |
| Student | Register a new account | anything (8+ chars) |

---

## Troubleshooting

**"password authentication failed"** — Wrong password in DATABASE_URL. Double-check the password you used when installing PostgreSQL.

**"database studypod does not exist"** — Run `CREATE DATABASE studypod;` in psql first.

**App starts but shows blank page** — Check the terminal for errors. Make sure `.env` file exists and has correct values.

**Vercel deploy fails** — Check build logs. Most common cause: DATABASE_URL not set in Vercel environment variables.
