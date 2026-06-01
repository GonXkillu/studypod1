# StudyPod Scheduler

A full-stack web application for online study room reservation, built for WGU D424 Software Engineering Capstone Task 3. Replaces paper-based scheduling with a digital system where students reserve rooms online.

**Student:** Amal Osman (ID: 012291062)

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run typecheck:libs` — rebuild shared libs (run after schema changes)
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- `pnpm --filter @workspace/scripts run seed` — seed admin user + 6 study rooms
- `pnpm --filter @workspace/scripts run generate-task3-doc` — regenerate Task 3 .docx
- Required env: `DATABASE_URL` — Postgres connection string
- Required env: `SESSION_SECRET` — secret for signing session cookies

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5 (serves both API and static HTML frontend)
- DB: PostgreSQL + Drizzle ORM
- Auth: bcryptjs (password hashing) + express-session (session management)
- Validation: Zod (server-side) + HTML5/JS (client-side)
- Frontend: Vanilla HTML/CSS/JavaScript (no framework)
- Build: esbuild (CJS bundle)

## Where things live

- `lib/db/src/schema/index.ts` — DB schema (users, rooms, reservations tables + Zod schemas)
- `artifacts/api-server/src/models/` — OOP classes (User, Student, Admin, Room, Reservation)
- `artifacts/api-server/src/middleware/auth.ts` — requireAuth and requireAdmin middleware
- `artifacts/api-server/src/routes/` — auth, rooms, reservations, admin, reports routes
- `artifacts/api-server/public/` — static HTML/CSS/JS frontend (5 pages)
- `scripts/src/seed.ts` — seeds admin user + 6 rooms
- `scripts/src/generate_task3_doc.ts` — generates Task 3 Word document
- `D424_Task3_AmalOsman_012291062.docx` — Task 3 submission document

## Architecture Decisions

- **Single Express server for both API + static files**: artifact serves at `/` path (not `/api`), Express handles both `/api/*` routes and serves HTML pages at root paths
- **bcryptjs over bcrypt**: bcrypt is a native addon that's externalized in esbuild; bcryptjs is pure JS and bundles fine
- **MemoryStore for sessions**: adequate for Replit demo; for Vercel serverless deployment, swap to connect-pg-simple
- **OOP design for rubric compliance**: User base class → Student + Admin subclasses with method overriding for polymorphism; private TypeScript fields for encapsulation
- **Drizzle ORM for SQL**: provides type-safe parameterized queries (SQL injection prevention); join syntax used for admin/report queries

## Product

- **Login/Register** — student accounts (self-register) and admin accounts (seeded)
- **Dashboard** — shows 6 active study rooms; click a room to check availability and book
- **My Reservations** — view all personal reservations; cancel active ones
- **Admin Panel** — manage rooms (add/edit/deactivate); view/search/cancel all reservations
- **Reports** — generate Reservation Activity Report with date range filter, all columns, timestamps
- **OOP**: User → Student + Admin (inheritance, polymorphism, encapsulation demonstrated)

## User Preferences

- Keep frontend as plain HTML/CSS/JS — no React or other frameworks
- Student ID: 012291062, Name: Amal Osman
- Demo credentials: admin@studypod.com / admin123!

## Gotchas

- Always run `pnpm run typecheck:libs` after changing `lib/db/src/schema/index.ts` before building api-server
- `pnpm --filter @workspace/db run push` must be run after any schema changes to update the database
- `req.params["id"] as string` required in Express 5 route handlers (typed as `string | string[]`)
- `import { z } from "zod"` not `zod/v4` in api-server routes — zod v4 sub-export path not needed here
- The artifact routes traffic from `/` (not `/api`) so the Express server handles all paths
- Static HTML files are at `artifacts/api-server/public/` and served relative to `dist/` (one level up)

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
