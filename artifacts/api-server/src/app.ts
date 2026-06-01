import express, { type Express } from "express";
import cors from "cors";
import pinoHttp from "pino-http";
import session from "express-session";
import connectPgSimple from "connect-pg-simple";
import pg from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import router from "./routes";
import { logger } from "./lib/logger";

const _filename = fileURLToPath(import.meta.url);
const _dirname = path.dirname(_filename);

const app: Express = express();

// ── Session store: pg in production, memory in dev ───────────
const PgSession = connectPgSimple(session);

const sessionStore =
  process.env["NODE_ENV"] === "production"
    ? new PgSession({
        pool: new pg.Pool({ connectionString: process.env["DATABASE_URL"] }),
        tableName: "session",
        createTableIfMissing: false,
      })
    : undefined;

app.use(
  session({
    store: sessionStore,
    secret: process.env["SESSION_SECRET"] ?? "studypod-fallback-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: process.env["NODE_ENV"] === "production",
      sameSite: process.env["NODE_ENV"] === "production" ? "none" : "lax",
      maxAge: 24 * 60 * 60 * 1000,
    },
  }),
);

app.use(
  pinoHttp({
    logger,
    serializers: {
      req(req) {
        return { id: req.id, method: req.method, url: req.url?.split("?")[0] };
      },
      res(res) {
        return { statusCode: res.statusCode };
      },
    },
  }),
);

app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const publicDir = path.join(_dirname, "../public");
app.use(express.static(publicDir));

app.use("/api", router);

const serveHtml =
  (file: string) =>
  (_req: express.Request, res: express.Response): void => {
    res.sendFile(path.join(publicDir, file));
  };

app.get("/", serveHtml("index.html"));
app.get("/dashboard", serveHtml("dashboard.html"));
app.get("/my-reservations", serveHtml("my-reservations.html"));
app.get("/admin", serveHtml("admin.html"));
app.get("/reports", serveHtml("reports.html"));

export default app;
