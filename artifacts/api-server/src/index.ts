import path from "node:path";
import { fileURLToPath } from "node:url";
import { createApp } from "./app";
import { logger } from "./lib/logger";

const _dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.resolve(_dirname, "../public");

const app = createApp(publicDir);

const rawPort = process.env["PORT"];

if (!rawPort) {
  throw new Error(
    "PORT environment variable is required but was not provided.",
  );
}

const port = Number(rawPort);

if (Number.isNaN(port) || port <= 0) {
  throw new Error(`Invalid PORT value: "${rawPort}"`);
}

app.listen(port, (err) => {
  if (err) {
    logger.error({ err }, "Error listening on port");
    process.exit(1);
  }

  logger.info({ port }, "Server listening");
});