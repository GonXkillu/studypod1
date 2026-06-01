import path from "node:path";
import { fileURLToPath } from "node:url";
import { createApp } from "./app";

const _dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(_dirname, "public");

export default createApp(publicDir);
