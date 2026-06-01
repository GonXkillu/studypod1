import type { Request, Response, NextFunction } from "express";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";
import { createUser } from "../models/User";

declare module "express-session" {
  interface SessionData {
    userId: number;
    role: string;
    name: string;
  }
}

export async function requireAuth(
  req: Request,
  res: Response,
  next: NextFunction,
): Promise<void> {
  if (!req.session?.userId) {
    res.status(401).json({ error: "Authentication required" });
    return;
  }
  next();
}

export async function requireAdmin(
  req: Request,
  res: Response,
  next: NextFunction,
): Promise<void> {
  if (!req.session?.userId) {
    res.status(401).json({ error: "Authentication required" });
    return;
  }

  const [userRow] = await db
    .select()
    .from(usersTable)
    .where(eq(usersTable.id, req.session.userId));

  if (!userRow) {
    res.status(401).json({ error: "User not found" });
    return;
  }

  const user = createUser(userRow);
  if (!user.canManageRooms()) {
    res.status(403).json({ error: "Administrator access required" });
    return;
  }

  next();
}
