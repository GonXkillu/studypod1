import { Router } from "express";
import bcrypt from "bcryptjs";
import { db, usersTable } from "@workspace/db";
import { eq } from "drizzle-orm";
import { createUser } from "../models/User";
import { z } from "zod";

const router = Router();

const registerSchema = z.object({
  name: z.string().min(2, "Name must be at least 2 characters"),
  email: z.string().email("Invalid email address"),
  password: z.string().min(8, "Password must be at least 8 characters"),
});

const loginSchema = z.object({
  email: z.string().email("Invalid email address"),
  password: z.string().min(1, "Password is required"),
});

router.post("/register", async (req, res) => {
  const result = registerSchema.safeParse(req.body);
  if (!result.success) {
    res.status(400).json({ error: result.error.issues[0]?.message ?? "Validation error" });
    return;
  }

  const { name, email, password } = result.data;

  const [existing] = await db
    .select()
    .from(usersTable)
    .where(eq(usersTable.email, email.toLowerCase()));

  if (existing) {
    res.status(409).json({ error: "An account with this email already exists" });
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  const [inserted] = await db
    .insert(usersTable)
    .values({ name, email: email.toLowerCase(), passwordHash, role: "student" })
    .returning();

  const user = createUser(inserted);
  req.session.userId = user.getId();
  req.session.role = user.getRole();
  req.session.name = user.getName();

  res.status(201).json({ user: user.toPublicJSON() });
});

router.post("/login", async (req, res) => {
  const result = loginSchema.safeParse(req.body);
  if (!result.success) {
    res.status(400).json({ error: result.error.issues[0]?.message ?? "Validation error" });
    return;
  }

  const { email, password } = result.data;
  const [userRow] = await db
    .select()
    .from(usersTable)
    .where(eq(usersTable.email, email.toLowerCase()));

  if (!userRow || !(await bcrypt.compare(password, userRow.passwordHash))) {
    res.status(401).json({ error: "Invalid email or password" });
    return;
  }

  const user = createUser(userRow);
  req.session.userId = user.getId();
  req.session.role = user.getRole();
  req.session.name = user.getName();

  res.json({ user: user.toPublicJSON() });
});

router.post("/logout", (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      res.status(500).json({ error: "Logout failed" });
      return;
    }
    res.clearCookie("connect.sid");
    res.json({ message: "Logged out" });
  });
});

router.get("/me", async (req, res) => {
  if (!req.session?.userId) {
    res.status(401).json({ error: "Not authenticated" });
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
  res.json({ user: createUser(userRow).toPublicJSON() });
});

export default router;
