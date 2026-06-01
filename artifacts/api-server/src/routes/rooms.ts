import { Router } from "express";
import { db, roomsTable, reservationsTable } from "@workspace/db";
import { eq, and } from "drizzle-orm";
import { requireAuth } from "../middleware/auth";
import { Room } from "../models/Room";

const router = Router();

router.get("/", requireAuth, async (_req, res) => {
  const rows = await db
    .select()
    .from(roomsTable)
    .where(eq(roomsTable.isActive, true))
    .orderBy(roomsTable.name);
  res.json({ rooms: rows.map((r) => new Room(r).toJSON()) });
});

router.get("/:id/availability", requireAuth, async (req, res) => {
  const roomId = parseInt(req.params["id"] as string, 10);
  const { date } = req.query;

  if (!date || typeof date !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    res.status(400).json({ error: "A valid date (YYYY-MM-DD) is required" });
    return;
  }

  const reservations = await db
    .select()
    .from(reservationsTable)
    .where(
      and(
        eq(reservationsTable.roomId, roomId),
        eq(reservationsTable.date, date),
        eq(reservationsTable.status, "active"),
      ),
    )
    .orderBy(reservationsTable.startTime);

  res.json({ reservations });
});

export default router;
