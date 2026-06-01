import { Router } from "express";
import { db, roomsTable, usersTable, reservationsTable, insertRoomSchema } from "@workspace/db";
import { eq, and, ilike, or } from "drizzle-orm";
import { requireAdmin } from "../middleware/auth";
import { Room } from "../models/Room";

const router = Router();

router.get("/rooms", requireAdmin, async (_req, res) => {
  const rows = await db.select().from(roomsTable).orderBy(roomsTable.id);
  res.json({ rooms: rows.map((r) => new Room(r).toJSON()) });
});

router.post("/rooms", requireAdmin, async (req, res) => {
  const result = insertRoomSchema.safeParse(req.body);
  if (!result.success) {
    res.status(400).json({ error: result.error.issues[0]?.message ?? "Validation error" });
    return;
  }
  const [inserted] = await db.insert(roomsTable).values(result.data).returning();
  res.status(201).json({ room: new Room(inserted).toJSON() });
});

router.put("/rooms/:id", requireAdmin, async (req, res) => {
  const id = parseInt(req.params["id"] as string, 10);
  if (isNaN(id)) {
    res.status(400).json({ error: "Invalid room ID" });
    return;
  }
  const result = insertRoomSchema.partial().safeParse(req.body);
  if (!result.success) {
    res.status(400).json({ error: result.error.issues[0]?.message ?? "Validation error" });
    return;
  }
  const [updated] = await db
    .update(roomsTable)
    .set(result.data)
    .where(eq(roomsTable.id, id))
    .returning();
  if (!updated) {
    res.status(404).json({ error: "Room not found" });
    return;
  }
  res.json({ room: new Room(updated).toJSON() });
});

router.delete("/rooms/:id", requireAdmin, async (req, res) => {
  const id = parseInt(req.params["id"] as string, 10);
  if (isNaN(id)) {
    res.status(400).json({ error: "Invalid room ID" });
    return;
  }
  const [updated] = await db
    .update(roomsTable)
    .set({ isActive: false })
    .where(eq(roomsTable.id, id))
    .returning();
  if (!updated) {
    res.status(404).json({ error: "Room not found" });
    return;
  }
  res.json({ message: "Room deactivated successfully", room: new Room(updated).toJSON() });
});

router.get("/reservations", requireAdmin, async (req, res) => {
  const { search, status } = req.query;

  const baseQuery = db
    .select({
      id: reservationsTable.id,
      userId: reservationsTable.userId,
      roomId: reservationsTable.roomId,
      date: reservationsTable.date,
      startTime: reservationsTable.startTime,
      endTime: reservationsTable.endTime,
      status: reservationsTable.status,
      notes: reservationsTable.notes,
      createdAt: reservationsTable.createdAt,
      updatedAt: reservationsTable.updatedAt,
      userName: usersTable.name,
      userEmail: usersTable.email,
      roomName: roomsTable.name,
    })
    .from(reservationsTable)
    .leftJoin(usersTable, eq(reservationsTable.userId, usersTable.id))
    .leftJoin(roomsTable, eq(reservationsTable.roomId, roomsTable.id));

  const conditions = [];

  if (status && typeof status === "string" && status !== "all") {
    conditions.push(eq(reservationsTable.status, status));
  }

  if (search && typeof search === "string" && search.trim()) {
    const term = `%${search.trim()}%`;
    conditions.push(or(ilike(usersTable.name, term), ilike(roomsTable.name, term)));
  }

  const rows =
    conditions.length > 0
      ? await baseQuery
          .where(and(...conditions))
          .orderBy(reservationsTable.date, reservationsTable.startTime)
      : await baseQuery.orderBy(reservationsTable.date, reservationsTable.startTime);

  res.json({ reservations: rows });
});

router.delete("/reservations/:id", requireAdmin, async (req, res) => {
  const id = parseInt(req.params["id"] as string, 10);
  if (isNaN(id)) {
    res.status(400).json({ error: "Invalid reservation ID" });
    return;
  }
  const [updated] = await db
    .update(reservationsTable)
    .set({ status: "cancelled", updatedAt: new Date() })
    .where(eq(reservationsTable.id, id))
    .returning();
  if (!updated) {
    res.status(404).json({ error: "Reservation not found" });
    return;
  }
  res.json({ message: "Reservation cancelled successfully" });
});

export default router;
