import { Router } from "express";
import { db, roomsTable, usersTable, reservationsTable } from "@workspace/db";
import { eq, and, lt, gt } from "drizzle-orm";
import { requireAuth } from "../middleware/auth";
import { Reservation } from "../models/Reservation";
import { createUser } from "../models/User";
import { z } from "zod";

const router = Router();

const createReservationSchema = z.object({
  roomId: z.number().int().positive("Room ID must be a positive integer"),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "Date must be in YYYY-MM-DD format"),
  startTime: z.string().regex(/^\d{2}:\d{2}$/, "Start time must be in HH:MM format"),
  endTime: z.string().regex(/^\d{2}:\d{2}$/, "End time must be in HH:MM format"),
  notes: z.string().max(500, "Notes cannot exceed 500 characters").optional().default(""),
});

router.get("/my", requireAuth, async (req, res) => {
  const userId = req.session.userId!;

  const rows = await db
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
      roomName: roomsTable.name,
      roomCapacity: roomsTable.capacity,
    })
    .from(reservationsTable)
    .leftJoin(roomsTable, eq(reservationsTable.roomId, roomsTable.id))
    .where(eq(reservationsTable.userId, userId))
    .orderBy(reservationsTable.date, reservationsTable.startTime);

  res.json({ reservations: rows });
});

router.post("/", requireAuth, async (req, res) => {
  const userId = req.session.userId!;

  const result = createReservationSchema.safeParse(req.body);
  if (!result.success) {
    res.status(400).json({ error: result.error.issues[0]?.message ?? "Validation error" });
    return;
  }

  const { roomId, date, startTime, endTime, notes } = result.data;

  if (startTime >= endTime) {
    res.status(400).json({ error: "End time must be after start time" });
    return;
  }

  const today = new Date().toISOString().split("T")[0]!;
  if (date < today) {
    res.status(400).json({ error: "Cannot make reservations for past dates" });
    return;
  }

  const [room] = await db
    .select()
    .from(roomsTable)
    .where(and(eq(roomsTable.id, roomId), eq(roomsTable.isActive, true)));

  if (!room) {
    res.status(404).json({ error: "Room not found or is not available" });
    return;
  }

  const conflicts = await db
    .select()
    .from(reservationsTable)
    .where(
      and(
        eq(reservationsTable.roomId, roomId),
        eq(reservationsTable.date, date),
        eq(reservationsTable.status, "active"),
        lt(reservationsTable.startTime, endTime),
        gt(reservationsTable.endTime, startTime),
      ),
    );

  if (conflicts.length > 0) {
    res.status(409).json({
      error: "This time slot conflicts with an existing reservation. Please choose a different time.",
    });
    return;
  }

  const [inserted] = await db
    .insert(reservationsTable)
    .values({ userId, roomId, date, startTime, endTime, notes, status: "active" })
    .returning();

  res.status(201).json({ reservation: new Reservation(inserted).toJSON() });
});

router.delete("/:id", requireAuth, async (req, res) => {
  const userId = req.session.userId!;
  const reservationId = parseInt(req.params["id"] as string, 10);

  if (isNaN(reservationId)) {
    res.status(400).json({ error: "Invalid reservation ID" });
    return;
  }

  const [row] = await db
    .select()
    .from(reservationsTable)
    .where(eq(reservationsTable.id, reservationId));

  if (!row) {
    res.status(404).json({ error: "Reservation not found" });
    return;
  }

  const [userRow] = await db
    .select()
    .from(usersTable)
    .where(eq(usersTable.id, userId));

  if (!userRow) {
    res.status(401).json({ error: "User not found" });
    return;
  }

  const user = createUser(userRow);
  if (!user.canCancelReservation(row.userId)) {
    res.status(403).json({ error: "You can only cancel your own reservations" });
    return;
  }

  if (row.status !== "active") {
    res.status(400).json({ error: "This reservation is already cancelled" });
    return;
  }

  const [updated] = await db
    .update(reservationsTable)
    .set({ status: "cancelled", updatedAt: new Date() })
    .where(eq(reservationsTable.id, reservationId))
    .returning();

  res.json({ reservation: new Reservation(updated!).toJSON() });
});

export default router;
