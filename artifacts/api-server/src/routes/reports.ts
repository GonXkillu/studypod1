import { Router } from "express";
import { db, roomsTable, usersTable, reservationsTable } from "@workspace/db";
import { eq, and, gte, lte } from "drizzle-orm";
import { requireAdmin } from "../middleware/auth";

const router = Router();

router.get("/reservations", requireAdmin, async (req, res) => {
  const { startDate, endDate, status } = req.query;

  const conditions = [];

  if (startDate && typeof startDate === "string" && /^\d{4}-\d{2}-\d{2}$/.test(startDate)) {
    conditions.push(gte(reservationsTable.date, startDate));
  }
  if (endDate && typeof endDate === "string" && /^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
    conditions.push(lte(reservationsTable.date, endDate));
  }
  if (status && typeof status === "string" && status !== "all") {
    conditions.push(eq(reservationsTable.status, status));
  }

  const baseQuery = db
    .select({
      id: reservationsTable.id,
      date: reservationsTable.date,
      startTime: reservationsTable.startTime,
      endTime: reservationsTable.endTime,
      status: reservationsTable.status,
      notes: reservationsTable.notes,
      createdAt: reservationsTable.createdAt,
      userName: usersTable.name,
      userEmail: usersTable.email,
      roomName: roomsTable.name,
      roomCapacity: roomsTable.capacity,
    })
    .from(reservationsTable)
    .leftJoin(usersTable, eq(reservationsTable.userId, usersTable.id))
    .leftJoin(roomsTable, eq(reservationsTable.roomId, roomsTable.id));

  const rows =
    conditions.length > 0
      ? await baseQuery
          .where(and(...conditions))
          .orderBy(reservationsTable.date, reservationsTable.startTime)
      : await baseQuery.orderBy(reservationsTable.date, reservationsTable.startTime);

  res.json({
    title: "StudyPod Scheduler — Reservation Activity Report",
    generatedAt: new Date().toISOString(),
    filters: { startDate: startDate ?? null, endDate: endDate ?? null, status: status ?? "all" },
    totalRecords: rows.length,
    data: rows,
  });
});

export default router;
