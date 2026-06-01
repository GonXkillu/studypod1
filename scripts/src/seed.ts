import pg from "pg";
import bcrypt from "bcryptjs";

const { Pool } = pg;

if (!process.env.DATABASE_URL) throw new Error("DATABASE_URL is required");

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

async function seed() {
  console.log("Seeding database…");

  const adminHash = await bcrypt.hash("admin123!", 12);
  await pool.query(
    `INSERT INTO users (name, email, password_hash, role)
     VALUES ($1, $2, $3, $4)
     ON CONFLICT (email) DO NOTHING`,
    ["Admin User", "admin@studypod.com", adminHash, "admin"],
  );
  console.log("✓ Admin: admin@studypod.com / admin123!");

  const rooms = [
    ["Study Pod A", 2, "Quiet individual/pair study space"],
    ["Study Pod B", 2, "Quiet individual/pair study space"],
    ["Group Room 1", 6, "Collaborative group room with whiteboard"],
    ["Group Room 2", 8, "Large group room with projector"],
    ["Focus Room", 1, "Private silent study room"],
    ["Conference Room", 12, "Large meeting and presentation room"],
  ];

  for (const [name, capacity, description] of rooms) {
    await pool.query(
      `INSERT INTO rooms (name, capacity, description, is_active)
       VALUES ($1, $2, $3, true)
       ON CONFLICT DO NOTHING`,
      [name, capacity, description],
    );
  }
  console.log(`✓ Seeded ${rooms.length} rooms`);

  await pool.end();
  console.log("Done.");
}

seed().catch((err) => { console.error(err); process.exit(1); });
