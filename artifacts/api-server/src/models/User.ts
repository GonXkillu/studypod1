export interface UserData {
  id: number;
  name: string;
  email: string;
  role: string;
  createdAt: Date;
}

export class User {
  private readonly _id: number;
  private readonly _name: string;
  private readonly _email: string;
  private readonly _role: string;
  private readonly _createdAt: Date;

  constructor(data: UserData) {
    this._id = data.id;
    this._name = data.name;
    this._email = data.email;
    this._role = data.role;
    this._createdAt = data.createdAt;
  }

  getId(): number { return this._id; }
  getName(): string { return this._name; }
  getEmail(): string { return this._email; }
  getRole(): string { return this._role; }
  getCreatedAt(): Date { return this._createdAt; }

  canManageRooms(): boolean { return false; }
  canViewAllReservations(): boolean { return false; }

  canCancelReservation(reservationUserId: number): boolean {
    return reservationUserId === this._id;
  }

  toPublicJSON(): Record<string, unknown> {
    return {
      id: this._id,
      name: this._name,
      email: this._email,
      role: this._role,
    };
  }
}

export class Student extends User {
  override canManageRooms(): boolean { return false; }
  override canViewAllReservations(): boolean { return false; }
}

export class Admin extends User {
  override canManageRooms(): boolean { return true; }
  override canViewAllReservations(): boolean { return true; }
  override canCancelReservation(_reservationUserId: number): boolean { return true; }
}

export function createUser(data: UserData): User {
  return data.role === "admin" ? new Admin(data) : new Student(data);
}
