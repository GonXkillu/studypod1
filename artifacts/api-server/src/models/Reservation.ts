export interface ReservationData {
  id: number;
  userId: number;
  roomId: number;
  date: string;
  startTime: string;
  endTime: string;
  status: string;
  notes: string | null;
  createdAt: Date;
  updatedAt: Date;
}

export class Reservation {
  private readonly _id: number;
  private readonly _userId: number;
  private readonly _roomId: number;
  private readonly _date: string;
  private readonly _startTime: string;
  private readonly _endTime: string;
  private _status: string;
  private readonly _notes: string | null;
  private readonly _createdAt: Date;
  private readonly _updatedAt: Date;

  constructor(data: ReservationData) {
    this._id = data.id;
    this._userId = data.userId;
    this._roomId = data.roomId;
    this._date = data.date;
    this._startTime = data.startTime;
    this._endTime = data.endTime;
    this._status = data.status;
    this._notes = data.notes;
    this._createdAt = data.createdAt;
    this._updatedAt = data.updatedAt;
  }

  getId(): number { return this._id; }
  getUserId(): number { return this._userId; }
  getRoomId(): number { return this._roomId; }
  getDate(): string { return this._date; }
  getStartTime(): string { return this._startTime; }
  getEndTime(): string { return this._endTime; }
  getStatus(): string { return this._status; }
  getNotes(): string | null { return this._notes; }
  getCreatedAt(): Date { return this._createdAt; }
  isActive(): boolean { return this._status === "active"; }

  cancel(): void { this._status = "cancelled"; }

  toJSON(): Record<string, unknown> {
    return {
      id: this._id,
      userId: this._userId,
      roomId: this._roomId,
      date: this._date,
      startTime: this._startTime,
      endTime: this._endTime,
      status: this._status,
      notes: this._notes,
      createdAt: this._createdAt,
      updatedAt: this._updatedAt,
    };
  }
}
