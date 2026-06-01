export interface RoomData {
  id: number;
  name: string;
  capacity: number;
  description: string;
  isActive: boolean;
  createdAt: Date;
}

export class Room {
  private readonly _id: number;
  private readonly _name: string;
  private readonly _capacity: number;
  private readonly _description: string;
  private _isActive: boolean;
  private readonly _createdAt: Date;

  constructor(data: RoomData) {
    this._id = data.id;
    this._name = data.name;
    this._capacity = data.capacity;
    this._description = data.description;
    this._isActive = data.isActive;
    this._createdAt = data.createdAt;
  }

  getId(): number { return this._id; }
  getName(): string { return this._name; }
  getCapacity(): number { return this._capacity; }
  getDescription(): string { return this._description; }
  getIsActive(): boolean { return this._isActive; }
  getCreatedAt(): Date { return this._createdAt; }

  deactivate(): void { this._isActive = false; }

  toJSON(): Record<string, unknown> {
    return {
      id: this._id,
      name: this._name,
      capacity: this._capacity,
      description: this._description,
      isActive: this._isActive,
      createdAt: this._createdAt,
    };
  }
}
