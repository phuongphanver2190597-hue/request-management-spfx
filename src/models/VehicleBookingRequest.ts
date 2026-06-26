export interface IVehicleBookingRequest {
  ID: number;
  Title: string;
  RequestCode: string;
  RequesterId: string;
  RequesterName: string;
  RequesterEmail: string;
  Department: string;
  PhoneNumber: string;
  PickupLocation: string;
  DropoffLocation: string;
  PickupDateTime: string;
  ReturnDateTime: string;
  IsRoundTrip: boolean;
  NumberOfPassengers: number;
  Purpose: string;
  VehicleType: string;
  SpecialRequirement: string;
  Status: string;
  CurrentApproverId: string;
  CurrentOwnerId: string;
  AssignedVehicleId: number;
  AssignedVehiclePlate: string;
  AssignedDriverId: number;
  AssignedDriverName: string;
  AssignedDriverPhone: string;
  ActualStartTime: string;
  ActualEndTime: string;
  StartOdometer: number;
  EndOdometer: number;
  TotalDistance: number;
  CancelReason: string;
  AdminNote: string;
  SubmittedDate: string;
  CompletedDate: string;
  IsDeleted: boolean;
  Created: string;
  AuthorTitle?: string;
}

export interface ICreateBookingRequest {
  pickupLocation: string;
  dropoffLocation: string;
  pickupDateTime: string;
  returnDateTime: string;
  isRoundTrip: boolean;
  numberOfPassengers: number;
  purpose: string;
  vehicleType: string;
  specialRequirement: string;
  phoneNumber: string;
  department: string;
}

export interface IAssignmentPayload {
  assignedVehicleId: number;
  assignedVehiclePlate: string;
  assignedDriverId: number;
  assignedDriverName: string;
  assignedDriverPhone: string;
  adminNote: string;
}

export interface ITripCompletionPayload {
  actualStartTime: string;
  actualEndTime: string;
  startOdometer: number;
  endOdometer: number;
  totalDistance: number;
  adminNote: string;
}
