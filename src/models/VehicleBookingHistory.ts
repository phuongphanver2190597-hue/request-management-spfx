export interface IVehicleBookingHistory {
  ID: number;
  Title: string;
  BookingRequestID: number;
  RequestCode: string;
  Action: string;
  OldStatus: string;
  NewStatus: string;
  ActionByUserId: string;
  ActionByName: string;
  ActionByEmail: string;
  ActionDate: string;
  Note: string;
}
