export interface IVehicleBookingComment {
  ID: number;
  Title: string;
  BookingRequestID: number;
  RequestCode: string;
  Comment: string;
  CreatedByUserId: string;
  CreatedByName: string;
  CreatedByEmail: string;
  CommentDate: string;
  IsDeleted: boolean;
}
