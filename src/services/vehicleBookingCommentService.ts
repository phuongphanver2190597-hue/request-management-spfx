import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { IVehicleBookingComment } from '../models/VehicleBookingComment';
import { nowISO } from '../common/helpers/dateHelper';

function mapItem(i: Record<string, unknown>): IVehicleBookingComment {
  return {
    ID:              i.ID as number,
    Title:           (i.Title as string) || '',
    BookingRequestID:(i.BookingRequestID as number) || 0,
    RequestCode:     (i.RequestCode as string) || '',
    Comment:         (i.Comment as string) || '',
    CreatedByUserId: (i.CreatedByUserId as string) || '',
    CreatedByName:   (i.CreatedByName as string) || '',
    CreatedByEmail:  (i.CreatedByEmail as string) || '',
    CommentDate:     (i.CommentDate as string) || '',
    IsDeleted:       (i.IsDeleted as boolean) || false,
  };
}

export class VehicleBookingCommentService extends SpService {
  constructor(context: WebPartContext) {
    super(context);
  }

  async getByRequestId(bookingRequestId: number): Promise<IVehicleBookingComment[]> {
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_COMMENT,
      `?$select=ID,Title,BookingRequestID,RequestCode,Comment,CreatedByUserId,CreatedByName,CreatedByEmail,CommentDate,IsDeleted` +
      `&$filter=BookingRequestID eq ${bookingRequestId} and IsDeleted eq false&$orderby=CommentDate asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async addComment(
    bookingRequestId: number,
    requestCode: string,
    comment: string,
    user: { id: string; name: string; email: string },
  ): Promise<IVehicleBookingComment> {
    const body = {
      Title:           `${requestCode}-CMT-${Date.now()}`,
      BookingRequestID: bookingRequestId,
      RequestCode:      requestCode,
      Comment:          comment,
      CreatedByUserId:  user.id,
      CreatedByName:    user.name,
      CreatedByEmail:   user.email,
      CommentDate:      nowISO(),
      IsDeleted:        false,
    };
    const created = await this.postItem<Record<string, unknown>>(
      this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_COMMENT), body
    );
    return mapItem(created);
  }

  async softDelete(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_COMMENT, id, { IsDeleted: true });
  }
}
