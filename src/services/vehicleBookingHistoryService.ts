import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { IVehicleBookingHistory } from '../models/VehicleBookingHistory';
import { nowISO } from '../common/helpers/dateHelper';

function mapItem(i: Record<string, unknown>): IVehicleBookingHistory {
  return {
    ID:              i.ID as number,
    Title:           (i.Title as string) || '',
    BookingRequestID:(i.BookingRequestID as number) || 0,
    RequestCode:     (i.RequestCode as string) || '',
    Action:          (i.Action as string) || '',
    OldStatus:       (i.OldStatus as string) || '',
    NewStatus:       (i.NewStatus as string) || '',
    ActionByUserId:  (i.ActionByUserId as string) || '',
    ActionByName:    (i.ActionByName as string) || '',
    ActionByEmail:   (i.ActionByEmail as string) || '',
    ActionDate:      (i.ActionDate as string) || '',
    Note:            (i.Note as string) || '',
  };
}

export class VehicleBookingHistoryService extends SpService {
  constructor(context: WebPartContext) {
    super(context);
  }

  async getByRequestId(bookingRequestId: number): Promise<IVehicleBookingHistory[]> {
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_HISTORY,
      `?$select=ID,Title,BookingRequestID,RequestCode,Action,OldStatus,NewStatus,ActionByUserId,ActionByName,ActionByEmail,ActionDate,Note` +
      `&$filter=BookingRequestID eq ${bookingRequestId}&$orderby=ActionDate asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async log(
    bookingRequestId: number,
    requestCode: string,
    action: string,
    oldStatus: string,
    newStatus: string,
    user: { id: string; name: string; email: string },
    note: string,
  ): Promise<void> {
    const body = {
      Title:           `${requestCode}-${action}-${Date.now()}`,
      BookingRequestID: bookingRequestId,
      RequestCode:      requestCode,
      Action:           action,
      OldStatus:        oldStatus,
      NewStatus:        newStatus,
      ActionByUserId:   user.id,
      ActionByName:     user.name,
      ActionByEmail:    user.email,
      ActionDate:       nowISO(),
      Note:             note,
    };
    await this.postItem(this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_HISTORY), body);
  }
}
