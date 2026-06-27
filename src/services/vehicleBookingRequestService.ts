import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { STATUS } from '../common/constants/statuses';
import { ACTION } from '../common/constants/actions';
import { nowISO } from '../common/helpers/dateHelper';
import { IVehicleBookingRequest, ICreateBookingRequest, IAssignmentPayload, ITripCompletionPayload } from '../models/VehicleBookingRequest';
import { VehicleBookingHistoryService } from './vehicleBookingHistoryService';

const SELECT = 'ID,Title,RequestCode,RequesterId,RequesterName,RequesterEmail,Department,PhoneNumber,' +
  'PickupLocation,DropoffLocation,PickupDateTime,ReturnDateTime,IsRoundTrip,NumberOfPassengers,' +
  'Purpose,VehicleType,SpecialRequirement,Status,CurrentApproverId,CurrentOwnerId,' +
  'AssignedVehicleId,AssignedVehiclePlate,AssignedDriverId,AssignedDriverName,AssignedDriverPhone,' +
  'ActualStartTime,ActualEndTime,StartOdometer,EndOdometer,TotalDistance,' +
  'CancelReason,AdminNote,SubmittedDate,CompletedDate,IsDeleted,Created,Author/Title';
const EXPAND = 'Author';

function buildSelect(extra: string = ''): string {
  return `?$select=${SELECT}${extra}&$expand=${EXPAND}&$orderby=Created desc&$top=500`;
}

function mapItem(i: Record<string, unknown>): IVehicleBookingRequest {
  return {
    ID:                  i.ID as number,
    Title:               (i.Title as string) || '',
    RequestCode:         (i.RequestCode as string) || '',
    RequesterId:         (i.RequesterId as string) || '',
    RequesterName:       (i.RequesterName as string) || '',
    RequesterEmail:      (i.RequesterEmail as string) || '',
    Department:          (i.Department as string) || '',
    PhoneNumber:         (i.PhoneNumber as string) || '',
    PickupLocation:      (i.PickupLocation as string) || '',
    DropoffLocation:     (i.DropoffLocation as string) || '',
    PickupDateTime:      (i.PickupDateTime as string) || '',
    ReturnDateTime:      (i.ReturnDateTime as string) || '',
    IsRoundTrip:         (i.IsRoundTrip as boolean) || false,
    NumberOfPassengers:  (i.NumberOfPassengers as number) || 0,
    Purpose:             (i.Purpose as string) || '',
    VehicleType:         (i.VehicleType as string) || '',
    SpecialRequirement:  (i.SpecialRequirement as string) || '',
    Status:              (i.Status as string) || STATUS.DRAFT,
    CurrentApproverId:   (i.CurrentApproverId as string) || '',
    CurrentOwnerId:      (i.CurrentOwnerId as string) || '',
    AssignedVehicleId:   (i.AssignedVehicleId as number) || 0,
    AssignedVehiclePlate:(i.AssignedVehiclePlate as string) || '',
    AssignedDriverId:    (i.AssignedDriverId as number) || 0,
    AssignedDriverName:  (i.AssignedDriverName as string) || '',
    AssignedDriverPhone: (i.AssignedDriverPhone as string) || '',
    ActualStartTime:     (i.ActualStartTime as string) || '',
    ActualEndTime:       (i.ActualEndTime as string) || '',
    StartOdometer:       (i.StartOdometer as number) || 0,
    EndOdometer:         (i.EndOdometer as number) || 0,
    TotalDistance:       (i.TotalDistance as number) || 0,
    CancelReason:        (i.CancelReason as string) || '',
    AdminNote:           (i.AdminNote as string) || '',
    SubmittedDate:       (i.SubmittedDate as string) || '',
    CompletedDate:       (i.CompletedDate as string) || '',
    IsDeleted:           (i.IsDeleted as boolean) || false,
    Created:             (i.Created as string) || '',
    AuthorTitle:         ((i.Author as { Title?: string })?.Title) || '',
  };
}

function generateRequestCode(): string {
  const now = new Date();
  const yr  = now.getFullYear();
  const rnd = Math.floor(Math.random() * 90000) + 10000;
  return `VBR-${yr}-${rnd}`;
}

export class VehicleBookingRequestService extends SpService {
  private historyService: VehicleBookingHistoryService;

  constructor(context: WebPartContext) {
    super(context);
    this.historyService = new VehicleBookingHistoryService(context);
  }

  async getMyRequests(userEmail: string): Promise<IVehicleBookingRequest[]> {
    const filter = `&$filter=RequesterEmail eq '${userEmail}' and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, `${buildSelect(filter)}`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getPendingApproval(approverEmail: string): Promise<IVehicleBookingRequest[]> {
    const statuses = [STATUS.PENDING_MANAGER_APPROVAL, STATUS.RESUBMITTED].map(s => `Status eq '${s}'`).join(' or ');
    const filter = `&$filter=(${statuses}) and CurrentApproverId eq '${approverEmail}' and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getPendingTransportAssignment(): Promise<IVehicleBookingRequest[]> {
    const statuses = [
      STATUS.PENDING_TRANSPORT_ASSIGNMENT, STATUS.MANAGER_APPROVED,
      STATUS.VEHICLE_ASSIGNED, STATUS.CONFIRMED,
    ].map(s => `Status eq '${s}'`).join(' or ');
    const filter = `&$filter=(${statuses}) and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getDriverTrips(driverEmail: string): Promise<IVehicleBookingRequest[]> {
    const statuses = [STATUS.DRIVER_CONFIRMED, STATUS.IN_PROGRESS].map(s => `Status eq '${s}'`).join(' or ');
    const filter = `&$filter=(${statuses}) and AssignedDriverPhone ne null and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem).filter(r => r.AssignedDriverName !== '');
  }

  async getDriverTripsByEmail(driverName: string): Promise<IVehicleBookingRequest[]> {
    // Bao gồm VEHICLE_ASSIGNED và CONFIRMED để tài xế thấy chuyến cần xác nhận
    const statuses = [
      STATUS.VEHICLE_ASSIGNED, STATUS.CONFIRMED,
      STATUS.DRIVER_CONFIRMED, STATUS.IN_PROGRESS,
    ].map(s => `Status eq '${s}'`).join(' or ');
    // Filter trực tiếp theo AssignedDriverName ở server nếu có tên
    const nameFilter = driverName
      ? ` and AssignedDriverName eq '${driverName}'`
      : '';
    const filter = `&$filter=(${statuses})${nameFilter} and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem).filter(r => r.AssignedDriverName !== '');
  }

  async getCompletedTrips(): Promise<IVehicleBookingRequest[]> {
    const filter = `&$filter=Status eq '${STATUS.COMPLETED}' and IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getAllActive(): Promise<IVehicleBookingRequest[]> {
    const filter = `&$filter=IsDeleted eq false`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST, buildSelect(filter));
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getById(id: number): Promise<IVehicleBookingRequest> {
    const url = `${this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST)}(${id})?$select=${SELECT}&$expand=${EXPAND}`;
    const raw = await this.spHttpClient.get(url, (await import('@microsoft/sp-http')).SPHttpClient.configurations.v1);
    if (!raw.ok) throw new Error(`HTTP ${raw.status}`);
    const data = await raw.json() as Record<string, unknown>;
    return mapItem(data);
  }

  async createDraft(form: ICreateBookingRequest, user: { id: string; name: string; email: string }): Promise<IVehicleBookingRequest> {
    const code = generateRequestCode();
    const body = {
      Title:              code,
      RequestCode:        code,
      RequesterId:        user.id,
      RequesterName:      user.name,
      RequesterEmail:     user.email,
      Department:         form.department,
      PhoneNumber:        form.phoneNumber,
      PickupLocation:     form.pickupLocation,
      DropoffLocation:    form.dropoffLocation,
      PickupDateTime:     form.pickupDateTime || null,
      ReturnDateTime:     form.returnDateTime || null,
      IsRoundTrip:        form.isRoundTrip,
      NumberOfPassengers: form.numberOfPassengers,
      Purpose:            form.purpose,
      VehicleType:        form.vehicleType,
      SpecialRequirement: form.specialRequirement,
      Status:             STATUS.DRAFT,
      IsDeleted:          false,
    };
    const created = await this.postItem<Record<string, unknown>>(
      this.itemsUrl(LIST_NAMES.VEHICLE_BOOKING_REQUEST), body
    );
    const result = mapItem(created);
    await this.historyService.log(result.ID, code, ACTION.CREATE, '', STATUS.DRAFT, user, 'Tạo phiếu đặt xe');
    return result;
  }

  async submitRequest(
    id: number,
    request: IVehicleBookingRequest,
    user: { id: string; name: string; email: string },
    managerEmail: string,
    approvalLink?: string,
  ): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.PENDING_MANAGER_APPROVAL;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:            newStatus,
      CurrentApproverId: managerEmail,
      SubmittedDate:     nowISO(),
      ...(approvalLink ? { ApprovalLink: approvalLink } : {}),
    });
    await this.historyService.log(id, request.RequestCode, ACTION.SUBMIT, oldStatus, newStatus, user, '');
  }

  async saveDraft(id: number, form: ICreateBookingRequest): Promise<void> {
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      PickupLocation:     form.pickupLocation,
      DropoffLocation:    form.dropoffLocation,
      PickupDateTime:     form.pickupDateTime || null,
      ReturnDateTime:     form.returnDateTime || null,
      IsRoundTrip:        form.isRoundTrip,
      NumberOfPassengers: form.numberOfPassengers,
      Purpose:            form.purpose,
      VehicleType:        form.vehicleType,
      SpecialRequirement: form.specialRequirement,
      PhoneNumber:        form.phoneNumber,
      Department:         form.department,
    });
  }

  async approveByManager(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, note: string): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.PENDING_TRANSPORT_ASSIGNMENT;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:          newStatus,
      CurrentOwnerId:  user.email,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.APPROVE_MANAGER, oldStatus, newStatus, user, note);
  }

  async rejectByManager(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, note: string): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.REJECTED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:    newStatus,
      AdminNote: note,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.REJECT_MANAGER, oldStatus, newStatus, user, note);
  }

  async requestMoreInfo(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, note: string): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.NEED_MORE_INFORMATION;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:    newStatus,
      AdminNote: note,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.REQUEST_MORE_INFO, oldStatus, newStatus, user, note);
  }

  async resubmit(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.RESUBMITTED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, { Status: newStatus });
    await this.historyService.log(id, request.RequestCode, ACTION.RESUBMIT, oldStatus, newStatus, user, '');
  }

  async assignVehicleAndDriver(id: number, request: IVehicleBookingRequest, payload: IAssignmentPayload, user: { id: string; name: string; email: string }): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.VEHICLE_ASSIGNED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:              newStatus,
      AssignedVehicleId:   payload.assignedVehicleId,
      AssignedVehiclePlate:payload.assignedVehiclePlate,
      AssignedDriverId:    payload.assignedDriverId,
      AssignedDriverName:  payload.assignedDriverName,
      AssignedDriverPhone: payload.assignedDriverPhone,
      AdminNote:           payload.adminNote,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.ASSIGN_VEHICLE, oldStatus, newStatus, user,
      `Xe: ${payload.assignedVehiclePlate}, Tài xế: ${payload.assignedDriverName}`);
  }

  async confirmBooking(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.CONFIRMED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, { Status: newStatus });
    await this.historyService.log(id, request.RequestCode, ACTION.CONFIRM_BOOKING, oldStatus, newStatus, user, '');
  }

  async driverAcknowledge(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.DRIVER_CONFIRMED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, { Status: newStatus });
    await this.historyService.log(id, request.RequestCode, ACTION.DRIVER_ACKNOWLEDGE, oldStatus, newStatus, user, 'Tài xế xác nhận nhận chuyến');
  }

  async startTrip(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, startOdometer: number): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.IN_PROGRESS;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:          newStatus,
      ActualStartTime: nowISO(),
      StartOdometer:   startOdometer,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.START_TRIP, oldStatus, newStatus, user, `ODO: ${startOdometer}`);
  }

  async completeTrip(id: number, request: IVehicleBookingRequest, payload: ITripCompletionPayload, user: { id: string; name: string; email: string }): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.COMPLETED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:         newStatus,
      ActualEndTime:  payload.actualEndTime,
      EndOdometer:    payload.endOdometer,
      TotalDistance:  payload.totalDistance,
      CompletedDate:  nowISO(),
      AdminNote:      payload.adminNote,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.COMPLETE_TRIP, oldStatus, newStatus, user,
      `ODO: ${payload.startOdometer} → ${payload.endOdometer}, km: ${payload.totalDistance}`);
  }

  async rejectNoVehicle(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, note: string): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.REJECTED_NO_VEHICLE;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:    newStatus,
      AdminNote: note,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.REJECT_NO_VEHICLE, oldStatus, newStatus, user, note);
  }

  async cancelRequest(id: number, request: IVehicleBookingRequest, user: { id: string; name: string; email: string }, reason: string): Promise<void> {
    const oldStatus = request.Status;
    const newStatus = STATUS.CANCELLED;
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, {
      Status:       newStatus,
      CancelReason: reason,
    });
    await this.historyService.log(id, request.RequestCode, ACTION.CANCEL, oldStatus, newStatus, user, reason);
  }

  async softDelete(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.VEHICLE_BOOKING_REQUEST, id, { IsDeleted: true });
  }
}
