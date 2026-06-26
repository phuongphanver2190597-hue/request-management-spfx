import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { VEHICLE_STATUS } from '../common/constants/statuses';
import { IVehicleMaster } from '../models/VehicleMaster';

const SELECT = 'ID,Title,VehicleCode,PlateNumber,VehicleType,Brand,Model,Capacity,Status,CurrentOdometer,Note,IsDeleted';

function mapItem(i: Record<string, unknown>): IVehicleMaster {
  return {
    ID:              i.ID as number,
    Title:           (i.Title as string) || '',
    VehicleCode:     (i.VehicleCode as string) || '',
    PlateNumber:     (i.PlateNumber as string) || '',
    VehicleType:     (i.VehicleType as string) || '',
    Brand:           (i.Brand as string) || '',
    Model:           (i.Model as string) || '',
    Capacity:        (i.Capacity as number) || 0,
    Status:          (i.Status as string) || '',
    CurrentOdometer: (i.CurrentOdometer as number) || 0,
    Note:            (i.Note as string) || '',
    IsDeleted:       (i.IsDeleted as boolean) || false,
  };
}

export class VehicleMasterService extends SpService {
  constructor(context: WebPartContext) { super(context); }

  async getAll(): Promise<IVehicleMaster[]> {
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_MASTER,
      `?$select=${SELECT}&$filter=IsDeleted eq false&$orderby=VehicleCode asc&$top=500`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getAvailable(vehicleType?: string): Promise<IVehicleMaster[]> {
    let filter = `IsDeleted eq false and Status eq '${VEHICLE_STATUS.AVAILABLE}'`;
    if (vehicleType) filter += ` and VehicleType eq '${vehicleType}'`;
    const url = this.itemsUrl(LIST_NAMES.VEHICLE_MASTER, `?$select=${SELECT}&$filter=${filter}&$orderby=VehicleCode asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async updateStatus(id: number, status: string, odometer?: number): Promise<void> {
    const body: Record<string, unknown> = { Status: status };
    if (odometer !== undefined) body.CurrentOdometer = odometer;
    await this.updateItem(LIST_NAMES.VEHICLE_MASTER, id, body);
  }

  async create(vehicle: Omit<IVehicleMaster, 'ID'>): Promise<void> {
    await this.postItem(this.itemsUrl(LIST_NAMES.VEHICLE_MASTER), {
      Title:           vehicle.PlateNumber,
      VehicleCode:     vehicle.VehicleCode,
      PlateNumber:     vehicle.PlateNumber,
      VehicleType:     vehicle.VehicleType,
      Brand:           vehicle.Brand,
      Model:           vehicle.Model,
      Capacity:        vehicle.Capacity,
      Status:          vehicle.Status,
      CurrentOdometer: vehicle.CurrentOdometer,
      Note:            vehicle.Note,
      IsDeleted:       false,
    });
  }

  async update(id: number, vehicle: Partial<IVehicleMaster>): Promise<void> {
    await this.updateItem(LIST_NAMES.VEHICLE_MASTER, id, vehicle);
  }

  async softDelete(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.VEHICLE_MASTER, id, { IsDeleted: true });
  }
}
