import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { DRIVER_STATUS } from '../common/constants/statuses';
import { IDriverMaster } from '../models/DriverMaster';

const SELECT = 'ID,Title,DriverCode,DriverName,DriverPhone,DriverEmail,Status,LicenseNumber,Note,IsDeleted';

function mapItem(i: Record<string, unknown>): IDriverMaster {
  return {
    ID:            i.ID as number,
    Title:         (i.Title as string) || '',
    DriverCode:    (i.DriverCode as string) || '',
    DriverName:    (i.DriverName as string) || '',
    DriverPhone:   (i.DriverPhone as string) || '',
    DriverEmail:   (i.DriverEmail as string) || '',
    Status:        (i.Status as string) || '',
    LicenseNumber: (i.LicenseNumber as string) || '',
    Note:          (i.Note as string) || '',
    IsDeleted:     (i.IsDeleted as boolean) || false,
  };
}

export class DriverMasterService extends SpService {
  constructor(context: WebPartContext) { super(context); }

  async getAll(): Promise<IDriverMaster[]> {
    const url = this.itemsUrl(LIST_NAMES.DRIVER_MASTER,
      `?$select=${SELECT}&$filter=IsDeleted eq false&$orderby=DriverCode asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getAvailable(): Promise<IDriverMaster[]> {
    const url = this.itemsUrl(LIST_NAMES.DRIVER_MASTER,
      `?$select=${SELECT}&$filter=IsDeleted eq false and Status eq '${DRIVER_STATUS.AVAILABLE}'&$orderby=DriverCode asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async updateStatus(id: number, status: string): Promise<void> {
    await this.updateItem(LIST_NAMES.DRIVER_MASTER, id, { Status: status });
  }

  async create(driver: Omit<IDriverMaster, 'ID'>): Promise<void> {
    await this.postItem(this.itemsUrl(LIST_NAMES.DRIVER_MASTER), {
      Title:         driver.DriverName,
      DriverCode:    driver.DriverCode,
      DriverName:    driver.DriverName,
      DriverPhone:   driver.DriverPhone,
      DriverEmail:   driver.DriverEmail,
      Status:        driver.Status,
      LicenseNumber: driver.LicenseNumber,
      Note:          driver.Note,
      IsDeleted:     false,
    });
  }

  async update(id: number, driver: Partial<IDriverMaster>): Promise<void> {
    await this.updateItem(LIST_NAMES.DRIVER_MASTER, id, driver);
  }

  async softDelete(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.DRIVER_MASTER, id, { IsDeleted: true });
  }
}
