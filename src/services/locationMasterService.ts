import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { ILocationMaster } from '../models/LocationMaster';

const SELECT = 'ID,Title,LocationCode,LocationName,Address,Province,Status,IsDeleted';

function mapItem(i: Record<string, unknown>): ILocationMaster {
  return {
    ID:           i.ID as number,
    Title:        (i.Title as string) || '',
    LocationCode: (i.LocationCode as string) || '',
    LocationName: (i.LocationName as string) || '',
    Address:      (i.Address as string) || '',
    Province:     (i.Province as string) || '',
    Status:       (i.Status as string) || '',
    IsDeleted:    (i.IsDeleted as boolean) || false,
  };
}

export class LocationMasterService extends SpService {
  constructor(context: WebPartContext) { super(context); }

  async getAll(): Promise<ILocationMaster[]> {
    const url = this.itemsUrl(LIST_NAMES.LOCATION_MASTER,
      `?$select=${SELECT}&$filter=IsDeleted eq false and Status eq 'ACTIVE'&$orderby=LocationName asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getAllIncludingInactive(): Promise<ILocationMaster[]> {
    const url = this.itemsUrl(LIST_NAMES.LOCATION_MASTER,
      `?$select=${SELECT}&$filter=IsDeleted eq false&$orderby=LocationName asc&$top=200`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async create(loc: Omit<ILocationMaster, 'ID'>): Promise<void> {
    await this.postItem(this.itemsUrl(LIST_NAMES.LOCATION_MASTER), {
      Title:        loc.LocationName,
      LocationCode: loc.LocationCode,
      LocationName: loc.LocationName,
      Address:      loc.Address,
      Province:     loc.Province,
      Status:       loc.Status,
      IsDeleted:    false,
    });
  }

  async update(id: number, loc: Partial<ILocationMaster>): Promise<void> {
    await this.updateItem(LIST_NAMES.LOCATION_MASTER, id, loc);
  }

  async softDelete(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.LOCATION_MASTER, id, { IsDeleted: true });
  }
}
