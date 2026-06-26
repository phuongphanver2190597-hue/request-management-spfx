import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SpService } from './spService';
import { LIST_NAMES } from '../common/constants/listNames';
import { IUserRole } from '../models/UserRole';
import { IAppUser } from '../common/types/common';
import { ROLE } from '../common/constants/roles';

// Bao gồm cả tên internal có thể có space encoding
const SELECT = 'ID,Title,UserId,UserName,UserEmail,Role,Department,IsActive';

function mapItem(i: Record<string, unknown>): IUserRole {
  // SharePoint Yes/No field có thể trả về: true/false (boolean), 1/0 (number), hoặc undefined nếu $select miss
  // Thử cả hai tên internal phổ biến
  const rawActive = i['IsActive'] !== undefined ? i['IsActive'] : i['Is_x0020_Active'];
  const isActive = rawActive === true || rawActive === 1 || String(rawActive).toLowerCase() === 'true' || String(rawActive) === '1';
  return {
    ID:         i.ID as number,
    Title:      (i.Title as string) || '',
    UserId:     (i.UserId as string) || '',
    UserName:   (i.UserName as string) || '',
    UserEmail:  (i.UserEmail as string) || '',
    Role:       (i.Role as string) || '',
    Department: (i.Department as string) || '',
    IsActive:   isActive,
  };
}

export class UserRoleService extends SpService {
  constructor(context: WebPartContext) { super(context); }

  async getAll(): Promise<IUserRole[]> {
    const url = this.itemsUrl(LIST_NAMES.USER_ROLE,
      `?$select=${SELECT}&$orderby=UserName asc&$top=500`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem).filter(u => u.IsActive !== false);
  }

  async getAllIncludingInactive(): Promise<IUserRole[]> {
    const url = this.itemsUrl(LIST_NAMES.USER_ROLE,
      `?$select=${SELECT}&$orderby=UserName asc&$top=500`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getByEmail(email: string): Promise<IUserRole | null> {
    const url = this.itemsUrl(LIST_NAMES.USER_ROLE,
      `?$select=${SELECT}&$filter=UserEmail eq '${email}'&$top=1`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.length > 0 ? mapItem(raw[0]) : null;
  }

  async getManagersForDepartment(department: string): Promise<IUserRole[]> {
    const filter = department
      ? `Role eq '${ROLE.MANAGER}' and Department eq '${department}'`
      : `Role eq '${ROLE.MANAGER}'`;
    const url = this.itemsUrl(LIST_NAMES.USER_ROLE,
      `?$select=${SELECT}&$filter=${filter}&$orderby=UserName asc&$top=50`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getAllManagers(): Promise<IUserRole[]> {
    const url = this.itemsUrl(LIST_NAMES.USER_ROLE,
      `?$select=${SELECT}&$filter=Role eq '${ROLE.MANAGER}'&$orderby=Department asc,UserName asc&$top=100`);
    const raw = await this.getItems<Record<string, unknown>>(url);
    return raw.map(mapItem);
  }

  async getCurrentUserRole(context: WebPartContext): Promise<IAppUser> {
    const email = context.pageContext.user.email;
    const displayName = context.pageContext.user.displayName;
    const record = await this.getByEmail(email);
    return {
      userId:     email,
      userName:   record?.UserName || displayName,
      userEmail:  email,
      role:       record?.Role || ROLE.REQUESTER,
      department: record?.Department || '',
      isActive:   record?.IsActive ?? true,
    };
  }

  async create(user: Omit<IUserRole, 'ID'>): Promise<void> {
    await this.postItem(this.itemsUrl(LIST_NAMES.USER_ROLE), {
      Title:      user.UserName,
      UserId:     user.UserId,
      UserName:   user.UserName,
      UserEmail:  user.UserEmail,
      Role:       user.Role,
      Department: user.Department,
      IsActive:   user.IsActive,
    });
  }

  async update(id: number, user: Partial<IUserRole>): Promise<void> {
    await this.updateItem(LIST_NAMES.USER_ROLE, id, user);
  }

  async deactivate(id: number): Promise<void> {
    await this.updateItem(LIST_NAMES.USER_ROLE, id, { IsActive: false });
  }
}
