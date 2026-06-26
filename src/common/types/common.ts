export type AppScreen =
  | 'dashboard'
  | 'my-requests'
  | 'create-request'
  | 'request-detail'
  | 'pending-approval'
  | 'transport-assignment'
  | 'vehicle-schedule'
  | 'driver-trips'
  | 'completed-trips'
  | 'admin-settings'
  | 'user-role-setup';

export interface IAppUser {
  userId: string;
  userName: string;
  userEmail: string;
  role: string;
  department: string;
  isActive: boolean;
}

export interface ISelectOption {
  key: string;
  text: string;
}

export interface IServiceResult<T> {
  success: boolean;
  data?: T;
  error?: string;
}

export interface IPagedResult<T> {
  items: T[];
  total: number;
}
