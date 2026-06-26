export const ROLE = {
  REQUESTER:       'Requester',
  MANAGER:         'Manager',
  TRANSPORT_ADMIN: 'TransportAdmin',
  DRIVER:          'Driver',
  ADMIN:           'Admin',
} as const;

export type UserRole = typeof ROLE[keyof typeof ROLE];

export const ROLE_LABEL: Record<string, string> = {
  [ROLE.REQUESTER]:       'Người yêu cầu',
  [ROLE.MANAGER]:         'Trưởng phòng',
  [ROLE.TRANSPORT_ADMIN]: 'Quản lý vận tải',
  [ROLE.DRIVER]:          'Tài xế',
  [ROLE.ADMIN]:           'Quản trị viên',
};

export const MENU_BY_ROLE: Record<string, string[]> = {
  [ROLE.REQUESTER]: [
    'dashboard', 'my-requests', 'create-request', 'completed-trips',
  ],
  [ROLE.MANAGER]: [
    'dashboard', 'my-requests', 'create-request', 'pending-approval', 'completed-trips',
  ],
  [ROLE.TRANSPORT_ADMIN]: [
    'dashboard', 'transport-assignment', 'vehicle-schedule', 'completed-trips', 'admin-settings',
  ],
  [ROLE.DRIVER]: [
    'dashboard', 'driver-trips',
  ],
  [ROLE.ADMIN]: [
    'dashboard', 'my-requests', 'create-request', 'pending-approval',
    'transport-assignment', 'vehicle-schedule', 'driver-trips', 'completed-trips',
    'admin-settings', 'user-role-setup',
  ],
};
