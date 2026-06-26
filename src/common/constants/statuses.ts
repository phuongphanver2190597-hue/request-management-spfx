export const STATUS = {
  DRAFT:                        'DRAFT',
  SUBMITTED:                    'SUBMITTED',
  PENDING_MANAGER_APPROVAL:     'PENDING_MANAGER_APPROVAL',
  MANAGER_APPROVED:             'MANAGER_APPROVED',
  NEED_MORE_INFORMATION:        'NEED_MORE_INFORMATION',
  RESUBMITTED:                  'RESUBMITTED',
  PENDING_TRANSPORT_ASSIGNMENT: 'PENDING_TRANSPORT_ASSIGNMENT',
  VEHICLE_ASSIGNED:             'VEHICLE_ASSIGNED',
  CONFIRMED:                    'CONFIRMED',
  DRIVER_CONFIRMED:             'DRIVER_CONFIRMED',
  IN_PROGRESS:                  'IN_PROGRESS',
  COMPLETED:                    'COMPLETED',
  REJECTED:                     'REJECTED',
  REJECTED_NO_VEHICLE:          'REJECTED_NO_VEHICLE',
  CANCELLED:                    'CANCELLED',
} as const;

export type BookingStatus = typeof STATUS[keyof typeof STATUS];

export const STATUS_LABEL: Record<string, string> = {
  [STATUS.DRAFT]:                        'Nháp',
  [STATUS.SUBMITTED]:                    'Đã gửi',
  [STATUS.PENDING_MANAGER_APPROVAL]:     'Chờ trưởng phòng duyệt',
  [STATUS.MANAGER_APPROVED]:             'Trưởng phòng đã duyệt',
  [STATUS.NEED_MORE_INFORMATION]:        'Cần bổ sung thông tin',
  [STATUS.RESUBMITTED]:                  'Đã gửi lại',
  [STATUS.PENDING_TRANSPORT_ASSIGNMENT]: 'Chờ phân công xe',
  [STATUS.VEHICLE_ASSIGNED]:             'Đã phân công xe',
  [STATUS.CONFIRMED]:                    'Đã xác nhận',
  [STATUS.DRIVER_CONFIRMED]:             'Tài xế đã xác nhận',
  [STATUS.IN_PROGRESS]:                  'Đang thực hiện',
  [STATUS.COMPLETED]:                    'Hoàn thành',
  [STATUS.REJECTED]:                     'Từ chối',
  [STATUS.REJECTED_NO_VEHICLE]:          'Từ chối - Không có xe',
  [STATUS.CANCELLED]:                    'Đã hủy',
};

export const STATUS_COLOR: Record<string, { bg: string; color: string; dot: string }> = {
  [STATUS.DRAFT]:                        { bg: '#F3F4F6', color: '#6B7280', dot: '#9CA3AF' },
  [STATUS.SUBMITTED]:                    { bg: '#EFF6FF', color: '#1D4ED8', dot: '#3B82F6' },
  [STATUS.PENDING_MANAGER_APPROVAL]:     { bg: '#FFF8E6', color: '#B7791F', dot: '#F7B731' },
  [STATUS.MANAGER_APPROVED]:             { bg: '#ECFDF5', color: '#065F46', dot: '#10B981' },
  [STATUS.NEED_MORE_INFORMATION]:        { bg: '#FFF3CD', color: '#856404', dot: '#FFC107' },
  [STATUS.RESUBMITTED]:                  { bg: '#EFF6FF', color: '#1D4ED8', dot: '#60A5FA' },
  [STATUS.PENDING_TRANSPORT_ASSIGNMENT]: { bg: '#F5F3FF', color: '#5B21B6', dot: '#8B5CF6' },
  [STATUS.VEHICLE_ASSIGNED]:             { bg: '#F0FDF4', color: '#166534', dot: '#22C55E' },
  [STATUS.CONFIRMED]:                    { bg: '#ECFDF5', color: '#065F46', dot: '#10B981' },
  [STATUS.DRIVER_CONFIRMED]:             { bg: '#E0F2FE', color: '#0369A1', dot: '#0EA5E9' },
  [STATUS.IN_PROGRESS]:                  { bg: '#FFF8E6', color: '#B7791F', dot: '#F59E0B' },
  [STATUS.COMPLETED]:                    { bg: '#E8FFF2', color: '#276749', dot: '#20BF6B' },
  [STATUS.REJECTED]:                     { bg: '#FFF0F1', color: '#C53030', dot: '#FC5C65' },
  [STATUS.REJECTED_NO_VEHICLE]:          { bg: '#FFF0F1', color: '#C53030', dot: '#FC5C65' },
  [STATUS.CANCELLED]:                    { bg: '#F9FAFB', color: '#6B7280', dot: '#9CA3AF' },
};

export const VEHICLE_STATUS = {
  AVAILABLE:   'AVAILABLE',
  IN_USE:      'IN_USE',
  MAINTENANCE: 'MAINTENANCE',
  INACTIVE:    'INACTIVE',
} as const;

export const DRIVER_STATUS = {
  AVAILABLE: 'AVAILABLE',
  ON_TRIP:   'ON_TRIP',
  INACTIVE:  'INACTIVE',
} as const;

export const LOCATION_STATUS = {
  ACTIVE:   'ACTIVE',
  INACTIVE: 'INACTIVE',
} as const;

export const VEHICLE_TYPES = ['4-seat', '7-seat', '16-seat', 'Truck', 'Other'];
