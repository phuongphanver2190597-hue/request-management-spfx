export const ACTION = {
  CREATE:             'CREATE',
  SAVE_DRAFT:         'SAVE_DRAFT',
  SUBMIT:             'SUBMIT',
  APPROVE_MANAGER:    'APPROVE_MANAGER',
  REJECT_MANAGER:     'REJECT_MANAGER',
  REQUEST_MORE_INFO:  'REQUEST_MORE_INFO',
  RESUBMIT:           'RESUBMIT',
  ASSIGN_VEHICLE:     'ASSIGN_VEHICLE',
  ASSIGN_DRIVER:      'ASSIGN_DRIVER',
  CONFIRM_BOOKING:    'CONFIRM_BOOKING',
  DRIVER_ACKNOWLEDGE: 'DRIVER_ACKNOWLEDGE',
  START_TRIP:         'START_TRIP',
  COMPLETE_TRIP:      'COMPLETE_TRIP',
  REJECT_NO_VEHICLE:  'REJECT_NO_VEHICLE',
  CANCEL:             'CANCEL',
} as const;

export type BookingAction = typeof ACTION[keyof typeof ACTION];

export const ACTION_LABEL: Record<string, string> = {
  [ACTION.CREATE]:             'Tạo mới',
  [ACTION.SAVE_DRAFT]:         'Lưu nháp',
  [ACTION.SUBMIT]:             'Gửi yêu cầu',
  [ACTION.APPROVE_MANAGER]:    'Trưởng phòng duyệt',
  [ACTION.REJECT_MANAGER]:     'Trưởng phòng từ chối',
  [ACTION.REQUEST_MORE_INFO]:  'Yêu cầu bổ sung thông tin',
  [ACTION.RESUBMIT]:           'Gửi lại',
  [ACTION.ASSIGN_VEHICLE]:     'Phân công xe',
  [ACTION.ASSIGN_DRIVER]:      'Phân công tài xế',
  [ACTION.CONFIRM_BOOKING]:    'Xác nhận đặt xe',
  [ACTION.DRIVER_ACKNOWLEDGE]: 'Tài xế xác nhận',
  [ACTION.START_TRIP]:         'Bắt đầu chuyến',
  [ACTION.COMPLETE_TRIP]:      'Hoàn thành chuyến',
  [ACTION.REJECT_NO_VEHICLE]:  'Từ chối - Không có xe',
  [ACTION.CANCEL]:             'Hủy yêu cầu',
};
