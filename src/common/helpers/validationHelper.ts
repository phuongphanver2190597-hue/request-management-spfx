export interface IValidationError {
  field: string;
  message: string;
}

export function validateBookingRequest(form: {
  pickupLocation: string;
  dropoffLocation: string;
  pickupDateTime: string;
  returnDateTime: string;
  isRoundTrip: boolean;
  numberOfPassengers: number;
  purpose: string;
  vehicleType: string;
  phoneNumber: string;
}): IValidationError[] {
  const errors: IValidationError[] = [];

  if (!form.pickupLocation.trim())
    errors.push({ field: 'pickupLocation', message: 'Vui lòng nhập điểm đón' });

  if (!form.dropoffLocation.trim())
    errors.push({ field: 'dropoffLocation', message: 'Vui lòng nhập điểm đến' });

  if (!form.pickupDateTime)
    errors.push({ field: 'pickupDateTime', message: 'Vui lòng chọn thời gian đón' });

  if (form.isRoundTrip && !form.returnDateTime)
    errors.push({ field: 'returnDateTime', message: 'Vui lòng chọn thời gian về (chuyến khứ hồi)' });

  if (form.pickupDateTime && form.returnDateTime) {
    if (new Date(form.returnDateTime) <= new Date(form.pickupDateTime))
      errors.push({ field: 'returnDateTime', message: 'Thời gian về phải sau thời gian đón' });
  }

  if (!form.numberOfPassengers || form.numberOfPassengers < 1)
    errors.push({ field: 'numberOfPassengers', message: 'Số hành khách phải ít nhất là 1' });

  if (!form.purpose.trim())
    errors.push({ field: 'purpose', message: 'Vui lòng nhập mục đích chuyến đi' });

  if (!form.vehicleType)
    errors.push({ field: 'vehicleType', message: 'Vui lòng chọn loại xe' });

  if (!form.phoneNumber.trim())
    errors.push({ field: 'phoneNumber', message: 'Vui lòng nhập số điện thoại liên hệ' });

  return errors;
}

export function validateAssignment(form: {
  assignedVehicleId: number;
  assignedDriverId: number;
}): IValidationError[] {
  const errors: IValidationError[] = [];
  if (!form.assignedVehicleId)
    errors.push({ field: 'assignedVehicleId', message: 'Vui lòng chọn xe' });
  if (!form.assignedDriverId)
    errors.push({ field: 'assignedDriverId', message: 'Vui lòng chọn tài xế' });
  return errors;
}

export function validateTripCompletion(form: {
  startOdometer: number;
  endOdometer: number;
  actualStartTime: string;
  actualEndTime: string;
}): IValidationError[] {
  const errors: IValidationError[] = [];
  if (!form.actualStartTime)
    errors.push({ field: 'actualStartTime', message: 'Vui lòng nhập thời gian bắt đầu' });
  if (!form.actualEndTime)
    errors.push({ field: 'actualEndTime', message: 'Vui lòng nhập thời gian kết thúc' });
  if (form.endOdometer <= form.startOdometer)
    errors.push({ field: 'endOdometer', message: 'ODO kết thúc phải lớn hơn ODO bắt đầu' });
  return errors;
}

export function getFieldError(errors: IValidationError[], field: string): string | undefined {
  const e = errors.find(x => x.field === field);
  return e ? e.message : undefined;
}
