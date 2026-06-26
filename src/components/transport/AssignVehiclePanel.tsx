import * as React from 'react';
import {
  Panel, PanelType, PrimaryButton, DefaultButton, Dropdown, IDropdownOption,
  TextField, Spinner, SpinnerSize, MessageBar, MessageBarType, mergeStyleSets, Stack,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { VehicleMasterService } from '../../services/vehicleMasterService';
import { DriverMasterService } from '../../services/driverMasterService';
import { IVehicleBookingRequest, IAssignmentPayload } from '../../models/VehicleBookingRequest';
import { IVehicleMaster } from '../../models/VehicleMaster';
import { IDriverMaster } from '../../models/DriverMaster';
import { validateAssignment, getFieldError, IValidationError } from '../../common/helpers/validationHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import { StatusBadge } from '../shared/StatusBadge';
import { formatDateTime } from '../../common/helpers/dateHelper';

const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  info: { marginBottom: 20, padding: '12px 14px', background: '#F6F8FA', borderRadius: 8 },
  infoRow: { display: 'grid' as const, gridTemplateColumns: '110px 1fr', gap: '4px 8px', fontSize: 13, marginBottom: 2 },
  infoKey: { color: TEXT_MUTE, fontWeight: 600, fontSize: 12 },
});

interface IAssignVehiclePanelProps {
  context: WebPartContext;
  user: IAppUser;
  request: IVehicleBookingRequest | null;
  isOpen: boolean;
  onDismiss: () => void;
  onAssigned: () => void;
}

interface IAssignVehiclePanelState {
  vehicles: IVehicleMaster[];
  drivers: IDriverMaster[];
  selectedVehicleId: number;
  selectedDriverId: number;
  adminNote: string;
  errors: IValidationError[];
  isLoading: boolean;
  isSaving: boolean;
  error: string | null;
}

export default class AssignVehiclePanel extends React.Component<IAssignVehiclePanelProps, IAssignVehiclePanelState> {
  private bookingSvc: VehicleBookingRequestService;
  private vehicleSvc: VehicleMasterService;
  private driverSvc: DriverMasterService;

  constructor(props: IAssignVehiclePanelProps) {
    super(props);
    this.state = {
      vehicles: [], drivers: [], selectedVehicleId: 0, selectedDriverId: 0,
      adminNote: '', errors: [], isLoading: false, isSaving: false, error: null,
    };
    this.bookingSvc = new VehicleBookingRequestService(props.context);
    this.vehicleSvc = new VehicleMasterService(props.context);
    this.driverSvc = new DriverMasterService(props.context);
  }

  public componentDidUpdate(prevProps: IAssignVehiclePanelProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this._loadMeta().catch(console.error);
      this.setState({ selectedVehicleId: 0, selectedDriverId: 0, adminNote: '', errors: [], error: null });
    }
  }

  private async _loadMeta(): Promise<void> {
    this.setState({ isLoading: true });
    try {
      const { request } = this.props;
      const [vehicles, drivers] = await Promise.all([
        this.vehicleSvc.getAvailable(request?.VehicleType),
        this.driverSvc.getAvailable(),
      ]);
      this.setState({ vehicles, drivers, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private async _assign(): Promise<void> {
    const { selectedVehicleId, selectedDriverId, adminNote, vehicles, drivers } = this.state;
    const errors = validateAssignment({ assignedVehicleId: selectedVehicleId, assignedDriverId: selectedDriverId });
    if (errors.length > 0) { this.setState({ errors }); return; }

    const vehicle = vehicles.find(v => v.ID === selectedVehicleId)!;
    const driver  = drivers.find(d => d.ID === selectedDriverId)!;
    const { request, user } = this.props;
    if (!request) return;

    const payload: IAssignmentPayload = {
      assignedVehicleId:   vehicle.ID,
      assignedVehiclePlate:vehicle.PlateNumber,
      assignedDriverId:    driver.ID,
      assignedDriverName:  driver.DriverName,
      assignedDriverPhone: driver.DriverPhone,
      adminNote,
    };

    this.setState({ isSaving: true, error: null });
    try {
      await this.bookingSvc.assignVehicleAndDriver(request.ID, request, payload, {
        id: user.userId, name: user.userName, email: user.userEmail,
      });
      await this.bookingSvc.confirmBooking(request.ID, { ...request, Status: 'VEHICLE_ASSIGNED' }, {
        id: user.userId, name: user.userName, email: user.userEmail,
      });
      this.setState({ isSaving: false });
      this.props.onAssigned();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  public render(): React.ReactElement {
    const { isOpen, onDismiss, request } = this.props;
    const { vehicles, drivers, selectedVehicleId, selectedDriverId, adminNote, errors, isLoading, isSaving, error } = this.state;

    const vehicleOptions: IDropdownOption[] = vehicles.map(v => ({
      key: v.ID, text: `${v.PlateNumber} - ${v.Brand} ${v.Model} (${v.VehicleType}, ${v.Capacity} chỗ, ODO: ${v.CurrentOdometer})`,
    }));
    const driverOptions: IDropdownOption[] = drivers.map(d => ({
      key: d.ID, text: `${d.DriverName} - ${d.DriverPhone} (${d.LicenseNumber})`,
    }));

    return (
      <Panel isOpen={isOpen} type={PanelType.medium}
        headerText="Phân công xe & Tài xế"
        onDismiss={onDismiss}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text={isSaving ? 'Đang lưu...' : 'Xác nhận phân công'} disabled={isSaving}
              onClick={() => this._assign().catch(console.error)} styles={{ root: { borderRadius: 8 } }} />
            <DefaultButton text="Hủy" onClick={onDismiss} styles={{ root: { borderRadius: 8 } }} />
          </Stack>
        )}
        isFooterAtBottom>

        {request && (
          <div className={styles.info}>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Mã YC:</span>
              <span style={{ fontWeight: 600 }}>{request.RequestCode}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Người YC:</span>
              <span>{request.RequesterName}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Điểm đón:</span>
              <span>{request.PickupLocation}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Điểm đến:</span>
              <span>{request.DropoffLocation}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Thời gian:</span>
              <span>{formatDateTime(request.PickupDateTime)}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Loại xe YC:</span>
              <span>{request.VehicleType}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoKey}>Trạng thái:</span>
              <StatusBadge status={request.Status} />
            </div>
          </div>
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 8, marginBottom: 14 } }}>
            {error}
          </MessageBar>
        )}

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải danh sách xe & tài xế..." />
        ) : (
          <Stack tokens={{ childrenGap: 18 }}>
            <Dropdown label="Chọn xe *"
              placeholder={vehicles.length === 0 ? '-- Không có xe phù hợp --' : '-- Chọn xe --'}
              options={vehicleOptions} selectedKey={selectedVehicleId || null}
              onChange={(_, o) => this.setState({ selectedVehicleId: Number(o?.key || 0), errors: [] })}
              errorMessage={getFieldError(errors, 'assignedVehicleId')}
              disabled={isSaving || vehicles.length === 0} />

            {vehicles.length === 0 && (
              <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 8 } }}>
                Không có xe khả dụng phù hợp loại xe yêu cầu. Có thể chọn xe loại khác sau khi trao đổi với người yêu cầu.
              </MessageBar>
            )}

            <Dropdown label="Chọn tài xế *"
              placeholder={drivers.length === 0 ? '-- Không có tài xế --' : '-- Chọn tài xế --'}
              options={driverOptions} selectedKey={selectedDriverId || null}
              onChange={(_, o) => this.setState({ selectedDriverId: Number(o?.key || 0), errors: [] })}
              errorMessage={getFieldError(errors, 'assignedDriverId')}
              disabled={isSaving || drivers.length === 0} />

            <TextField label="Ghi chú" multiline rows={3}
              placeholder="Ghi chú thêm về lịch trình, yêu cầu đặc biệt..."
              value={adminNote} onChange={(_, v) => this.setState({ adminNote: v || '' })}
              disabled={isSaving} />
          </Stack>
        )}
      </Panel>
    );
  }
}
