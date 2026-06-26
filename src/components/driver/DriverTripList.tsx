import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, TextField, mergeStyleSets,
  Dialog, DialogType, DialogFooter, FontIcon, Text,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { STATUS } from '../../common/constants/statuses';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime, nowISO } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const PRIMARY = '#0078D4';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  grid: { display: 'grid' as const, gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 16 },
  card: {
    background: WHITE, borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden',
  },
  cardHead: {
    padding: '12px 16px', background: '#F6F8FA',
    borderBottom: `1px solid ${BORDER}`,
    display: 'flex' as const, justifyContent: 'space-between' as const, alignItems: 'center' as const,
  },
  cardCode: { fontWeight: 700, fontSize: 14, color: PRIMARY },
  cardBody: { padding: '14px 16px' },
  row: {
    display: 'flex' as const, fontSize: 13, marginBottom: 6, gap: 8,
  },
  label: { color: TEXT_MUTE, fontWeight: 600, fontSize: 12, flexShrink: 0, minWidth: 90 },
  actions: { padding: '12px 16px', borderTop: `1px solid ${BORDER}`, display: 'flex' as const, gap: 8, flexWrap: 'wrap' as const },
  empty: { padding: '60px 0', textAlign: 'center' as const, color: TEXT_MUTE },
});

interface IDriverTripListProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IDriverTripListState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  activeDialog: 'start' | 'complete' | null;
  dialogRequest: IVehicleBookingRequest | null;
  startOdometer: number;
  endOdometer: number;
  adminNote: string;
  isActioning: boolean;
}

export default class DriverTripList extends React.Component<IDriverTripListProps, IDriverTripListState> {
  private svc: VehicleBookingRequestService;

  constructor(props: IDriverTripListProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null,
      activeDialog: null, dialogRequest: null,
      startOdometer: 0, endOdometer: 0, adminNote: '',
      isActioning: false,
    };
    this.svc = new VehicleBookingRequestService(props.context);
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
  }

  private async _load(): Promise<void> {
    this.setState({ isLoading: true, error: null });
    try {
      const driverName = this.props.user.userName;
      const requests = await this.svc.getDriverTripsByEmail(driverName);
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _user() {
    const { user } = this.props;
    return { id: user.userId, name: user.userName, email: user.userEmail };
  }

  private async _acknowledge(request: IVehicleBookingRequest): Promise<void> {
    this.setState({ isActioning: true });
    try {
      await this.svc.driverAcknowledge(request.ID, request, this._user());
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    } finally {
      this.setState({ isActioning: false });
    }
  }

  private async _startTrip(): Promise<void> {
    const { dialogRequest, startOdometer } = this.state;
    if (!dialogRequest) return;
    this.setState({ isActioning: true });
    try {
      await this.svc.startTrip(dialogRequest.ID, dialogRequest, this._user(), startOdometer);
      this.setState({ activeDialog: null, dialogRequest: null });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    } finally {
      this.setState({ isActioning: false });
    }
  }

  private async _completeTrip(): Promise<void> {
    const { dialogRequest, startOdometer, endOdometer, adminNote } = this.state;
    if (!dialogRequest) return;
    const total = Math.max(0, endOdometer - (startOdometer || dialogRequest.StartOdometer));
    this.setState({ isActioning: true });
    try {
      await this.svc.completeTrip(dialogRequest.ID, dialogRequest, {
        actualStartTime:  dialogRequest.ActualStartTime || nowISO(),
        actualEndTime:    nowISO(),
        startOdometer:    startOdometer || dialogRequest.StartOdometer,
        endOdometer,
        totalDistance:    total,
        adminNote,
      }, this._user());
      this.setState({ activeDialog: null, dialogRequest: null });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    } finally {
      this.setState({ isActioning: false });
    }
  }

  public render(): React.ReactElement {
    const { requests, isLoading, error, activeDialog, isActioning, startOdometer, endOdometer, adminNote } = this.state;

    return (
      <div>
        <PageHeader icon="Road" title="Chuyến của tôi"
          subtitle={`${requests.length} chuyến đang hoạt động`} />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : requests.length === 0 ? (
          <div className={styles.empty}>
            <FontIcon iconName="Road" style={{ fontSize: 40, color: '#C0C5CC', display: 'block', marginBottom: 12 }} />
            <Text>Bạn không có chuyến nào đang hoạt động</Text>
          </div>
        ) : (
          <div className={styles.grid}>
            {requests.map(r => (
              <div key={r.ID} className={styles.card}>
                <div className={styles.cardHead}>
                  <span className={styles.cardCode}>{r.RequestCode}</span>
                  <StatusBadge status={r.Status} />
                </div>
                <div className={styles.cardBody}>
                  <div className={styles.row}>
                    <span className={styles.label}>Điểm đón:</span>
                    <span>{r.PickupLocation}</span>
                  </div>
                  <div className={styles.row}>
                    <span className={styles.label}>Điểm đến:</span>
                    <span>{r.DropoffLocation}</span>
                  </div>
                  <div className={styles.row}>
                    <span className={styles.label}>Thời gian:</span>
                    <span style={{ color: TEXT_MUTE }}>{formatDateTime(r.PickupDateTime)}</span>
                  </div>
                  <div className={styles.row}>
                    <span className={styles.label}>Hành khách:</span>
                    <span>{r.NumberOfPassengers} người</span>
                  </div>
                  <div className={styles.row}>
                    <span className={styles.label}>Mục đích:</span>
                    <span>{r.Purpose}</span>
                  </div>
                  {r.StartOdometer > 0 && (
                    <div className={styles.row}>
                      <span className={styles.label}>ODO bắt đầu:</span>
                      <span>{r.StartOdometer} km</span>
                    </div>
                  )}
                </div>
                <div className={styles.actions}>
                  {(r.Status === STATUS.VEHICLE_ASSIGNED || r.Status === STATUS.CONFIRMED) && (
                    <PrimaryButton text="Xác nhận nhận chuyến" disabled={isActioning}
                      onClick={() => this._acknowledge(r).catch(console.error)}
                      styles={{ root: { borderRadius: 8, flex: 1 } }} />
                  )}
                  {r.Status === STATUS.DRIVER_CONFIRMED && (
                    <PrimaryButton text="Bắt đầu chuyến" disabled={isActioning}
                      onClick={() => this.setState({ activeDialog: 'start', dialogRequest: r, startOdometer: 0 })}
                      styles={{ root: { borderRadius: 8, flex: 1 } }} />
                  )}
                  {r.Status === STATUS.IN_PROGRESS && (
                    <PrimaryButton text="Hoàn thành chuyến" disabled={isActioning}
                      onClick={() => this.setState({ activeDialog: 'complete', dialogRequest: r, endOdometer: 0 })}
                      styles={{ root: { borderRadius: 8, flex: 1, background: '#20BF6B', border: 'none' } }} />
                  )}
                  <DefaultButton text="Xem chi tiết"
                    onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}
                    styles={{ root: { borderRadius: 8 } }} />
                </div>
              </div>
            ))}
          </div>
        )}

        <Dialog hidden={activeDialog !== 'start'}
          dialogContentProps={{ type: DialogType.normal, title: 'Bắt đầu chuyến đi' }}
          onDismiss={() => this.setState({ activeDialog: null })}>
          <TextField label="ODO đồng hồ km hiện tại *" type="number"
            value={String(startOdometer)}
            onChange={(_, v) => this.setState({ startOdometer: parseInt(v || '0', 10) })} />
          <DialogFooter>
            <PrimaryButton text={isActioning ? 'Đang xử lý...' : 'Bắt đầu'}
              disabled={isActioning || !startOdometer}
              onClick={() => this._startTrip().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ activeDialog: null })} />
          </DialogFooter>
        </Dialog>

        <Dialog hidden={activeDialog !== 'complete'}
          dialogContentProps={{ type: DialogType.normal, title: 'Hoàn thành chuyến đi' }}
          onDismiss={() => this.setState({ activeDialog: null })}>
          <TextField label="ODO đồng hồ km kết thúc *" type="number"
            value={String(endOdometer)}
            onChange={(_, v) => this.setState({ endOdometer: parseInt(v || '0', 10) })} />
          <TextField label="Ghi chú (tùy chọn)" multiline rows={2}
            value={adminNote} onChange={(_, v) => this.setState({ adminNote: v || '' })} />
          <DialogFooter>
            <PrimaryButton text={isActioning ? 'Đang xử lý...' : 'Xác nhận hoàn thành'}
              disabled={isActioning || !endOdometer}
              onClick={() => this._completeTrip().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ activeDialog: null })} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
