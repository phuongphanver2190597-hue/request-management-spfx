import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  mergeStyleSets, FontIcon, Text, Dropdown, IDropdownOption,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { VehicleMasterService } from '../../services/vehicleMasterService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { IVehicleMaster } from '../../models/VehicleMaster';
import { STATUS } from '../../common/constants/statuses';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const PRIMARY = '#0078D4';
const TEXT_MUTE = '#6A737D';

const ACTIVE_STATUSES: string[] = [
  STATUS.VEHICLE_ASSIGNED, STATUS.CONFIRMED, STATUS.DRIVER_CONFIRMED, STATUS.IN_PROGRESS,
];

const styles = mergeStyleSets({
  toolbar: {
    background: WHITE, borderRadius: 10, padding: '10px 14px',
    marginBottom: 16, boxShadow: '0 1px 4px rgba(0,0,0,0.06)',
    display: 'flex' as const, alignItems: 'center' as const, gap: 12, flexWrap: 'wrap' as const,
  },
  grid: {
    display: 'grid' as const,
    gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))',
    gap: 14,
  },
  vehicleCard: {
    background: WHITE, borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden',
  },
  vehicleHead: {
    padding: '12px 16px', background: PRIMARY,
    display: 'flex' as const, justifyContent: 'space-between' as const, alignItems: 'center' as const,
  },
  vehicleName: { color: WHITE, fontWeight: 700, fontSize: 14 },
  vehicleType: { color: 'rgba(255,255,255,0.75)', fontSize: 12 },
  tripItem: {
    padding: '10px 14px', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
    cursor: 'pointer',
    ':hover': { background: '#F6F8FA' },
  },
  tripCode: { fontWeight: 600, fontSize: 12, color: PRIMARY },
  tripRoute: { fontSize: 13, color: '#1B1F23', margin: '2px 0' },
  tripTime: { fontSize: 11, color: TEXT_MUTE },
  noTrip: { padding: '20px 14px', color: TEXT_MUTE, fontSize: 13, textAlign: 'center' as const },
});

interface IVehicleScheduleCalendarProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IVehicleScheduleState {
  requests: IVehicleBookingRequest[];
  vehicles: IVehicleMaster[];
  isLoading: boolean;
  error: string | null;
  filterVehicle: string;
}

export default class VehicleScheduleCalendar extends React.Component<IVehicleScheduleCalendarProps, IVehicleScheduleState> {
  private bookingSvc: VehicleBookingRequestService;
  private vehicleSvc: VehicleMasterService;

  constructor(props: IVehicleScheduleCalendarProps) {
    super(props);
    this.state = { requests: [], vehicles: [], isLoading: false, error: null, filterVehicle: 'all' };
    this.bookingSvc = new VehicleBookingRequestService(props.context);
    this.vehicleSvc = new VehicleMasterService(props.context);
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
  }

  private async _load(): Promise<void> {
    this.setState({ isLoading: true, error: null });
    try {
      const [requests, vehicles] = await Promise.all([
        this.bookingSvc.getAllActive(),
        this.vehicleSvc.getAll(),
      ]);
      const active = requests.filter(r => ACTIVE_STATUSES.indexOf(r.Status) !== -1 && r.AssignedVehicleId > 0);
      this.setState({ requests: active, vehicles, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  public render(): React.ReactElement {
    const { requests, vehicles, isLoading, error, filterVehicle } = this.state;

    const usedVehicleIds = Array.from(new Set(requests.map(r => r.AssignedVehicleId)));
    const displayVehicles = vehicles.filter(v =>
      usedVehicleIds.indexOf(v.ID) !== -1 && (filterVehicle === 'all' || String(v.ID) === filterVehicle)
    );

    const vehicleFilterOptions: IDropdownOption[] = [
      { key: 'all', text: 'Tất cả xe' },
      ...displayVehicles.map(v => ({ key: String(v.ID), text: `${v.PlateNumber} - ${v.Brand} ${v.Model}` })),
    ];

    return (
      <div>
        <PageHeader icon="Calendar" title="Lịch xe"
          subtitle="Xem lịch sử dụng xe theo từng phương tiện" />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}>
            {error}
          </MessageBar>
        )}

        <div className={styles.toolbar}>
          <div style={{ minWidth: 200, flex: 1 }}>
            <Dropdown options={vehicleFilterOptions} selectedKey={filterVehicle}
              onChange={(_, o) => this.setState({ filterVehicle: String(o?.key || 'all') })} />
          </div>
          <span style={{ fontSize: 12, color: TEXT_MUTE, flexShrink: 0 }}>
            {requests.length} chuyến · {usedVehicleIds.length} xe
          </span>
        </div>

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : displayVehicles.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '60px 0', color: TEXT_MUTE }}>
            <FontIcon iconName="Calendar" style={{ fontSize: 40, display: 'block', marginBottom: 12, color: '#C0C5CC' }} />
            <Text>Không có xe nào đang được sử dụng</Text>
          </div>
        ) : (
          <div className={styles.grid}>
            {displayVehicles.map(v => {
              const trips = requests.filter(r => r.AssignedVehicleId === v.ID);
              return (
                <div key={v.ID} className={styles.vehicleCard}>
                  <div className={styles.vehicleHead}>
                    <div>
                      <div className={styles.vehicleName}>{v.PlateNumber}</div>
                      <div className={styles.vehicleType}>{v.Brand} {v.Model} · {v.VehicleType}</div>
                    </div>
                    <StatusBadge status={v.Status} />
                  </div>
                  {trips.length === 0 ? (
                    <div className={styles.noTrip}>Không có chuyến nào</div>
                  ) : trips.map(r => (
                    <div key={r.ID} className={styles.tripItem}
                      onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                      <div className={styles.tripCode}>{r.RequestCode}</div>
                      <div className={styles.tripRoute}>{r.PickupLocation} → {r.DropoffLocation}</div>
                      <div className={styles.tripTime}>{formatDateTime(r.PickupDateTime)}</div>
                      <div style={{ marginTop: 4 }}><StatusBadge status={r.Status} /></div>
                    </div>
                  ))}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }
}
