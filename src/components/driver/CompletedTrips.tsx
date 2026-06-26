import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  mergeStyleSets, FontIcon, Text, SearchBox,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const PRIMARY = '#0078D4';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  toolbar: {
    background: WHITE, borderRadius: 10, padding: '10px 14px',
    marginBottom: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.06)',
    display: 'flex' as const, alignItems: 'center' as const, gap: 10,
  },
  tableWrap: { background: WHITE, borderRadius: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden' },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 560 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 14px', fontSize: 13, borderBottom: `1px solid ${BORDER}` },
  empty: { padding: '48px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  mobileCard: {
    padding: '12px 14px', borderBottom: `1px solid ${BORDER}`,
    cursor: 'pointer', ':hover': { background: '#F6F8FA' },
    ':last-child': { borderBottom: 'none' },
  },
  mobileCode: { fontWeight: 700, color: PRIMARY, fontSize: 13, marginBottom: 4 },
  mobileLine: { fontSize: 12, color: '#1B1F23', marginBottom: 3 },
  mobileMeta: {
    display: 'flex' as const, justifyContent: 'space-between' as const,
    alignItems: 'center' as const, fontSize: 11, color: TEXT_MUTE, marginTop: 6,
  },
  mobileKm: { fontWeight: 700, color: PRIMARY },
});

interface ICompletedTripsProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface ICompletedTripsState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  searchText: string;
  isMobile: boolean;
}

export default class CompletedTrips extends React.Component<ICompletedTripsProps, ICompletedTripsState> {
  private svc: VehicleBookingRequestService;
  private _mql: MediaQueryList | null = null;

  constructor(props: ICompletedTripsProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null, searchText: '',
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 600,
    };
    this.svc = new VehicleBookingRequestService(props.context);
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
    if (typeof window !== 'undefined') {
      this._mql = window.matchMedia('(max-width: 600px)');
      this._mql.addListener(this._onMql);
    }
  }

  public componentWillUnmount(): void {
    if (this._mql) this._mql.removeListener(this._onMql);
  }

  private _onMql = (e: MediaQueryListEvent): void => { this.setState({ isMobile: e.matches }); }

  private async _load(): Promise<void> {
    this.setState({ isLoading: true, error: null });
    try {
      const requests = await this.svc.getCompletedTrips();
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _filtered(): IVehicleBookingRequest[] {
    const { requests, searchText } = this.state;
    if (!searchText.trim()) return requests;
    const q = searchText.toLowerCase();
    return requests.filter(r =>
      r.RequestCode.toLowerCase().includes(q) || r.RequesterName.toLowerCase().includes(q) ||
      r.AssignedVehiclePlate.toLowerCase().includes(q) || r.AssignedDriverName.toLowerCase().includes(q)
    );
  }

  public render(): React.ReactElement {
    const { isLoading, error, searchText, requests, isMobile } = this.state;
    const filtered = this._filtered();
    const totalKm = requests.reduce((s, r) => s + (r.TotalDistance || 0), 0);

    return (
      <div>
        <PageHeader icon="CheckMark" title="Chuyến hoàn thành"
          subtitle={`${requests.length} chuyến · Tổng ${totalKm.toLocaleString('vi-VN')} km`} />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}

        <div className={styles.toolbar}>
          <div style={{ flex: 1 }}>
            <SearchBox placeholder="Tìm theo mã, người yêu cầu, xe, tài xế..."
              value={searchText} onChange={(_, v) => this.setState({ searchText: v || '' })}
              styles={{ root: { borderRadius: 8 } }} />
          </div>
        </div>

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <div className={styles.tableWrap}>
            {isMobile ? (
              <div>
                {filtered.length === 0 ? (
                  <div className={styles.empty}>
                    <FontIcon iconName="CheckMark" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                    <Text>Chưa có chuyến nào hoàn thành</Text>
                  </div>
                ) : filtered.map(r => (
                  <div key={r.ID} className={styles.mobileCard}
                    onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                    <div className={styles.mobileCode}>{r.RequestCode}</div>
                    <div className={styles.mobileLine}>{r.RequesterName}</div>
                    <div className={styles.mobileLine}>{r.PickupLocation} → {r.DropoffLocation}</div>
                    <div className={styles.mobileLine} style={{ color: TEXT_MUTE }}>
                      {r.AssignedVehiclePlate} · {r.AssignedDriverName}
                    </div>
                    <div className={styles.mobileMeta}>
                      <span>{formatDateTime(r.CompletedDate)}</span>
                      <span className={styles.mobileKm}>
                        {r.TotalDistance > 0 ? `${r.TotalDistance} km` : '-'}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className={styles.tableScroll}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th className={styles.th}>Mã yêu cầu</th>
                      <th className={styles.th}>Người yêu cầu</th>
                      <th className={styles.th}>Hành trình</th>
                      <th className={styles.th}>Thời gian thực tế</th>
                      <th className={styles.th}>Xe / Tài xế</th>
                      <th className={styles.th}>Tổng km</th>
                      <th className={styles.th}>Hoàn thành</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.length === 0 ? (
                      <tr>
                        <td colSpan={7}>
                          <div className={styles.empty}>
                            <FontIcon iconName="CheckMark" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                            <Text>Chưa có chuyến nào hoàn thành</Text>
                          </div>
                        </td>
                      </tr>
                    ) : filtered.map(r => (
                      <tr key={r.ID} style={{ cursor: 'pointer' }}
                        onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                        <td className={styles.td} style={{ fontWeight: 600, color: PRIMARY }}>{r.RequestCode}</td>
                        <td className={styles.td}>{r.RequesterName}</td>
                        <td className={styles.td}>{r.PickupLocation} → {r.DropoffLocation}</td>
                        <td className={styles.td} style={{ color: TEXT_MUTE, fontSize: 12 }}>
                          <div>{formatDateTime(r.ActualStartTime)}</div>
                          <div>{formatDateTime(r.ActualEndTime)}</div>
                        </td>
                        <td className={styles.td}>
                          <div style={{ fontWeight: 600 }}>{r.AssignedVehiclePlate}</div>
                          <div style={{ color: TEXT_MUTE, fontSize: 12 }}>{r.AssignedDriverName}</div>
                        </td>
                        <td className={styles.td} style={{ fontWeight: 600 }}>
                          {r.TotalDistance > 0 ? `${r.TotalDistance} km` : '-'}
                        </td>
                        <td className={styles.td} style={{ color: TEXT_MUTE, fontSize: 12 }}>
                          {formatDateTime(r.CompletedDate)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }
}
