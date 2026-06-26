import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, IconButton, mergeStyleSets, FontIcon, Text, SearchBox,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import AssignVehiclePanel from './AssignVehiclePanel';

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
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 640 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 14px', fontSize: 13, borderBottom: `1px solid ${BORDER}` },
  empty: { padding: '48px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  mobileCard: {
    padding: '14px', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
  },
  mobileCode: { fontWeight: 700, color: PRIMARY, fontSize: 13, marginBottom: 4, cursor: 'pointer' },
  mobileLine: { fontSize: 12, color: '#1B1F23', marginBottom: 3 },
  mobileMeta: { fontSize: 11, color: TEXT_MUTE, marginBottom: 10 },
  mobileBottom: { display: 'flex' as const, justifyContent: 'space-between' as const, alignItems: 'center' as const },
});

interface ITransportAssignmentProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface ITransportAssignmentState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  searchText: string;
  panelRequest: IVehicleBookingRequest | null;
  isPanelOpen: boolean;
  isMobile: boolean;
}

export default class TransportAssignment extends React.Component<ITransportAssignmentProps, ITransportAssignmentState> {
  private svc: VehicleBookingRequestService;
  private _mql: MediaQueryList | null = null;

  constructor(props: ITransportAssignmentProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null, searchText: '',
      panelRequest: null, isPanelOpen: false,
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 768,
    };
    this.svc = new VehicleBookingRequestService(props.context);
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
    if (typeof window !== 'undefined') {
      this._mql = window.matchMedia('(max-width: 768px)');
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
      const requests = await this.svc.getPendingTransportAssignment();
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
      r.RequestCode.toLowerCase().includes(q) ||
      r.RequesterName.toLowerCase().includes(q) ||
      r.PickupLocation.toLowerCase().includes(q)
    );
  }

  public render(): React.ReactElement {
    const { isLoading, error, searchText, panelRequest, isPanelOpen, isMobile } = this.state;
    const filtered = this._filtered();

    return (
      <div>
        <PageHeader icon="Car" title="Phân công xe"
          subtitle="Danh sách yêu cầu cần phân công xe & tài xế"
          actions={
            <IconButton iconProps={{ iconName: 'Refresh' }} title="Làm mới"
              onClick={() => this._load().catch(console.error)}
              styles={{ root: { borderRadius: 8, border: `1px solid ${BORDER}` } }} />
          } />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}

        <div className={styles.toolbar}>
          <div style={{ flex: 1 }}>
            <SearchBox placeholder="Tìm theo mã, người yêu cầu, điểm đón..."
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
                    <FontIcon iconName="Car" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                    <Text>Không có yêu cầu nào cần phân công</Text>
                  </div>
                ) : filtered.map(r => (
                  <div key={r.ID} className={styles.mobileCard}>
                    <div className={styles.mobileCode}
                      onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                      {r.RequestCode}
                    </div>
                    <div className={styles.mobileLine}>{r.RequesterName} · {r.Department}</div>
                    <div className={styles.mobileLine}>{r.PickupLocation} → {r.DropoffLocation}</div>
                    <div className={styles.mobileMeta}>{formatDateTime(r.PickupDateTime)} · {r.VehicleType}</div>
                    {r.AssignedVehiclePlate && (
                      <div style={{ fontSize: 12, marginBottom: 8, color: TEXT_MUTE }}>
                        Xe: {r.AssignedVehiclePlate} {r.AssignedDriverName && `· ${r.AssignedDriverName}`}
                      </div>
                    )}
                    <div className={styles.mobileBottom}>
                      <StatusBadge status={r.Status} />
                      <PrimaryButton text="Phân công" styles={{ root: { borderRadius: 6, height: 30 } }}
                        onClick={() => this.setState({ panelRequest: r, isPanelOpen: true })} />
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
                      <th className={styles.th}>Phòng ban</th>
                      <th className={styles.th}>Hành trình</th>
                      <th className={styles.th}>Thời gian đón</th>
                      <th className={styles.th}>Loại xe</th>
                      <th className={styles.th}>Xe đã gán</th>
                      <th className={styles.th}>Trạng thái</th>
                      <th className={styles.th}>Thao tác</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.length === 0 ? (
                      <tr>
                        <td colSpan={9}>
                          <div className={styles.empty}>
                            <FontIcon iconName="Car" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                            <Text>Không có yêu cầu nào cần phân công</Text>
                          </div>
                        </td>
                      </tr>
                    ) : filtered.map(r => (
                      <tr key={r.ID}>
                        <td className={styles.td}>
                          <span style={{ fontWeight: 600, color: PRIMARY, cursor: 'pointer' }}
                            onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                            {r.RequestCode}
                          </span>
                        </td>
                        <td className={styles.td}>{r.RequesterName}</td>
                        <td className={styles.td}>{r.Department}</td>
                        <td className={styles.td}>
                          {r.PickupLocation} <span style={{ color: TEXT_MUTE }}>→</span> {r.DropoffLocation}
                        </td>
                        <td className={styles.td} style={{ color: TEXT_MUTE }}>
                          {formatDateTime(r.PickupDateTime)}
                        </td>
                        <td className={styles.td}>{r.VehicleType}</td>
                        <td className={styles.td} style={{ color: r.AssignedVehiclePlate ? '#1B1F23' : TEXT_MUTE }}>
                          {r.AssignedVehiclePlate || '-'}
                          {r.AssignedDriverName && <div style={{ fontSize: 11, color: TEXT_MUTE }}>{r.AssignedDriverName}</div>}
                        </td>
                        <td className={styles.td}><StatusBadge status={r.Status} /></td>
                        <td className={styles.td}>
                          <PrimaryButton text="Phân công" styles={{ root: { borderRadius: 6, height: 28, minWidth: 80 } }}
                            onClick={() => this.setState({ panelRequest: r, isPanelOpen: true })} />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        <AssignVehiclePanel
          context={this.props.context}
          user={this.props.user}
          request={panelRequest}
          isOpen={isPanelOpen}
          onDismiss={() => this.setState({ isPanelOpen: false, panelRequest: null })}
          onAssigned={() => {
            this.setState({ isPanelOpen: false, panelRequest: null });
            this._load().catch(console.error);
          }}
        />
      </div>
    );
  }
}
