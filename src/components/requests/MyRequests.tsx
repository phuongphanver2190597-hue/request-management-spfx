import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  SearchBox, Dropdown, IDropdownOption, IconButton,
  PrimaryButton, DefaultButton, mergeStyleSets, Text, FontIcon,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { STATUS_LABEL } from '../../common/constants/statuses';
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
    display: 'flex' as const, alignItems: 'center' as const, gap: 10, flexWrap: 'wrap' as const,
  },
  tableWrap: {
    background: WHITE, borderRadius: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden',
  },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 560 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 14px', fontSize: 13, borderBottom: `1px solid ${BORDER}` },
  trHover: { cursor: 'pointer', ':hover': { background: '#F6F8FA' } },
  empty: { padding: '48px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  mobileCard: {
    padding: '12px 14px', borderBottom: `1px solid ${BORDER}`,
    cursor: 'pointer', ':hover': { background: '#F6F8FA' },
    ':last-child': { borderBottom: 'none' },
  },
  mobileCode: { fontWeight: 700, color: PRIMARY, fontSize: 13, marginBottom: 4 },
  mobileRoute: { fontSize: 12, color: '#1B1F23', marginBottom: 4 },
  mobileMeta: {
    display: 'flex' as const, justifyContent: 'space-between' as const,
    alignItems: 'center' as const, fontSize: 11, color: TEXT_MUTE,
  },
  mobileType: { fontSize: 11, color: TEXT_MUTE, marginBottom: 4 },
});

const STATUS_FILTER_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'Tất cả trạng thái' },
  ...([
    'DRAFT','SUBMITTED','PENDING_MANAGER_APPROVAL','MANAGER_APPROVED',
    'NEED_MORE_INFORMATION','RESUBMITTED','PENDING_TRANSPORT_ASSIGNMENT',
    'VEHICLE_ASSIGNED','CONFIRMED','DRIVER_CONFIRMED','IN_PROGRESS',
    'COMPLETED','REJECTED','REJECTED_NO_VEHICLE','CANCELLED',
  ] as string[]).map(k => ({ key: k, text: STATUS_LABEL[k] || k })),
];

interface IMyRequestsProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IMyRequestsState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  searchText: string;
  statusFilter: string;
  isMobile: boolean;
}

export default class MyRequests extends React.Component<IMyRequestsProps, IMyRequestsState> {
  private svc: VehicleBookingRequestService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IMyRequestsProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null, searchText: '', statusFilter: 'all',
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
      const email = this.props.context.pageContext.user.email;
      const requests = await this.svc.getMyRequests(email);
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _filtered(): IVehicleBookingRequest[] {
    const { requests, searchText, statusFilter } = this.state;
    const q = searchText.toLowerCase().trim();
    return requests.filter(r => {
      const matchSearch = !q || r.RequestCode.toLowerCase().includes(q) ||
        r.PickupLocation.toLowerCase().includes(q) || r.DropoffLocation.toLowerCase().includes(q);
      const matchStatus = statusFilter === 'all' || r.Status === statusFilter;
      return matchSearch && matchStatus;
    });
  }

  public render(): React.ReactElement {
    const { isLoading, error, searchText, statusFilter, requests, isMobile } = this.state;
    const filtered = this._filtered();

    return (
      <div>
        <PageHeader icon="DocumentSet" title="Yêu cầu của tôi"
          subtitle={`${requests.length} yêu cầu`}
          actions={
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="Tạo yêu cầu mới"
              onClick={() => this.props.onNavigate('create-request')}
              styles={{ root: { borderRadius: 8 } }} />
          } />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}

        <div className={styles.toolbar}>
          <div style={{ flex: 3, minWidth: 160 }}>
            <SearchBox placeholder="Tìm theo mã, điểm đón, điểm đến..."
              value={searchText} onChange={(_, v) => this.setState({ searchText: v || '' })}
              styles={{ root: { borderRadius: 8 } }} />
          </div>
          {!isMobile && (
            <div style={{ flex: 1, minWidth: 160 }}>
              <Dropdown options={STATUS_FILTER_OPTIONS} selectedKey={statusFilter}
                onChange={(_, o) => this.setState({ statusFilter: String(o?.key ?? 'all') })}
                styles={{ dropdown: { borderRadius: 8 } }} />
            </div>
          )}
          <IconButton iconProps={{ iconName: 'Refresh' }} title="Làm mới"
            onClick={() => this._load().catch(console.error)}
            styles={{ root: { borderRadius: 8, border: `1px solid ${BORDER}` } }} />
        </div>

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <div className={styles.tableWrap}>
            {isMobile ? (
              <div>
                {filtered.length === 0 ? (
                  <div className={styles.empty}>
                    <FontIcon iconName="InboxCheck" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                    <Text>{requests.length === 0 ? 'Bạn chưa có yêu cầu nào' : 'Không tìm thấy kết quả'}</Text>
                  </div>
                ) : filtered.map(r => (
                  <div key={r.ID} className={styles.mobileCard}
                    onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                    <div className={styles.mobileCode}>{r.RequestCode}</div>
                    <div className={styles.mobileRoute}>{r.PickupLocation} → {r.DropoffLocation}</div>
                    <div className={styles.mobileType}>{r.VehicleType} · {formatDateTime(r.PickupDateTime)}</div>
                    <div className={styles.mobileMeta}>
                      <StatusBadge status={r.Status} />
                      <DefaultButton text="Xem" styles={{ root: { borderRadius: 6, minWidth: 50, height: 26, fontSize: 12 } }}
                        onClick={(e) => { e.stopPropagation(); this.props.onNavigate('request-detail', { id: r.ID }); }} />
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
                      <th className={styles.th}>Điểm đón</th>
                      <th className={styles.th}>Điểm đến</th>
                      <th className={styles.th}>Thời gian đón</th>
                      <th className={styles.th}>Loại xe</th>
                      <th className={styles.th}>Trạng thái</th>
                      <th className={styles.th}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.length === 0 ? (
                      <tr>
                        <td colSpan={7}>
                          <div className={styles.empty}>
                            <FontIcon iconName="InboxCheck" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                            <Text>{requests.length === 0 ? 'Bạn chưa có yêu cầu nào' : 'Không tìm thấy kết quả'}</Text>
                          </div>
                        </td>
                      </tr>
                    ) : filtered.map(r => (
                      <tr key={r.ID} className={styles.trHover}
                        onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                        <td className={styles.td} style={{ fontWeight: 600, color: PRIMARY }}>{r.RequestCode}</td>
                        <td className={styles.td}>{r.PickupLocation}</td>
                        <td className={styles.td}>{r.DropoffLocation}</td>
                        <td className={styles.td} style={{ color: TEXT_MUTE }}>{formatDateTime(r.PickupDateTime)}</td>
                        <td className={styles.td}>{r.VehicleType}</td>
                        <td className={styles.td}><StatusBadge status={r.Status} /></td>
                        <td className={styles.td}>
                          <DefaultButton text="Xem" styles={{ root: { borderRadius: 6, minWidth: 50, height: 28 } }}
                            onClick={(e) => { e.stopPropagation(); this.props.onNavigate('request-detail', { id: r.ID }); }} />
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
