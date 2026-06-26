import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  mergeStyleSets, FontIcon, Text,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { STATUS, STATUS_LABEL } from '../../common/constants/statuses';
import { ROLE } from '../../common/constants/roles';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';

const PRIMARY  = '#0078D4';
const WHITE    = '#FFFFFF';
const BORDER   = '#E1E4E8';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  cardGrid: {
    display: 'grid' as const,
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    gap: 12,
    marginBottom: 20,
  },
  card: {
    background: WHITE,
    borderRadius: 12,
    padding: '14px 16px',
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 12,
  },
  cardIcon: {
    width: 40, height: 40, borderRadius: 10,
    display: 'flex' as const, alignItems: 'center' as const, justifyContent: 'center' as const,
    fontSize: 18, flexShrink: 0,
  },
  cardValue: { fontSize: 24, fontWeight: 700, lineHeight: '1', color: '#1B1F23' },
  cardLabel: { fontSize: 11, color: TEXT_MUTE, marginTop: 3, fontWeight: 500 },
  tableWrap: {
    background: WHITE, borderRadius: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
    overflow: 'hidden',
  },
  tableTitle: {
    padding: '12px 16px', borderBottom: `1px solid ${BORDER}`,
    fontWeight: 700, fontSize: 14, color: '#1B1F23',
  },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 520 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const,
    fontSize: 11, fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: {
    padding: '10px 14px', fontSize: 13, color: '#1B1F23',
    borderBottom: `1px solid ${BORDER}`,
  },
  clickRow: { cursor: 'pointer', ':hover': { background: '#F6F8FA' } },
  empty: { padding: '32px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  mobileCard: {
    padding: '12px 14px',
    borderBottom: `1px solid ${BORDER}`,
    cursor: 'pointer',
    ':hover': { background: '#F6F8FA' },
    ':last-child': { borderBottom: 'none' },
  },
  mobileCode: { fontWeight: 700, color: PRIMARY, fontSize: 13, marginBottom: 4 },
  mobileRoute: { fontSize: 12, color: '#1B1F23', marginBottom: 4 },
  mobileMeta: { fontSize: 11, color: TEXT_MUTE, display: 'flex' as const, justifyContent: 'space-between' as const, alignItems: 'center' as const },
});

const STAT_CARDS = [
  { key: 'total',     label: 'Tổng cộng',      icon: 'BulletedList', bg: '#EBF4FF', color: PRIMARY },
  { key: 'pending',   label: 'Chờ xử lý',      icon: 'Clock',        bg: '#FFF8E6', color: '#F59E0B' },
  { key: 'inprogress',label: 'Đang thực hiện', icon: 'Running',      bg: '#F0FDF4', color: '#22C55E' },
  { key: 'completed', label: 'Hoàn thành',     icon: 'CheckMark',    bg: '#E8FFF2', color: '#20BF6B' },
  { key: 'rejected',  label: 'Từ chối/Hủy',   icon: 'Cancel',       bg: '#FFF0F1', color: '#FC5C65' },
];

interface IDashboardProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IDashboardState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  isMobile: boolean;
}

export default class Dashboard extends React.Component<IDashboardProps, IDashboardState> {
  private svc: VehicleBookingRequestService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IDashboardProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null,
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
      const { user, context } = this.props;
      let requests: IVehicleBookingRequest[];
      if (user.role === ROLE.REQUESTER) {
        requests = await this.svc.getMyRequests(context.pageContext.user.email);
      } else {
        requests = await this.svc.getAllActive();
      }
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _counts(): Record<string, number> {
    const { requests } = this.state;
    const pending: string[] = [
      STATUS.SUBMITTED, STATUS.PENDING_MANAGER_APPROVAL, STATUS.RESUBMITTED,
      STATUS.PENDING_TRANSPORT_ASSIGNMENT, STATUS.VEHICLE_ASSIGNED, STATUS.CONFIRMED,
    ];
    const rejected: string[] = [STATUS.REJECTED, STATUS.REJECTED_NO_VEHICLE, STATUS.CANCELLED];
    return {
      total:      requests.length,
      pending:    requests.filter(r => pending.indexOf(r.Status) !== -1).length,
      inprogress: requests.filter(r => r.Status === STATUS.IN_PROGRESS || r.Status === STATUS.DRIVER_CONFIRMED).length,
      completed:  requests.filter(r => r.Status === STATUS.COMPLETED).length,
      rejected:   requests.filter(r => rejected.indexOf(r.Status) !== -1).length,
    };
  }

  public render(): React.ReactElement {
    const { isLoading, error, requests, isMobile } = this.state;
    const counts = this._counts();
    const recent = requests.slice(0, 10);

    return (
      <div>
        <PageHeader
          icon="ViewDashboard"
          title="Tổng quan"
          subtitle="Hệ thống quản lý đặt xe công ty"
        />

        {error && (
          <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 8, marginBottom: 16 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <>
            <div className={styles.cardGrid}>
              {STAT_CARDS.map(c => (
                <div key={c.key} className={styles.card}>
                  <div className={styles.cardIcon} style={{ background: c.bg }}>
                    <FontIcon iconName={c.icon} style={{ color: c.color }} />
                  </div>
                  <div>
                    <div className={styles.cardValue}>{counts[c.key] || 0}</div>
                    <div className={styles.cardLabel}>{c.label}</div>
                  </div>
                </div>
              ))}
            </div>

            <div className={styles.tableWrap}>
              <div className={styles.tableTitle}>Yêu cầu gần đây</div>
              {isMobile ? (
                <div>
                  {recent.length === 0 ? (
                    <div className={styles.empty}><Text>Chưa có yêu cầu nào</Text></div>
                  ) : recent.map(r => (
                    <div key={r.ID} className={styles.mobileCard}
                      onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                      <div className={styles.mobileCode}>{r.RequestCode}</div>
                      <div className={styles.mobileRoute}>{r.PickupLocation} → {r.DropoffLocation}</div>
                      <div className={styles.mobileMeta}>
                        <span>{formatDateTime(r.PickupDateTime)}</span>
                        <StatusBadge status={r.Status} />
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
                        <th className={styles.th}>Điểm đón</th>
                        <th className={styles.th}>Điểm đến</th>
                        <th className={styles.th}>Thời gian đón</th>
                        <th className={styles.th}>Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {recent.length === 0 ? (
                        <tr>
                          <td colSpan={6} className={styles.empty}>
                            <Text>Chưa có yêu cầu nào</Text>
                          </td>
                        </tr>
                      ) : recent.map(r => (
                        <tr key={r.ID} className={styles.clickRow}
                          onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                          <td className={styles.td} style={{ fontWeight: 600, color: PRIMARY }}>{r.RequestCode}</td>
                          <td className={styles.td}>{r.RequesterName}</td>
                          <td className={styles.td}>{r.PickupLocation}</td>
                          <td className={styles.td}>{r.DropoffLocation}</td>
                          <td className={styles.td} style={{ color: TEXT_MUTE }}>{formatDateTime(r.PickupDateTime)}</td>
                          <td className={styles.td}><StatusBadge status={r.Status} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </>
        )}
      </div>
    );
  }
}
