import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, TextField, mergeStyleSets,
  Dialog, DialogType, DialogFooter, IconButton, FontIcon, Text,
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
  tableWrap: { background: WHITE, borderRadius: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden' },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 600 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 14px', fontSize: 13, borderBottom: `1px solid ${BORDER}` },
  empty: { padding: '48px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  actionCell: { display: 'flex' as const, gap: 6, flexWrap: 'wrap' as const },
  mobileCard: {
    padding: '14px', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
  },
  mobileCode: { fontWeight: 700, color: PRIMARY, fontSize: 13, marginBottom: 4 },
  mobileLine: { fontSize: 12, color: '#1B1F23', marginBottom: 3 },
  mobileMeta: { fontSize: 11, color: TEXT_MUTE, marginBottom: 10 },
  mobileActions: { display: 'flex' as const, gap: 8, flexWrap: 'wrap' as const },
});

interface IPendingMyApprovalProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IPendingMyApprovalState {
  requests: IVehicleBookingRequest[];
  isLoading: boolean;
  error: string | null;
  selectedId: number | null;
  confirmAction: 'approve' | 'reject' | 'moreinfo' | null;
  note: string;
  isActioning: boolean;
  isMobile: boolean;
}

export default class PendingMyApproval extends React.Component<IPendingMyApprovalProps, IPendingMyApprovalState> {
  private svc: VehicleBookingRequestService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IPendingMyApprovalProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null,
      selectedId: null, confirmAction: null, note: '', isActioning: false,
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
      const requests = await this.svc.getPendingApproval(this.props.context.pageContext.user.email);
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _openAction(id: number, action: 'approve' | 'reject' | 'moreinfo'): void {
    this.setState({ selectedId: id, confirmAction: action, note: '' });
  }

  private async _confirm(): Promise<void> {
    const { selectedId, confirmAction, note, requests } = this.state;
    if (!selectedId || !confirmAction) return;
    const request = requests.find(r => r.ID === selectedId);
    if (!request) return;
    const { user } = this.props;
    const u = { id: user.userId, name: user.userName, email: user.userEmail };
    this.setState({ isActioning: true });
    try {
      if (confirmAction === 'approve') await this.svc.approveByManager(request.ID, request, u, note);
      else if (confirmAction === 'reject') await this.svc.rejectByManager(request.ID, request, u, note);
      else if (confirmAction === 'moreinfo') await this.svc.requestMoreInfo(request.ID, request, u, note);
      this.setState({ isActioning: false, confirmAction: null, selectedId: null, note: '' });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isActioning: false, confirmAction: null });
    }
  }

  public render(): React.ReactElement {
    const { requests, isLoading, error, confirmAction, note, isActioning, isMobile } = this.state;

    return (
      <div>
        <PageHeader icon="Clock" title="Chờ phê duyệt"
          subtitle={`${requests.length} yêu cầu đang chờ`}
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

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <div className={styles.tableWrap}>
            {isMobile ? (
              <div>
                {requests.length === 0 ? (
                  <div className={styles.empty}>
                    <FontIcon iconName="CheckMark" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                    <Text>Không có yêu cầu nào đang chờ phê duyệt</Text>
                  </div>
                ) : requests.map(r => (
                  <div key={r.ID} className={styles.mobileCard}>
                    <div className={styles.mobileCode}
                      onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}
                      style={{ cursor: 'pointer' }}>
                      {r.RequestCode}
                    </div>
                    <div className={styles.mobileLine}>{r.RequesterName} · {r.Department}</div>
                    <div className={styles.mobileLine}>{r.PickupLocation} → {r.DropoffLocation}</div>
                    <div className={styles.mobileMeta}>{formatDateTime(r.PickupDateTime)} · {r.VehicleType}</div>
                    <div style={{ marginBottom: 8 }}><StatusBadge status={r.Status} /></div>
                    <div className={styles.mobileActions}>
                      <PrimaryButton text="Duyệt"
                        styles={{ root: { borderRadius: 6, height: 30, background: '#20BF6B', border: 'none' } }}
                        onClick={() => this._openAction(r.ID, 'approve')} />
                      <DefaultButton text="Từ chối"
                        styles={{ root: { borderRadius: 6, height: 30, color: '#C53030' } }}
                        onClick={() => this._openAction(r.ID, 'reject')} />
                      <DefaultButton text="Bổ sung"
                        styles={{ root: { borderRadius: 6, height: 30 } }}
                        onClick={() => this._openAction(r.ID, 'moreinfo')} />
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
                      <th className={styles.th}>Điểm đón → Điểm đến</th>
                      <th className={styles.th}>Thời gian đón</th>
                      <th className={styles.th}>Loại xe</th>
                      <th className={styles.th}>Trạng thái</th>
                      <th className={styles.th}>Thao tác</th>
                    </tr>
                  </thead>
                  <tbody>
                    {requests.length === 0 ? (
                      <tr>
                        <td colSpan={8}>
                          <div className={styles.empty}>
                            <FontIcon iconName="CheckMark" style={{ fontSize: 36, color: '#C0C5CC', display: 'block', marginBottom: 10 }} />
                            <Text>Không có yêu cầu nào đang chờ phê duyệt</Text>
                          </div>
                        </td>
                      </tr>
                    ) : requests.map(r => (
                      <tr key={r.ID}>
                        <td className={styles.td}>
                          <span style={{ fontWeight: 600, color: PRIMARY, cursor: 'pointer' }}
                            onClick={() => this.props.onNavigate('request-detail', { id: r.ID })}>
                            {r.RequestCode}
                          </span>
                        </td>
                        <td className={styles.td}>{r.RequesterName}</td>
                        <td className={styles.td}>{r.Department}</td>
                        <td className={styles.td}>{r.PickupLocation} → {r.DropoffLocation}</td>
                        <td className={styles.td} style={{ color: TEXT_MUTE }}>{formatDateTime(r.PickupDateTime)}</td>
                        <td className={styles.td}>{r.VehicleType}</td>
                        <td className={styles.td}><StatusBadge status={r.Status} /></td>
                        <td className={styles.td}>
                          <div className={styles.actionCell}>
                            <PrimaryButton text="Duyệt" styles={{ root: { borderRadius: 6, height: 28, minWidth: 55, background: '#20BF6B', border: 'none' } }}
                              onClick={() => this._openAction(r.ID, 'approve')} />
                            <DefaultButton text="Từ chối" styles={{ root: { borderRadius: 6, height: 28, minWidth: 60, color: '#C53030' } }}
                              onClick={() => this._openAction(r.ID, 'reject')} />
                            <DefaultButton text="Bổ sung" styles={{ root: { borderRadius: 6, height: 28, minWidth: 60 } }}
                              onClick={() => this._openAction(r.ID, 'moreinfo')} />
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        <Dialog hidden={!confirmAction}
          dialogContentProps={{
            type: DialogType.normal,
            title: confirmAction === 'approve' ? 'Xác nhận duyệt yêu cầu' :
                   confirmAction === 'reject' ? 'Xác nhận từ chối' : 'Yêu cầu bổ sung thông tin',
            subText: confirmAction === 'approve' ? 'Yêu cầu sẽ được chuyển sang bộ phận vận tải.' :
                     confirmAction === 'reject' ? 'Yêu cầu sẽ bị từ chối và người dùng sẽ được thông báo.' :
                     'Người yêu cầu sẽ cần bổ sung thông tin trước khi tái gửi.',
          }}
          onDismiss={() => this.setState({ confirmAction: null })}>
          <TextField label="Ghi chú" multiline rows={3}
            placeholder="Nhập lý do hoặc ghi chú (tùy chọn)..."
            value={note} onChange={(_, v) => this.setState({ note: v || '' })} />
          <DialogFooter>
            <PrimaryButton text={isActioning ? 'Đang xử lý...' : 'Xác nhận'}
              disabled={isActioning} onClick={() => this._confirm().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ confirmAction: null })} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
