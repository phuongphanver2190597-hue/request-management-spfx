import * as React from 'react';
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, TextField, mergeStyleSets,
  Dialog, DialogType, DialogFooter, Dropdown, IDropdownOption,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { VehicleBookingHistoryService } from '../../services/vehicleBookingHistoryService';
import { VehicleBookingCommentService } from '../../services/vehicleBookingCommentService';
import { UserRoleService } from '../../services/userRoleService';
import { IUserRole } from '../../models/UserRole';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { IVehicleBookingHistory } from '../../models/VehicleBookingHistory';
import { IVehicleBookingComment } from '../../models/VehicleBookingComment';
import { STATUS } from '../../common/constants/statuses';
import { ROLE } from '../../common/constants/roles';
import { StatusBadge } from '../shared/StatusBadge';
import { PageHeader } from '../shared/PageHeader';
import { formatDateTime, nowISO } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import { buildApprovalLink } from '../../common/helpers/linkHelper';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const TEXT_MUTE = '#6A737D';
const PRIMARY = '#0078D4';

const styles = mergeStyleSets({
  layout: { display: 'grid' as const, gridTemplateColumns: '1fr 340px', gap: 16, alignItems: 'start' },
  layoutMobile: { display: 'grid' as const, gridTemplateColumns: '1fr', gap: 16 },
  card: {
    background: WHITE, borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden',
  },
  cardHead: {
    padding: '12px 16px', borderBottom: `1px solid ${BORDER}`,
    fontWeight: 700, fontSize: 14,
  },
  cardBody: { padding: '14px 16px' },
  detailRow: {
    display: 'grid' as const, gridTemplateColumns: '120px 1fr',
    gap: '6px 10px', padding: '7px 0',
    borderBottom: `1px solid ${BORDER}`,
    alignItems: 'start' as const,
    ':last-child': { borderBottom: 'none' },
  },
  detailKey: { fontSize: 11, fontWeight: 600, color: TEXT_MUTE, textTransform: 'uppercase' as const, paddingTop: 1 },
  detailVal: { fontSize: 13, color: '#1B1F23', wordBreak: 'break-word' as const },
  actionBar: { display: 'flex' as const, gap: 8, flexWrap: 'wrap' as const, padding: '12px 16px', borderTop: `1px solid ${BORDER}` },
  historyItem: {
    padding: '10px 0', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
  },
  historyAction: { fontSize: 12, fontWeight: 600, color: PRIMARY },
  historyMeta: { fontSize: 11, color: TEXT_MUTE, marginTop: 2 },
  commentItem: {
    padding: '10px 12px', background: '#F6F8FA', borderRadius: 8, marginBottom: 8,
  },
  commentText: { fontSize: 13, color: '#1B1F23' },
  commentMeta: { fontSize: 11, color: TEXT_MUTE, marginTop: 4 },
  commentInput: { marginTop: 12 },
});

interface IRequestDetailProps {
  context: WebPartContext;
  user: IAppUser;
  requestId: number;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IRequestDetailState {
  request: IVehicleBookingRequest | null;
  history: IVehicleBookingHistory[];
  comments: IVehicleBookingComment[];
  isLoading: boolean;
  error: string | null;
  actionNote: string;
  commentText: string;
  isActioning: boolean;
  confirmAction: string | null;
  isMobile: boolean;
  isSubmitDialogOpen: boolean;
  submitManagerEmail: string;
  managers: IUserRole[];
  linkCopied: boolean;
}

export default class RequestDetail extends React.Component<IRequestDetailProps, IRequestDetailState> {
  private bookingSvc: VehicleBookingRequestService;
  private historySvc: VehicleBookingHistoryService;
  private commentSvc: VehicleBookingCommentService;
  private userSvc: UserRoleService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IRequestDetailProps) {
    super(props);
    this.state = {
      request: null, history: [], comments: [],
      isLoading: false, error: null,
      actionNote: '', commentText: '', isActioning: false, confirmAction: null,
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 768,
      isSubmitDialogOpen: false, submitManagerEmail: '', managers: [], linkCopied: false,
    };
    this.bookingSvc = new VehicleBookingRequestService(props.context);
    this.historySvc = new VehicleBookingHistoryService(props.context);
    this.commentSvc = new VehicleBookingCommentService(props.context);
    this.userSvc = new UserRoleService(props.context);
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
      const [request, history, comments] = await Promise.all([
        this.bookingSvc.getById(this.props.requestId),
        this.historySvc.getByRequestId(this.props.requestId),
        this.commentSvc.getByRequestId(this.props.requestId),
      ]);
      this.setState({ request, history, comments, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _user() {
    const { user } = this.props;
    return { id: user.userId, name: user.userName, email: user.userEmail };
  }

  private async _doAction(action: string): Promise<void> {
    const { request, actionNote } = this.state;
    if (!request) return;
    this.setState({ isActioning: true, confirmAction: null, error: null });
    try {
      const u = this._user();
      if (action === 'approve') await this.bookingSvc.approveByManager(request.ID, request, u, actionNote);
      else if (action === 'reject') await this.bookingSvc.rejectByManager(request.ID, request, u, actionNote);
      else if (action === 'moreinfo') await this.bookingSvc.requestMoreInfo(request.ID, request, u, actionNote);
      else if (action === 'resubmit') await this.bookingSvc.resubmit(request.ID, request, u);
      else if (action === 'cancel') await this.bookingSvc.cancelRequest(request.ID, request, u, actionNote);
      this.setState({ isActioning: false, actionNote: '' });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isActioning: false });
    }
  }

  private async _addComment(): Promise<void> {
    const { commentText, request } = this.state;
    if (!commentText.trim() || !request) return;
    this.setState({ isActioning: true });
    try {
      await this.commentSvc.addComment(request.ID, request.RequestCode, commentText.trim(), this._user());
      const comments = await this.commentSvc.getByRequestId(request.ID);
      this.setState({ comments, commentText: '', isActioning: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isActioning: false });
    }
  }

  private async _openSubmitDialog(request: IVehicleBookingRequest): Promise<void> {
    const managers = await this.userSvc.getManagersForDepartment(request.Department);
    const all = managers.length > 0 ? managers : await this.userSvc.getAllManagers();
    this.setState({ isSubmitDialogOpen: true, managers: all, submitManagerEmail: all[0]?.UserEmail || '' });
  }

  private async _doSubmit(): Promise<void> {
    const { request, submitManagerEmail } = this.state;
    if (!request || !submitManagerEmail) return;
    this.setState({ isActioning: true, isSubmitDialogOpen: false, error: null });
    try {
      const approvalLink = buildApprovalLink(request.ID);
      await this.bookingSvc.submitRequest(request.ID, request, this._user(), submitManagerEmail, approvalLink);
      this.setState({ isActioning: false });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isActioning: false });
    }
  }

  private _renderActions(request: IVehicleBookingRequest): React.ReactNode {
    const { user } = this.props;
    const { isActioning } = this.state;
    const actions: React.ReactNode[] = [];

    const isRequester = user.userEmail === request.RequesterEmail;

    const canEdit   = isRequester && request.Status === STATUS.DRAFT;
    const canSubmit = isRequester && request.Status === STATUS.DRAFT;

    const canApprove = user.role === ROLE.MANAGER && user.userEmail === request.CurrentApproverId &&
      (request.Status === STATUS.PENDING_MANAGER_APPROVAL || request.Status === STATUS.RESUBMITTED);

    const canCancel = user.userEmail === request.RequesterEmail &&
      ([STATUS.DRAFT, STATUS.SUBMITTED, STATUS.PENDING_MANAGER_APPROVAL, STATUS.NEED_MORE_INFORMATION, STATUS.RESUBMITTED] as string[]).indexOf(request.Status) !== -1;

    const canResubmit = user.userEmail === request.RequesterEmail && request.Status === STATUS.NEED_MORE_INFORMATION;

    if (canSubmit) {
      actions.push(
        <PrimaryButton key="submit" text="Gửi yêu cầu" disabled={isActioning}
          iconProps={{ iconName: 'Send' }}
          onClick={() => this._openSubmitDialog(request).catch(console.error)}
          styles={{ root: { borderRadius: 8 } }} />
      );
    }

    if (canEdit) {
      actions.push(
        <DefaultButton key="edit" text="Chỉnh sửa" disabled={isActioning}
          iconProps={{ iconName: 'Edit' }}
          onClick={() => this.props.onNavigate('create-request', { editId: request.ID })}
          styles={{ root: { borderRadius: 8 } }} />
      );
    }

    if (canApprove) {
      actions.push(
        <PrimaryButton key="approve" text="Duyệt" disabled={isActioning}
          onClick={() => this.setState({ confirmAction: 'approve' })}
          styles={{ root: { borderRadius: 8, background: '#20BF6B', border: 'none' } }} />,
        <DefaultButton key="reject" text="Từ chối" disabled={isActioning}
          onClick={() => this.setState({ confirmAction: 'reject' })}
          styles={{ root: { borderRadius: 8, color: '#C53030' } }} />,
        <DefaultButton key="moreinfo" text="Yêu cầu bổ sung" disabled={isActioning}
          onClick={() => this.setState({ confirmAction: 'moreinfo' })}
          styles={{ root: { borderRadius: 8 } }} />,
      );
    }

    if (canResubmit) {
      actions.push(
        <PrimaryButton key="resubmit" text="Gửi lại" disabled={isActioning}
          onClick={() => this._doAction('resubmit').catch(console.error)}
          styles={{ root: { borderRadius: 8 } }} />
      );
    }

    if (canCancel) {
      actions.push(
        <DefaultButton key="cancel" text="Hủy yêu cầu" disabled={isActioning}
          onClick={() => this.setState({ confirmAction: 'cancel' })}
          styles={{ root: { borderRadius: 8, color: '#C53030' } }} />
      );
    }

    const { linkCopied } = this.state;
    actions.push(
      <DefaultButton key="copy-link"
        text={linkCopied ? 'Đã sao chép!' : 'Sao chép link'}
        iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Link' }}
        disabled={isActioning}
        onClick={() => this._copyLink(request)}
        styles={{ root: { borderRadius: 8, marginLeft: 'auto' } }} />
    );

    return <div className={styles.actionBar}>{actions}</div>;
  }

  private _copyLink(request: IVehicleBookingRequest): void {
    const link = buildApprovalLink(request.ID);
    navigator.clipboard.writeText(link).then(() => {
      this.setState({ linkCopied: true });
      setTimeout(() => this.setState({ linkCopied: false }), 2500);
    }).catch(() => {
      // fallback cho browser cũ
      const el = document.createElement('textarea');
      el.value = link;
      document.body.appendChild(el);
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
      this.setState({ linkCopied: true });
      setTimeout(() => this.setState({ linkCopied: false }), 2500);
    });
  }

  private _field(label: string, value: React.ReactNode): React.ReactNode {
    return (
      <div className={styles.detailRow}>
        <span className={styles.detailKey}>{label}</span>
        <span className={styles.detailVal}>{value}</span>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { request, history, comments, isLoading, error, confirmAction, actionNote, commentText, isActioning, isMobile } = this.state;

    if (isLoading) return <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '80px 0' } }} />;
    if (error) return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
    if (!request) return <MessageBar>Không tìm thấy yêu cầu</MessageBar>;

    const rightPanel = (
      <div>
        <div className={styles.card} style={{ marginBottom: 16 }}>
          <div className={styles.cardHead}>Lịch sử xử lý ({history.length})</div>
          <div className={styles.cardBody}>
            {history.length === 0 ? (
              <span style={{ color: TEXT_MUTE, fontSize: 13 }}>Chưa có lịch sử</span>
            ) : history.map(h => (
              <div key={h.ID} className={styles.historyItem}>
                <div className={styles.historyAction}>{h.Action}</div>
                <div style={{ fontSize: 12, color: '#1B1F23', margin: '2px 0' }}>
                  {h.OldStatus && <><StatusBadge status={h.OldStatus} /> → </>}
                  <StatusBadge status={h.NewStatus} />
                </div>
                {h.Note && <div style={{ fontSize: 12, marginTop: 2 }}>{h.Note}</div>}
                <div className={styles.historyMeta}>{h.ActionByName} · {formatDateTime(h.ActionDate)}</div>
              </div>
            ))}
          </div>
        </div>

        <div className={styles.card}>
          <div className={styles.cardHead}>Bình luận ({comments.length})</div>
          <div className={styles.cardBody}>
            {comments.map(c => (
              <div key={c.ID} className={styles.commentItem}>
                <div className={styles.commentText}>{c.Comment}</div>
                <div className={styles.commentMeta}>{c.CreatedByName} · {formatDateTime(c.CommentDate)}</div>
              </div>
            ))}
            <div className={styles.commentInput}>
              <TextField multiline rows={2} placeholder="Nhập bình luận..."
                value={commentText}
                onChange={(_, v) => this.setState({ commentText: v || '' })}
                disabled={isActioning} />
              <PrimaryButton text="Gửi" style={{ marginTop: 8, borderRadius: 8 }}
                disabled={!commentText.trim() || isActioning}
                onClick={() => this._addComment().catch(console.error)} />
            </div>
          </div>
        </div>
      </div>
    );

    return (
      <div>
        <PageHeader icon="DocumentSet" title={request.RequestCode} subtitle="Chi tiết yêu cầu đặt xe" />

        <div className={isMobile ? styles.layoutMobile : styles.layout}>
          {/* Left: detail + actions */}
          <div>
            <div className={styles.card} style={{ marginBottom: 16 }}>
              <div className={styles.cardHead}>Thông tin chuyến đi</div>
              <div className={styles.cardBody}>
                {this._field('Mã yêu cầu', request.RequestCode)}
                {this._field('Trạng thái', <StatusBadge status={request.Status} />)}
                {this._field('Người yêu cầu', request.RequesterName)}
                {this._field('Phòng ban', request.Department)}
                {this._field('Điện thoại', request.PhoneNumber)}
                {this._field('Điểm đón', request.PickupLocation)}
                {this._field('Điểm đến', request.DropoffLocation)}
                {this._field('Thời gian đón', formatDateTime(request.PickupDateTime))}
                {request.IsRoundTrip && this._field('Thời gian về', formatDateTime(request.ReturnDateTime))}
                {this._field('Khứ hồi', request.IsRoundTrip ? 'Có' : 'Không')}
                {this._field('Số hành khách', String(request.NumberOfPassengers))}
                {this._field('Loại xe yêu cầu', request.VehicleType)}
                {this._field('Mục đích', request.Purpose)}
                {request.SpecialRequirement && this._field('Yêu cầu đặc biệt', request.SpecialRequirement)}
              </div>
              {this._renderActions(request)}
            </div>

            {(request.AssignedVehiclePlate || request.AssignedDriverName) && (
              <div className={styles.card} style={{ marginBottom: 16 }}>
                <div className={styles.cardHead}>Thông tin phân công</div>
                <div className={styles.cardBody}>
                  {request.AssignedVehiclePlate && this._field('Biển số xe', request.AssignedVehiclePlate)}
                  {request.AssignedDriverName && this._field('Tài xế', request.AssignedDriverName)}
                  {request.AssignedDriverPhone && this._field('SĐT tài xế', request.AssignedDriverPhone)}
                  {request.ActualStartTime && this._field('Giờ xuất phát TT', formatDateTime(request.ActualStartTime))}
                  {request.ActualEndTime && this._field('Giờ kết thúc TT', formatDateTime(request.ActualEndTime))}
                  {request.TotalDistance > 0 && this._field('Tổng km', `${request.TotalDistance} km`)}
                </div>
              </div>
            )}

            {(request.AdminNote || request.CancelReason) && (
              <div className={styles.card} style={{ marginBottom: 16 }}>
                <div className={styles.cardHead}>Ghi chú</div>
                <div className={styles.cardBody}>
                  {request.AdminNote && this._field('Ghi chú admin', request.AdminNote)}
                  {request.CancelReason && this._field('Lý do hủy', request.CancelReason)}
                </div>
              </div>
            )}

            {isMobile && rightPanel}
          </div>

          {/* Right panel — desktop only */}
          {!isMobile && rightPanel}
        </div>

        <Dialog hidden={!this.state.isSubmitDialogOpen}
          dialogContentProps={{ type: DialogType.normal, title: 'Gửi yêu cầu phê duyệt' }}
          onDismiss={() => this.setState({ isSubmitDialogOpen: false })}>
          <Dropdown
            label="Người phê duyệt"
            required
            selectedKey={this.state.submitManagerEmail}
            options={this.state.managers.map((m): IDropdownOption => ({
              key: m.UserEmail, text: `${m.UserName} — ${m.Department || ''}`,
            }))}
            onChange={(_, o) => this.setState({ submitManagerEmail: String(o?.key || '') })}
            styles={{ dropdown: { borderRadius: 8 } }}
          />
          <DialogFooter>
            <PrimaryButton text="Gửi" disabled={!this.state.submitManagerEmail}
              onClick={() => this._doSubmit().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ isSubmitDialogOpen: false })} />
          </DialogFooter>
        </Dialog>

        <Dialog hidden={!confirmAction}
          dialogContentProps={{
            type: DialogType.normal,
            title: confirmAction === 'approve' ? 'Xác nhận duyệt' :
                   confirmAction === 'reject' ? 'Xác nhận từ chối' :
                   confirmAction === 'moreinfo' ? 'Yêu cầu bổ sung thông tin' : 'Xác nhận hủy',
          }}
          onDismiss={() => this.setState({ confirmAction: null })}>
          <TextField multiline rows={3} label="Ghi chú (tùy chọn)"
            value={actionNote}
            onChange={(_, v) => this.setState({ actionNote: v || '' })} />
          <DialogFooter>
            <PrimaryButton text="Xác nhận" disabled={isActioning}
              onClick={() => this._doAction(confirmAction!).catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ confirmAction: null })} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
