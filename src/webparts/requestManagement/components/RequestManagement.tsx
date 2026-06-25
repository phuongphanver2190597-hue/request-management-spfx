import * as React from 'react';
import {
  Stack,
  SearchBox,
  Dropdown,
  IDropdownOption,
  IconButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
  Modal,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  mergeStyleSets,
  DefaultButton,
  PrimaryButton,
  TextField,
  NormalPeoplePicker,
  IPersonaProps,
  Label,
  IStackTokens,
  FontIcon,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IRequestManagementProps } from './IRequestManagementProps';
import { IRequestManagementState, IRequest } from './IRequestManagementState';

const PRIMARY   = '#0078D4';
const PENDING   = '#F7B731';
const APPROVED  = '#20BF6B';
const REJECTED  = '#FC5C65';
const BG        = '#F0F2F5';
const WHITE     = '#FFFFFF';
const BORDER    = '#E1E4E8';
const TEXT_MAIN = '#1B1F23';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  root: {
    background: BG,
    minHeight: '100%',
    padding: '24px',
    boxSizing: 'border-box' as const,
    width: '100%',
    fontFamily: `'Segoe UI', -apple-system, sans-serif`,
  },

  /* ── Header ──────────────────────────────── */
  header: {
    display: 'flex' as const,
    alignItems: 'center' as const,
    justifyContent: 'space-between' as const,
    marginBottom: 24,
  },
  headerTitle: {
    fontSize: 22,
    fontWeight: 700,
    color: TEXT_MAIN,
    lineHeight: '1.2',
  },
  headerSub: {
    fontSize: 13,
    color: TEXT_MUTE,
    marginTop: 2,
  },

  /* ── Cards ───────────────────────────────── */
  cardGrid: {
    display: 'grid' as const,
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: 16,
    marginBottom: 20,
  },
  card: {
    background: WHITE,
    borderRadius: 12,
    padding: '18px 20px',
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 14,
    transition: 'box-shadow 0.15s',
    ':hover': { boxShadow: '0 4px 12px rgba(0,0,0,0.12)' },
  },
  cardIcon: {
    width: 44,
    height: 44,
    borderRadius: 10,
    display: 'flex' as const,
    alignItems: 'center' as const,
    justifyContent: 'center' as const,
    fontSize: 20,
    flexShrink: 0,
  },
  cardValue: {
    fontSize: 28,
    fontWeight: 700,
    lineHeight: '1',
    color: TEXT_MAIN,
  },
  cardLabel: {
    fontSize: 12,
    color: TEXT_MUTE,
    marginTop: 3,
    fontWeight: 500,
  },

  /* ── Toolbar ─────────────────────────────── */
  toolbar: {
    background: WHITE,
    borderRadius: 10,
    padding: '12px 16px',
    marginBottom: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.06)',
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 10,
    flexWrap: 'wrap' as const,
  },

  /* ── Table ───────────────────────────────── */
  tableWrapper: {
    background: WHITE,
    borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
    overflowX: 'auto' as const,
  },

  /* ── Status badge ────────────────────────── */
  badge: {
    display: 'inline-flex' as const,
    alignItems: 'center' as const,
    gap: 5,
    padding: '3px 10px',
    borderRadius: 20,
    fontSize: 12,
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
  },
  badgeDot: {
    width: 6,
    height: 6,
    borderRadius: '50%',
    display: 'inline-block' as const,
  },

  emptyState: {
    padding: '48px 0',
    textAlign: 'center' as const,
  },

  /* ── Modal ───────────────────────────────── */
  modalWrap: {
    width: 520,
    maxWidth: '95vw',
    maxHeight: '90vh',
    display: 'flex' as const,
    flexDirection: 'column' as const,
    borderRadius: 14,
    overflow: 'hidden' as const,
    background: WHITE,
  },
  modalHead: {
    padding: '20px 24px 16px',
    borderBottom: `1px solid ${BORDER}`,
    display: 'flex' as const,
    alignItems: 'center' as const,
    justifyContent: 'space-between' as const,
    flexShrink: 0,
  },
  modalHeadTitle: {
    fontSize: 17,
    fontWeight: 700,
    color: TEXT_MAIN,
  },
  modalBody: {
    padding: '20px 24px',
    overflowY: 'auto' as const,
    flex: 1,
  },
  modalFoot: {
    padding: '14px 24px',
    borderTop: `1px solid ${BORDER}`,
    flexShrink: 0,
    background: '#FAFBFC',
  },

  /* ── Detail fields ───────────────────────── */
  detailRow: {
    display: 'grid' as const,
    gridTemplateColumns: '130px 1fr',
    gap: '8px 16px',
    padding: '10px 0',
    borderBottom: `1px solid ${BORDER}`,
    alignItems: 'start' as const,
    ':last-child': { borderBottom: 'none' },
  },
  detailKey: {
    fontSize: 12,
    fontWeight: 600,
    color: TEXT_MUTE,
    textTransform: 'uppercase' as const,
    letterSpacing: '0.4px',
    paddingTop: 2,
  },
  detailVal: {
    fontSize: 14,
    color: TEXT_MAIN,
    wordBreak: 'break-word' as const,
  },
});

/* ── Card config ──────────────────────────── */
const CARDS = [
  { key: 'total',    label: 'Total',    icon: 'BulletedList',  bg: '#EBF4FF', color: PRIMARY  },
  { key: 'pending',  label: 'Pending',  icon: 'Clock',         bg: '#FFF8E6', color: PENDING  },
  { key: 'approved', label: 'Approved', icon: 'CheckMark',     bg: '#E8FFF2', color: APPROVED },
  { key: 'rejected', label: 'Rejected', icon: 'Cancel',        bg: '#FFF0F1', color: REJECTED },
];

const STATUS_STYLE: Record<string, { bg: string; color: string; dot: string }> = {
  Pending:  { bg: '#FFF8E6', color: '#B7791F', dot: PENDING  },
  Approved: { bg: '#E8FFF2', color: '#276749', dot: APPROVED },
  Rejected: { bg: '#FFF0F1', color: '#C53030', dot: REJECTED },
};

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'all',      text: 'All Status'  },
  { key: 'Pending',  text: 'Pending'     },
  { key: 'Approved', text: 'Approved'    },
  { key: 'Rejected', text: 'Rejected'    },
];

const GAP: IStackTokens = { childrenGap: 10 };

export default class RequestManagement extends React.Component<IRequestManagementProps, IRequestManagementState> {

  public constructor(props: IRequestManagementProps) {
    super(props);
    this.state = {
      requests: [], isLoading: false, error: null,
      searchText: '', statusFilter: 'all',
      selectedRequest: null, isViewPanelOpen: false,
      isFormPanelOpen: false,
      formTitle: '', formRequestCode: '',
      formAssignedToLoginName: '', formAssignedToPersonas: [],
      formKey: 0, formError: null, isSubmitting: false,
    };
  }

  public componentDidMount(): void {
    this._fetchRequests().catch(console.error);
  }

  /* ── Data ─────────────────────────────────── */

  private async _fetchRequests(): Promise<void> {
    this.setState({ isLoading: true, error: null });
    try {
      const { context, listName } = this.props;
      const siteUrl = context.pageContext.web.absoluteUrl;
      const encoded = encodeURIComponent(listName);
      const email   = context.pageContext.user.email;

      const url =
        `${siteUrl}/_api/web/lists/getByTitle('${encoded}')/items` +
        `?$select=ID,Title,RequestCode,Status,Author/Title,Created,AssignedTo/EMail,AssignedTo/Title` +
        `&$expand=Author,AssignedTo` +
        `&$orderby=Created%20desc` +
        `&$top=500` +
        `&$filter=AssignedTo/EMail eq '${email}'`;

      const res: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        const b = await res.json().catch(() => ({}));
        throw new Error((b as { error?: { message?: string } })?.error?.message || `HTTP ${res.status}`);
      }
      const data: { value: Record<string, unknown>[] } = await res.json();
      const requests: IRequest[] = (data.value || []).map((i) => ({
        ID:              i.ID as number,
        RequestCode:     (i.RequestCode as string)  || '',
        Title:           (i.Title as string)        || '',
        Status:          (i.Status as string)       || 'Pending',
        AuthorTitle:     ((i.Author  as { Title?: string })?.Title) || '',
        Created:         (i.Created as string)      || '',
        AssignedToEmail: ((i.AssignedTo as { EMail?: string })?.EMail) || '',
        AssignedToTitle: ((i.AssignedTo as { Title?: string })?.Title) || '',
      }));
      this.setState({ requests, isLoading: false });
    } catch (err) {
      this.setState({ error: (err as Error).message, isLoading: false });
    }
  }

  private async _ensureUser(loginName: string): Promise<number> {
    const { context } = this.props;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: JSON.stringify({ logonName: loginName }),
      }
    );
    if (!res.ok) throw new Error(`User not found: ${loginName}`);
    const d = await res.json();
    return (d as { Id: number }).Id;
  }

  private async _submitRequest(): Promise<void> {
    const { formTitle, formRequestCode, formAssignedToLoginName } = this.state;
    if (!formTitle.trim() || !formRequestCode.trim() || !formAssignedToLoginName.trim()) {
      this.setState({ formError: 'Please fill all required fields.' });
      return;
    }
    this.setState({ isSubmitting: true, formError: null });
    try {
      const { context, listName } = this.props;
      const siteUrl  = context.pageContext.web.absoluteUrl;
      const encoded  = encodeURIComponent(listName);
      const userId   = await this._ensureUser(formAssignedToLoginName.trim());
      const res = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getByTitle('${encoded}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify({
            Title: formTitle.trim(),
            RequestCode: formRequestCode.trim(),
            Status: 'Pending',
            AssignedToId: userId,
          }),
        }
      );
      if (!res.ok) {
        const b = await res.json().catch(() => ({}));
        throw new Error((b as { error?: { message?: string } })?.error?.message || `HTTP ${res.status}`);
      }
      this.setState((p) => ({
        isFormPanelOpen: false, formTitle: '', formRequestCode: '',
        formAssignedToLoginName: '', formAssignedToPersonas: [],
        formKey: p.formKey + 1, isSubmitting: false,
      }));
      await this._fetchRequests();
    } catch (err) {
      this.setState({ formError: (err as Error).message, isSubmitting: false });
    }
  }

  private _searchPeople = async (filterText: string): Promise<IPersonaProps[]> => {
    if (!filterText || filterText.length < 2) return [];
    try {
      const { context } = this.props;
      const res = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify({
            queryParams: {
              AllowEmailAddresses: true, AllowMultipleEntities: false,
              AllUrlZones: false, MaximumEntitySuggestions: 10,
              PrincipalSource: 15, PrincipalType: 1,
              QueryString: filterText,
            },
          }),
        }
      );
      if (!res.ok) return [];
      const data = await res.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const results: any[] = JSON.parse(data.value || '[]');
      return results.filter((r) => r.EntityType === 'User').map((r) => ({
        text: r.DisplayText,
        secondaryText: r.EntityData?.Email || '',
        loginName: r.Key,
        imageInitials: (r.DisplayText || '').charAt(0).toUpperCase(),
      }));
    } catch { return []; }
  };

  /* ── Helpers ──────────────────────────────── */

  private _getFiltered(): IRequest[] {
    const { requests, searchText, statusFilter } = this.state;
    const q = searchText.toLowerCase().trim();
    return requests.filter((r) => {
      const matchSearch = !q || r.Title.toLowerCase().includes(q) || r.RequestCode.toLowerCase().includes(q);
      const matchStatus = statusFilter === 'all' || r.Status === statusFilter;
      return matchSearch && matchStatus;
    });
  }

  private _badge(status: string): JSX.Element {
    const s = STATUS_STYLE[status] || { bg: '#F6F8FA', color: TEXT_MUTE, dot: '#ccc' };
    return (
      <span className={styles.badge} style={{ background: s.bg, color: s.color }}>
        <span className={styles.badgeDot} style={{ background: s.dot }} />
        {status || 'Unknown'}
      </span>
    );
  }

  private _date(iso: string): string {
    if (!iso) return '-';
    return new Date(iso).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  /* ── Columns ──────────────────────────────── */

  private _columns(): IColumn[] {
    return [
      { key: 'ID',          name: '#',            fieldName: 'ID',          minWidth: 36,  maxWidth: 48,  isResizable: true },
      { key: 'RequestCode', name: 'Request Code', fieldName: 'RequestCode', minWidth: 100, maxWidth: 130, isResizable: true },
      { key: 'Title',       name: 'Title',        fieldName: 'Title',       minWidth: 200, isResizable: true, isMultiline: true },
      {
        key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 120, isResizable: true,
        onRender: (item: IRequest) => this._badge(item.Status),
      },
      {
        key: 'AssignedTo', name: 'Assigned To', fieldName: 'AssignedToTitle', minWidth: 130, maxWidth: 160, isResizable: true,
        onRender: (item: IRequest) => (
          <span title={item.AssignedToEmail} style={{ color: TEXT_MAIN }}>
            {item.AssignedToTitle || item.AssignedToEmail || '-'}
          </span>
        ),
      },
      {
        key: 'Created', name: 'Created Date', fieldName: 'Created', minWidth: 90, maxWidth: 110, isResizable: true,
        onRender: (item: IRequest) => <span style={{ color: TEXT_MUTE }}>{this._date(item.Created)}</span>,
      },
      {
        key: 'action', name: '', minWidth: 64, maxWidth: 64,
        onRender: (item: IRequest) => (
          <DefaultButton
            text="View"
            styles={{
              root: { minWidth: 54, height: 28, padding: '0 10px', borderRadius: 6, border: `1px solid ${BORDER}`, background: WHITE },
              label: { fontSize: 12, fontWeight: 600, color: PRIMARY },
            }}
            onClick={() => this.setState({ selectedRequest: item, isViewPanelOpen: true })}
          />
        ),
      },
    ];
  }

  /* ── Render pieces ────────────────────────── */

  private _renderCards(): JSX.Element {
    const { requests } = this.state;
    const counts = {
      total:    requests.length,
      pending:  requests.filter((r) => r.Status === 'Pending').length,
      approved: requests.filter((r) => r.Status === 'Approved').length,
      rejected: requests.filter((r) => r.Status === 'Rejected').length,
    };
    return (
      <div className={styles.cardGrid}>
        {CARDS.map((c) => (
          <div key={c.key} className={styles.card}>
            <div className={styles.cardIcon} style={{ background: c.bg }}>
              <FontIcon iconName={c.icon} style={{ color: c.color }} />
            </div>
            <div>
              <div className={styles.cardValue}>{counts[c.key as keyof typeof counts]}</div>
              <div className={styles.cardLabel}>{c.label}</div>
            </div>
          </div>
        ))}
      </div>
    );
  }

  private _renderViewModal(): React.ReactNode {
    const { selectedRequest, isViewPanelOpen } = this.state;
    if (!selectedRequest) return null;
    const close = (): void => this.setState({ isViewPanelOpen: false, selectedRequest: null });
    const fields = [
      { k: 'Request Code', v: selectedRequest.RequestCode || '-' },
      { k: 'Title',        v: selectedRequest.Title               },
      { k: 'Status',       v: this._badge(selectedRequest.Status) },
      { k: 'Assigned To',  v: selectedRequest.AssignedToTitle || selectedRequest.AssignedToEmail || '-' },
      { k: 'Created By',   v: selectedRequest.AuthorTitle         },
      { k: 'Created Date', v: this._date(selectedRequest.Created) },
    ];
    return (
      <Modal isOpen={isViewPanelOpen} onDismiss={close} isBlocking={false}>
        <div className={styles.modalWrap}>
          <div className={styles.modalHead}>
            <div>
              <div className={styles.modalHeadTitle}>{selectedRequest.RequestCode || selectedRequest.Title}</div>
              <div style={{ fontSize: 12, color: TEXT_MUTE, marginTop: 2 }}>Request Details</div>
            </div>
            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Close" onClick={close}
              styles={{ root: { borderRadius: 8, color: TEXT_MUTE } }} />
          </div>
          <div className={styles.modalBody}>
            {fields.map((f) => (
              <div key={f.k} className={styles.detailRow}>
                <span className={styles.detailKey}>{f.k}</span>
                <span className={styles.detailVal}>{f.v}</span>
              </div>
            ))}
          </div>
          <div className={styles.modalFoot}>
            <DefaultButton text="Close" onClick={close}
              styles={{ root: { borderRadius: 8 } }} />
          </div>
        </div>
      </Modal>
    );
  }

  private _renderFormModal(): React.ReactNode {
    const { isFormPanelOpen, formTitle, formRequestCode, formAssignedToPersonas, formKey, formError, isSubmitting } = this.state;
    const close = (): void => this.setState({ isFormPanelOpen: false, formError: null });
    return (
      <Modal isOpen={isFormPanelOpen} onDismiss={close} isBlocking={false}>
        <div className={styles.modalWrap}>
          <div className={styles.modalHead}>
            <div>
              <div className={styles.modalHeadTitle}>New Request</div>
              <div style={{ fontSize: 12, color: TEXT_MUTE, marginTop: 2 }}>Fill in the details below</div>
            </div>
            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Close" onClick={close}
              disabled={isSubmitting} styles={{ root: { borderRadius: 8, color: TEXT_MUTE } }} />
          </div>
          <div className={styles.modalBody}>
            <Stack tokens={{ childrenGap: 18 }}>
              {formError && (
                <MessageBar messageBarType={MessageBarType.error}
                  onDismiss={() => this.setState({ formError: null })}
                  styles={{ root: { borderRadius: 8 } }}>
                  {formError}
                </MessageBar>
              )}
              <TextField label="Title" required value={formTitle}
                onChange={(_, v) => this.setState({ formTitle: v || '' })}
                disabled={isSubmitting}
                styles={{ fieldGroup: { borderRadius: 8 } }} />
              <TextField label="Request Code" required value={formRequestCode}
                onChange={(_, v) => this.setState({ formRequestCode: v || '' })}
                disabled={isSubmitting}
                styles={{ fieldGroup: { borderRadius: 8 } }} />
              <div>
                <Label required>Assigned To</Label>
                <NormalPeoplePicker
                  key={formKey}
                  onResolveSuggestions={this._searchPeople}
                  selectedItems={formAssignedToPersonas as IPersonaProps[]}
                  itemLimit={1}
                  disabled={isSubmitting}
                  pickerSuggestionsProps={{
                    suggestionsHeaderText: 'Suggested people',
                    noResultsFoundText: 'No results found',
                    loadingText: 'Loading...',
                  }}
                  onChange={(items?: IPersonaProps[]) => {
                    const ps = (items || []) as Array<IPersonaProps & { loginName?: string }>;
                    this.setState({
                      formAssignedToPersonas: ps.map((p) => ({ text: p.text || '', secondaryText: p.secondaryText || '', loginName: p.loginName || p.secondaryText || '' })),
                      formAssignedToLoginName: ps.length > 0 ? (ps[0].loginName || ps[0].secondaryText || '') : '',
                    });
                  }}
                />
              </div>
            </Stack>
          </div>
          <div className={styles.modalFoot}>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text={isSubmitting ? 'Submitting…' : 'Submit Request'}
                disabled={isSubmitting}
                onClick={() => this._submitRequest().catch(console.error)}
                styles={{ root: { borderRadius: 8, paddingLeft: 20, paddingRight: 20 } }}
                onRenderIcon={isSubmitting ? () => <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginRight: 6 } }} /> : undefined}
              />
              <DefaultButton text="Cancel" disabled={isSubmitting} onClick={close}
                styles={{ root: { borderRadius: 8 } }} />
            </Stack>
          </div>
        </div>
      </Modal>
    );
  }

  /* ── Main render ──────────────────────────── */

  public render(): React.ReactElement<IRequestManagementProps> {
    const { isLoading, error, searchText, statusFilter, requests } = this.state;
    const filtered = this._getFiltered();

    return (
      <div className={styles.root}>

        {/* Header */}
        <div className={styles.header}>
          <div>
            <div className={styles.headerTitle}>Request Management</div>
            <div className={styles.headerSub}>{this.props.listName} — {requests.length} record(s)</div>
          </div>
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            text="New Request"
            onClick={() => this.setState((p) => ({
              isFormPanelOpen: true, formTitle: '', formRequestCode: '',
              formAssignedToLoginName: '', formAssignedToPersonas: [],
              formKey: p.formKey + 1, formError: null,
            }))}
            styles={{ root: { borderRadius: 8, height: 36, paddingLeft: 16, paddingRight: 16 } }}
          />
        </div>

        {/* Error */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 16 } }}
            onDismiss={() => this.setState({ error: null })}
            dismissButtonAriaLabel="Dismiss">
            {error}
          </MessageBar>
        )}

        {isLoading ? (
          <Stack horizontalAlign="center" styles={{ root: { padding: '80px 0' } }}>
            <Spinner size={SpinnerSize.large} label="Loading…" />
          </Stack>
        ) : (
          <>
            {this._renderCards()}

            {/* Toolbar */}
            <div className={styles.toolbar}>
              <div style={{ flex: 3, minWidth: 200 }}>
                <SearchBox
                  placeholder="Search by title or request code…"
                  value={searchText}
                  onChange={(_, v) => this.setState({ searchText: v || '' })}
                  onClear={() => this.setState({ searchText: '' })}
                  styles={{ root: { borderRadius: 8 } }}
                />
              </div>
              <div style={{ flex: 1, minWidth: 150 }}>
                <Dropdown
                  options={STATUS_OPTIONS}
                  selectedKey={statusFilter}
                  onChange={(_, o) => this.setState({ statusFilter: String(o?.key ?? 'all') })}
                  styles={{ dropdown: { borderRadius: 8 } }}
                />
              </div>
              <IconButton
                iconProps={{ iconName: 'Refresh' }}
                title="Refresh"
                onClick={() => this._fetchRequests().catch(console.error)}
                styles={{ root: { borderRadius: 8, border: `1px solid ${BORDER}`, background: WHITE } }}
              />
            </div>

            {/* Table */}
            <div className={styles.tableWrapper}>
              <DetailsList
                items={filtered}
                columns={this._columns()}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                isHeaderVisible={true}
                styles={{
                  root: { selectors: { '.ms-DetailsRow': { borderBottom: `1px solid ${BORDER}` } } },
                  headerWrapper: { selectors: { '.ms-DetailsHeader': { paddingTop: 0, background: '#FAFBFC', borderBottom: `1px solid ${BORDER}` } } },
                }}
              />
              {filtered.length === 0 && (
                <div className={styles.emptyState}>
                  <FontIcon iconName="InboxCheck" style={{ fontSize: 40, color: '#C0C5CC', display: 'block', marginBottom: 12 }} />
                  <Text styles={{ root: { color: TEXT_MUTE, fontSize: 14 } }}>
                    {requests.length === 0 ? 'No requests found.' : 'No results match your filter.'}
                  </Text>
                </div>
              )}
            </div>
          </>
        )}

        {this._renderViewModal()}
        {this._renderFormModal()}
      </div>
    );
  }
}
