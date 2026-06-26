import * as React from 'react';
import {
  NormalPeoplePicker, IPersonaProps,
  Dropdown, IDropdownOption, TextField,
  PrimaryButton, DefaultButton, IconButton,
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  mergeStyleSets, FontIcon, Dialog, DialogType, DialogFooter,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { UserRoleService } from '../../services/userRoleService';
import { IUserRole } from '../../models/UserRole';
import { ROLE, ROLE_LABEL } from '../../common/constants/roles';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import { PageHeader } from '../shared/PageHeader';

const PRIMARY = '#0078D4';
const BORDER = '#E1E4E8';
const TEXT_MUTE = '#6A737D';
const WHITE = '#FFFFFF';

const styles = mergeStyleSets({
  tableWrap: {
    background: WHITE, borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden', marginTop: 16,
  },
  toolbar: {
    padding: '12px 16px', borderBottom: `1px solid ${BORDER}`,
    display: 'flex' as const, justifyContent: 'space-between' as const,
    alignItems: 'center' as const, flexWrap: 'wrap' as const, gap: 10,
  },
  toolbarLeft: { fontSize: 13, color: TEXT_MUTE },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 560 },
  th: {
    padding: '10px 16px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 16px', fontSize: 13, borderBottom: `1px solid ${BORDER}`, verticalAlign: 'middle' as const },
  empty: { padding: '40px 0', textAlign: 'center' as const, color: TEXT_MUTE, fontSize: 14 },
  roleBadge: {
    display: 'inline-block' as const,
    padding: '2px 10px', borderRadius: 20, fontSize: 11, fontWeight: 700,
  },
  activeChip: {
    display: 'inline-block' as const,
    padding: '2px 10px', borderRadius: 20, fontSize: 11, fontWeight: 600,
  },
  mobileCard: {
    padding: '12px 16px', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
  },
  mobileName: { fontWeight: 700, fontSize: 14, marginBottom: 2 },
  mobileLine: { fontSize: 12, color: TEXT_MUTE, marginBottom: 2 },
  mobileActions: { marginTop: 8, display: 'flex' as const, gap: 8, alignItems: 'center' as const },
  formSection: { marginBottom: 16 },
  formLabel: { fontSize: 14, fontWeight: 600, color: '#24292E', marginBottom: 6 },
  pickerBox: {
    border: '1px solid #C8C6C4', borderRadius: 4,
    minHeight: 38, background: WHITE, padding: '2px 4px',
  },
  helpText: { fontSize: 12, color: TEXT_MUTE, marginTop: 4 },
  row2: { display: 'grid' as const, gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 },
});

const ROLE_COLORS: Record<string, { bg: string; color: string }> = {
  [ROLE.ADMIN]:           { bg: '#FEE2E2', color: '#991B1B' },
  [ROLE.MANAGER]:         { bg: '#DBEAFE', color: '#1E40AF' },
  [ROLE.TRANSPORT_ADMIN]: { bg: '#D1FAE5', color: '#065F46' },
  [ROLE.DRIVER]:          { bg: '#FEF3C7', color: '#92400E' },
  [ROLE.REQUESTER]:       { bg: '#F3F4F6', color: '#374151' },
};

const roleOptions: IDropdownOption[] = [
  ROLE.REQUESTER, ROLE.MANAGER, ROLE.TRANSPORT_ADMIN, ROLE.DRIVER, ROLE.ADMIN,
].map(r => ({ key: r, text: ROLE_LABEL[r] || r }));

interface ISpPickerUser {
  Key: string;
  DisplayText: string;
  EntityData?: { Email?: string; Department?: string; Title?: string };
}

interface IUserRoleSetupProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen) => void;
}

interface IFormState {
  pickedPersona: IPersonaProps | null;
  pickedEmail: string;
  pickedName: string;
  role: string;
  department: string;
}

interface IUserRoleSetupState {
  users: IUserRole[];
  isLoading: boolean;
  isSaving: boolean;
  showDialog: boolean;
  editTarget: IUserRole | null;
  form: IFormState;
  error: string | null;
  successMsg: string | null;
  isMobile: boolean;
}

const EMPTY_FORM: IFormState = {
  pickedPersona: null, pickedEmail: '', pickedName: '',
  role: ROLE.REQUESTER, department: '',
};

export default class UserRoleSetup extends React.Component<IUserRoleSetupProps, IUserRoleSetupState> {
  private userSvc: UserRoleService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IUserRoleSetupProps) {
    super(props);
    this.state = {
      users: [], isLoading: false, isSaving: false,
      showDialog: false, editTarget: null,
      form: { ...EMPTY_FORM },
      error: null, successMsg: null,
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 768,
    };
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
      const users = await this.userSvc.getAllIncludingInactive();
      this.setState({ users, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  // Search users trong SharePoint tenant qua ClientPeoplePicker REST API
  private _searchSpUsers = async (filter: string): Promise<IPersonaProps[]> => {
    if (!filter || filter.length < 2) return [];
    try {
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      const url = `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
      const body = JSON.stringify({
        queryParams: {
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: 10,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: filter,
        },
      });
      const res = await this.props.context.spHttpClient.post(
        url, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body,
        }
      );
      if (!res.ok) return [];
      const data = await res.json() as { value: string };
      const results: ISpPickerUser[] = JSON.parse(data.value) as ISpPickerUser[];
      return results
        .filter(u => u.EntityData && u.EntityData.Email)
        .map(u => ({
          key: u.EntityData!.Email!,
          text: u.DisplayText || u.EntityData!.Title || u.EntityData!.Email!,
          secondaryText: u.EntityData!.Department
            ? `${u.EntityData!.Department} — ${u.EntityData!.Email}`
            : u.EntityData!.Email,
          imageInitials: (u.DisplayText || '?').split(' ').map((n: string) => n[0]).join('').toUpperCase().slice(0, 2),
        }));
    } catch {
      return [];
    }
  }

  private _onPickerChange = (items?: IPersonaProps[]): void => {
    if (items && items.length > 0) {
      const p = items[0];
      const email = String(p.key || '');
      const name = p.text || '';
      const dept = (p.secondaryText || '').split(' — ')[0] !== email
        ? (p.secondaryText || '').split(' — ')[0]
        : '';
      this.setState(prev => ({
        form: {
          ...prev.form,
          pickedPersona: p,
          pickedEmail: email,
          pickedName: name,
          department: prev.form.department || dept,
        },
      }));
    } else {
      this.setState(prev => ({
        form: { ...prev.form, pickedPersona: null, pickedEmail: '', pickedName: '' },
      }));
    }
  }

  private _openAdd(): void {
    this.setState({ showDialog: true, editTarget: null, form: { ...EMPTY_FORM } });
  }

  private _openEdit(u: IUserRole): void {
    const persona: IPersonaProps = {
      key: u.UserEmail,
      text: u.UserName,
      secondaryText: u.Department ? `${u.Department} — ${u.UserEmail}` : u.UserEmail,
      imageInitials: u.UserName ? u.UserName.split(' ').map((n: string) => n[0]).join('').toUpperCase().slice(0, 2) : '?',
    };
    this.setState({
      showDialog: true,
      editTarget: u,
      form: {
        pickedPersona: persona,
        pickedEmail: u.UserEmail,
        pickedName: u.UserName,
        role: u.Role,
        department: u.Department,
      },
    });
  }

  private async _save(): Promise<void> {
    const { form, editTarget } = this.state;
    if (!form.pickedEmail || !form.role) {
      this.setState({ error: 'Vui lòng chọn người dùng và vai trò.' });
      return;
    }
    this.setState({ isSaving: true, error: null });
    try {
      if (editTarget) {
        await this.userSvc.update(editTarget.ID, {
          Role: form.role,
          Department: form.department,
          UserName: form.pickedName || editTarget.UserName,
          IsActive: true,
        });
        this.setState({ successMsg: 'Đã cập nhật phân quyền thành công.' });
      } else {
        await this.userSvc.create({
          Title: form.pickedName || form.pickedEmail,
          UserId: form.pickedEmail,
          UserName: form.pickedName,
          UserEmail: form.pickedEmail,
          Role: form.role,
          Department: form.department,
          IsActive: true,
        });
        this.setState({ successMsg: 'Đã thêm người dùng thành công.' });
      }
      this.setState({ showDialog: false, editTarget: null, form: { ...EMPTY_FORM }, isSaving: false });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  private async _deactivate(id: number): Promise<void> {
    try {
      await this.userSvc.deactivate(id);
      this.setState({ successMsg: 'Đã vô hiệu hóa người dùng.' });
      await this._load();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    }
  }

  private _roleBadge(role: string): React.ReactElement {
    const c = ROLE_COLORS[role] || { bg: '#F3F4F6', color: '#374151' };
    return (
      <span className={styles.roleBadge} style={{ background: c.bg, color: c.color }}>
        {ROLE_LABEL[role] || role}
      </span>
    );
  }

  public render(): React.ReactElement {
    const { users, isLoading, isSaving, showDialog, editTarget, form, error, successMsg, isMobile } = this.state;

    return (
      <div>
        <PageHeader icon="PeopleAdd" title="Phân quyền người dùng" subtitle="Gán vai trò cho người dùng trong hệ thống" />

        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ error: null })}>
            {error}
          </MessageBar>
        )}
        {successMsg && (
          <MessageBar messageBarType={MessageBarType.success}
            styles={{ root: { borderRadius: 8, marginBottom: 12 } }}
            onDismiss={() => this.setState({ successMsg: null })}>
            {successMsg}
          </MessageBar>
        )}

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <div className={styles.tableWrap}>
            <div className={styles.toolbar}>
              <span className={styles.toolbarLeft}>{users.length} người dùng đã phân quyền</span>
              <PrimaryButton
                iconProps={{ iconName: 'Add' }} text="Thêm người dùng"
                onClick={() => this._openAdd()}
                styles={{ root: { borderRadius: 8 } }}
              />
            </div>

            {users.length === 0 ? (
              <div className={styles.empty}>
                <FontIcon iconName="PeopleAdd" style={{ fontSize: 36, color: '#C8C6C4', display: 'block', marginBottom: 10 }} />
                Chưa có người dùng nào được phân quyền
              </div>
            ) : isMobile ? (
              <div>
                {users.map(u => (
                  <div key={u.ID} className={styles.mobileCard}>
                    <div className={styles.mobileName}>{u.UserName || u.UserEmail}</div>
                    <div className={styles.mobileLine}>{u.UserEmail}</div>
                    <div className={styles.mobileLine}>{u.Department && <span>{u.Department} · </span>}{this._roleBadge(u.Role)}</div>
                    <div className={styles.mobileActions}>
                      <span className={styles.activeChip}
                        style={{ background: u.IsActive ? '#D1FAE5' : '#FEE2E2', color: u.IsActive ? '#065F46' : '#991B1B' }}>
                        {u.IsActive ? 'Hoạt động' : 'Không HĐ'}
                      </span>
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Sửa"
                        onClick={() => this._openEdit(u)}
                        styles={{ root: { color: PRIMARY } }} />
                      {u.IsActive && (
                        <IconButton iconProps={{ iconName: 'UserRemove' }} title="Vô hiệu hóa"
                          onClick={() => this._deactivate(u.ID).catch(console.error)}
                          styles={{ root: { color: '#C53030' } }} />
                      )}
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className={styles.tableScroll}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th className={styles.th}>Họ tên</th>
                      <th className={styles.th}>Email</th>
                      <th className={styles.th}>Phòng ban</th>
                      <th className={styles.th}>Vai trò</th>
                      <th className={styles.th}>Trạng thái</th>
                      <th className={styles.th}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {users.map(u => (
                      <tr key={u.ID}>
                        <td className={styles.td} style={{ fontWeight: 600 }}>{u.UserName || '—'}</td>
                        <td className={styles.td} style={{ color: TEXT_MUTE }}>{u.UserEmail}</td>
                        <td className={styles.td}>{u.Department || '—'}</td>
                        <td className={styles.td}>{this._roleBadge(u.Role)}</td>
                        <td className={styles.td}>
                          <span className={styles.activeChip}
                            style={{ background: u.IsActive ? '#D1FAE5' : '#FEE2E2', color: u.IsActive ? '#065F46' : '#991B1B' }}>
                            {u.IsActive ? 'Hoạt động' : 'Không HĐ'}
                          </span>
                        </td>
                        <td className={styles.td} style={{ whiteSpace: 'nowrap' as const }}>
                          <IconButton iconProps={{ iconName: 'Edit' }} title="Sửa vai trò"
                            onClick={() => this._openEdit(u)}
                            styles={{ root: { color: PRIMARY } }} />
                          {u.IsActive && (
                            <IconButton iconProps={{ iconName: 'UserRemove' }} title="Vô hiệu hóa"
                              onClick={() => this._deactivate(u.ID).catch(console.error)}
                              styles={{ root: { color: '#C53030' } }} />
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* Dialog thêm / sửa */}
        <Dialog
          hidden={!showDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: editTarget ? 'Sửa phân quyền' : 'Thêm người dùng',
            subText: editTarget
              ? 'Cập nhật vai trò và phòng ban cho người dùng.'
              : 'Tìm người dùng trong hệ thống và gán vai trò.',
          }}
          modalProps={{ isBlocking: false }}
          onDismiss={() => this.setState({ showDialog: false })}
          minWidth={480}
        >
          {/* People Picker — chỉ hiện khi thêm mới */}
          {!editTarget && (
            <div className={styles.formSection}>
              <div className={styles.formLabel}>Chọn người dùng *</div>
              <div className={styles.pickerBox}>
                <NormalPeoplePicker
                  onResolveSuggestions={this._searchSpUsers}
                  onChange={this._onPickerChange}
                  itemLimit={1}
                  selectedItems={form.pickedPersona ? [form.pickedPersona] : []}
                  pickerSuggestionsProps={{
                    suggestionsHeaderText: 'Người dùng gợi ý',
                    noResultsFoundText: 'Không tìm thấy. Thử nhập tên hoặc email.',
                    loadingText: 'Đang tìm...',
                  }}
                  inputProps={{
                    placeholder: 'Nhập tên hoặc email để tìm kiếm...',
                    'aria-label': 'Tìm người dùng',
                  }}
                  styles={{
                    root: { width: '100%' },
                    text: { border: 'none', minWidth: 0 },
                  }}
                />
              </div>
              <div className={styles.helpText}>
                Tìm kiếm người dùng trong tổ chức theo tên hoặc email.
              </div>
            </div>
          )}

          {/* Khi edit: hiện thông tin read-only */}
          {editTarget && (
            <div className={styles.formSection}>
              <div style={{
                background: '#F6F8FA', borderRadius: 8, padding: '10px 14px',
                display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16,
              }}>
                <div style={{
                  width: 40, height: 40, borderRadius: '50%', background: PRIMARY,
                  display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
                }}>
                  <FontIcon iconName="Contact" style={{ fontSize: 20, color: '#fff' }} />
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>{editTarget.UserName || editTarget.UserEmail}</div>
                  <div style={{ fontSize: 12, color: TEXT_MUTE }}>{editTarget.UserEmail}</div>
                </div>
              </div>
            </div>
          )}

          <div className={styles.row2}>
            <Dropdown
              label="Vai trò *"
              options={roleOptions}
              selectedKey={form.role || null}
              onChange={(_, o) => this.setState(prev => ({ form: { ...prev.form, role: String(o?.key || '') } }))}
            />
            <TextField
              label="Phòng ban"
              value={form.department}
              onChange={(_, v) => this.setState(prev => ({ form: { ...prev.form, department: v || '' } }))}
              placeholder="VD: IT, Kế toán, Kinh doanh..."
            />
          </div>

          {/* Preview vai trò */}
          {form.role && (
            <div style={{ marginBottom: 14, fontSize: 13, color: TEXT_MUTE }}>
              Quyền truy cập: {this._roleBadge(form.role)}
            </div>
          )}

          <DialogFooter>
            <PrimaryButton
              text={isSaving ? 'Đang lưu...' : editTarget ? 'Cập nhật' : 'Thêm'}
              disabled={isSaving || !form.pickedEmail || !form.role}
              onClick={() => this._save().catch(console.error)}
            />
            <DefaultButton text="Hủy" onClick={() => this.setState({ showDialog: false })} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
