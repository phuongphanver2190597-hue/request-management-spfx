import * as React from 'react';
import {
  TextField, IDropdownOption, Toggle, PrimaryButton, DefaultButton,
  Spinner, SpinnerSize, MessageBar, MessageBarType, mergeStyleSets, FontIcon,
  NormalPeoplePicker, IPersonaProps, Dropdown,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { LocationMasterService } from '../../services/locationMasterService';
import { UserRoleService } from '../../services/userRoleService';
import { ILocationMaster } from '../../models/LocationMaster';
import { IUserRole } from '../../models/UserRole';
import { VEHICLE_TYPES } from '../../common/constants/statuses';
import { ROLE } from '../../common/constants/roles';
import { ICreateBookingRequest } from '../../models/VehicleBookingRequest';
import { validateBookingRequest, getFieldError, IValidationError } from '../../common/helpers/validationHelper';
import { fromInputDateTimeLocal } from '../../common/helpers/dateHelper';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import { PageHeader } from '../shared/PageHeader';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const PRIMARY = '#0078D4';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  card: {
    background: WHITE, borderRadius: 12,
    boxShadow: '0 1px 4px rgba(0,0,0,0.08)', padding: '20px',
  },
  section: {
    borderBottom: `1px solid ${BORDER}`, paddingBottom: 18, marginBottom: 18,
  },
  sectionTitle: {
    fontSize: 12, fontWeight: 700, color: '#6A737D', textTransform: 'uppercase' as const,
    letterSpacing: '0.4px', marginBottom: 12,
  },
  row2: { display: 'grid' as const, gridTemplateColumns: '1fr 1fr', gap: 14 },
  row1: { display: 'grid' as const, gridTemplateColumns: '1fr', gap: 14 },
  actions: { display: 'flex' as const, gap: 10, paddingTop: 8, flexWrap: 'wrap' as const },
  approverCard: {
    background: '#F0F7FF',
    border: `1px solid #C7E0F4`,
    borderRadius: 10,
    padding: '14px 16px',
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 12,
    marginTop: 12,
  },
  approverIcon: {
    width: 40, height: 40, borderRadius: '50%',
    background: PRIMARY,
    display: 'flex' as const, alignItems: 'center' as const, justifyContent: 'center' as const,
    flexShrink: 0,
  },
  approverName: { fontSize: 14, fontWeight: 700, color: '#1B1F23' },
  approverMeta: { fontSize: 12, color: TEXT_MUTE, marginTop: 2 },
  noApproverWarn: {
    background: '#FFF8E6', border: `1px solid #F3D55B`,
    borderRadius: 10, padding: '12px 14px',
    display: 'flex' as const, alignItems: 'center' as const, gap: 10,
    marginTop: 12, fontSize: 13, color: '#7A5800',
  },
  pickerWrap: {
    marginBottom: 12,
  },
  pickerLabel: {
    fontSize: 14, fontWeight: 600, color: '#24292E', marginBottom: 6,
  },
  pickerBox: {
    border: '1px solid #C8C6C4',
    borderRadius: 4,
    minHeight: 38,
    background: '#fff',
    padding: '2px 4px',
  },
});

const vehicleTypeOptions: IDropdownOption[] = VEHICLE_TYPES.map(v => ({ key: v, text: v }));

interface ICreateRequestFormProps {
  context: WebPartContext;
  user: IAppUser;
  editId?: number;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IFormState {
  pickupLocation: string;
  dropoffLocation: string;
  pickupDateTimeLocal: string;
  returnDateTimeLocal: string;
  isRoundTrip: boolean;
  numberOfPassengers: number;
  purpose: string;
  vehicleType: string;
  specialRequirement: string;
  phoneNumber: string;
  department: string;
  selectedManagerEmail: string;
}

interface ICreateRequestFormState {
  form: IFormState;
  locations: ILocationMaster[];
  managers: IUserRole[];
  allManagers: IUserRole[];
  errors: IValidationError[];
  isSaving: boolean;
  isSubmitting: boolean;
  globalError: string | null;
  successMsg: string | null;
  isMobile: boolean;
}

export default class CreateRequestForm extends React.Component<ICreateRequestFormProps, ICreateRequestFormState> {
  private bookingSvc: VehicleBookingRequestService;
  private locationSvc: LocationMasterService;
  private userSvc: UserRoleService;
  private _mql: MediaQueryList | null = null;

  constructor(props: ICreateRequestFormProps) {
    super(props);
    this.state = {
      form: {
        pickupLocation: '', dropoffLocation: '',
        pickupDateTimeLocal: '', returnDateTimeLocal: '',
        isRoundTrip: false, numberOfPassengers: 1,
        purpose: '', vehicleType: '', specialRequirement: '',
        phoneNumber: props.user.userId, department: props.user.department,
        selectedManagerEmail: '',
      },
      locations: [], managers: [], allManagers: [], errors: [],
      isSaving: false, isSubmitting: false,
      globalError: null, successMsg: null,
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 600,
    };
    this.bookingSvc = new VehicleBookingRequestService(props.context);
    this.locationSvc = new LocationMasterService(props.context);
    this.userSvc = new UserRoleService(props.context);
  }

  public componentDidMount(): void {
    this._loadMeta().catch(console.error);
    if (typeof window !== 'undefined') {
      this._mql = window.matchMedia('(max-width: 600px)');
      this._mql.addListener(this._onMql);
    }
  }

  public componentWillUnmount(): void {
    if (this._mql) this._mql.removeListener(this._onMql);
  }

  private _onMql = (e: MediaQueryListEvent): void => { this.setState({ isMobile: e.matches }); }

  private async _loadMeta(): Promise<void> {
    const { user } = this.props;
    const [locations, managers, allManagers] = await Promise.all([
      this.locationSvc.getAll(),
      this.userSvc.getManagersForDepartment(user.department),
      this.userSvc.getAllManagers(),
    ]);

    const effectiveManagers = managers.length > 0 ? managers : allManagers;

    this.setState({
      locations,
      managers: effectiveManagers,
      allManagers,
    });
  }

  private _setField<K extends keyof IFormState>(key: K, value: IFormState[K]): void {
    this.setState(prev => ({ form: { ...prev.form, [key]: value }, errors: [] }));
  }

  private _buildPayload(): ICreateBookingRequest {
    const { form } = this.state;
    return {
      pickupLocation:    form.pickupLocation,
      dropoffLocation:   form.dropoffLocation,
      pickupDateTime:    fromInputDateTimeLocal(form.pickupDateTimeLocal),
      returnDateTime:    form.isRoundTrip ? fromInputDateTimeLocal(form.returnDateTimeLocal) : '',
      isRoundTrip:       form.isRoundTrip,
      numberOfPassengers:form.numberOfPassengers,
      purpose:           form.purpose,
      vehicleType:       form.vehicleType,
      specialRequirement:form.specialRequirement,
      phoneNumber:       form.phoneNumber,
      department:        form.department,
    };
  }

  private _validate(): boolean {
    const { form } = this.state;
    const errors = validateBookingRequest({
      pickupLocation:    form.pickupLocation,
      dropoffLocation:   form.dropoffLocation,
      pickupDateTime:    form.pickupDateTimeLocal,
      returnDateTime:    form.returnDateTimeLocal,
      isRoundTrip:       form.isRoundTrip,
      numberOfPassengers:form.numberOfPassengers,
      purpose:           form.purpose,
      vehicleType:       form.vehicleType,
      phoneNumber:       form.phoneNumber,
    });
    this.setState({ errors });
    return errors.length === 0;
  }

  private async _saveDraft(): Promise<void> {
    this.setState({ isSaving: true, globalError: null, successMsg: null });
    try {
      const { user } = this.props;
      await this.bookingSvc.createDraft(this._buildPayload(), {
        id: user.userId, name: user.userName, email: user.userEmail,
      });
      this.setState({ isSaving: false, successMsg: 'Đã lưu nháp thành công!' });
    } catch (err) {
      this.setState({ isSaving: false, globalError: extractErrorMessage(err) });
    }
  }

  private async _submit(): Promise<void> {
    if (!this._validate()) return;
    const { form } = this.state;
    const managerEmail = form.selectedManagerEmail;

    if (!managerEmail && this.props.user.role !== ROLE.MANAGER) {
      this.setState({ globalError: 'Vui lòng chọn người phê duyệt trước khi gửi yêu cầu.' });
      return;
    }

    this.setState({ isSubmitting: true, globalError: null, successMsg: null });
    try {
      const { user } = this.props;
      const created = await this.bookingSvc.createDraft(this._buildPayload(), {
        id: user.userId, name: user.userName, email: user.userEmail,
      });
      await this.bookingSvc.submitRequest(created.ID, created, {
        id: user.userId, name: user.userName, email: user.userEmail,
      }, managerEmail || user.userEmail);
      this.setState({ isSubmitting: false });
      this.props.onNavigate('my-requests');
    } catch (err) {
      this.setState({ isSubmitting: false, globalError: extractErrorMessage(err) });
    }
  }

  private _getSelectedManager(): IUserRole | undefined {
    const { allManagers, form } = this.state;
    return allManagers.find(m => m.UserEmail === form.selectedManagerEmail);
  }

  private _managerToPersona(m: IUserRole): IPersonaProps {
    return {
      key: m.UserEmail,
      text: m.UserName,
      secondaryText: m.Department ? `${m.Department} — ${m.UserEmail}` : m.UserEmail,
      imageInitials: m.UserName ? m.UserName.split(' ').map((n: string) => n[0]).join('').toUpperCase().slice(0, 2) : '?',
    };
  }

  private _onResolveSuggestions = (filter: string, selectedItems?: IPersonaProps[]): IPersonaProps[] => {
    const { allManagers } = this.state;
    const q = filter.toLowerCase();
    const selectedKeys: string[] = (selectedItems || []).map(p => p.key as string);
    return allManagers
      .filter(m =>
        selectedKeys.indexOf(m.UserEmail) === -1 &&
        (m.UserName.toLowerCase().indexOf(q) !== -1 || m.UserEmail.toLowerCase().indexOf(q) !== -1 || m.Department.toLowerCase().indexOf(q) !== -1)
      )
      .map(m => this._managerToPersona(m));
  }

  private _onPickerChange = (items?: IPersonaProps[]): void => {
    const email = (items && items.length > 0) ? String(items[0].key || '') : '';
    this._setField('selectedManagerEmail', email);
  }

  public render(): React.ReactElement {
    const { form, locations, managers, allManagers, errors, isSaving, isSubmitting, globalError, successMsg, isMobile } = this.state;
    const busy = isSaving || isSubmitting;
    const rowClass = isMobile ? styles.row1 : styles.row2;
    const selectedManager = this._getSelectedManager();
    const hasSameDeptManagers = managers.length > 0 && managers[0].Department === this.props.user.department;

    const locationOptions: IDropdownOption[] = locations.map(l => ({ key: l.LocationName, text: l.LocationName }));

    return (
      <div>
        <PageHeader icon="Add" title="Tạo yêu cầu đặt xe" subtitle="Điền đầy đủ thông tin chuyến đi" />

        {globalError && (
          <MessageBar messageBarType={MessageBarType.error}
            styles={{ root: { borderRadius: 8, marginBottom: 16 } }}
            onDismiss={() => this.setState({ globalError: null })}>
            {globalError}
          </MessageBar>
        )}
        {successMsg && (
          <MessageBar messageBarType={MessageBarType.success}
            styles={{ root: { borderRadius: 8, marginBottom: 16 } }}
            onDismiss={() => this.setState({ successMsg: null })}>
            {successMsg}
          </MessageBar>
        )}

        <div className={styles.card}>
          {/* Thông tin người yêu cầu */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Thông tin người yêu cầu</div>
            <div className={rowClass}>
              <TextField label="Họ tên" value={this.props.user.userName} disabled />
              <TextField label="Email" value={this.props.user.userEmail} disabled />
              <TextField label="Số điện thoại liên hệ" required value={form.phoneNumber}
                onChange={(_, v) => this._setField('phoneNumber', v || '')}
                errorMessage={getFieldError(errors, 'phoneNumber')} disabled={busy} />
              <TextField label="Phòng ban" value={form.department}
                onChange={(_, v) => this._setField('department', v || '')} disabled={busy} />
            </div>
          </div>

          {/* Thông tin chuyến đi */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Thông tin chuyến đi</div>
            <div className={rowClass}>
              <Dropdown label="Điểm đón *" options={[{ key: '', text: '-- Chọn điểm đón --' }, ...locationOptions]}
                selectedKey={form.pickupLocation}
                onChange={(_, o) => this._setField('pickupLocation', String(o?.key || ''))}
                errorMessage={getFieldError(errors, 'pickupLocation')} disabled={busy} />
              <Dropdown label="Điểm đến *" options={[{ key: '', text: '-- Chọn điểm đến --' }, ...locationOptions]}
                selectedKey={form.dropoffLocation}
                onChange={(_, o) => this._setField('dropoffLocation', String(o?.key || ''))}
                errorMessage={getFieldError(errors, 'dropoffLocation')} disabled={busy} />
              <TextField label="Thời gian đón *" type="datetime-local"
                value={form.pickupDateTimeLocal}
                onChange={(_, v) => this._setField('pickupDateTimeLocal', v || '')}
                errorMessage={getFieldError(errors, 'pickupDateTime')} disabled={busy} />
              <div>
                <Toggle label="Chuyến khứ hồi" checked={form.isRoundTrip}
                  onChange={(_, v) => this._setField('isRoundTrip', !!v)} disabled={busy} />
                {form.isRoundTrip && (
                  <TextField label="Thời gian về *" type="datetime-local"
                    value={form.returnDateTimeLocal}
                    onChange={(_, v) => this._setField('returnDateTimeLocal', v || '')}
                    errorMessage={getFieldError(errors, 'returnDateTime')} disabled={busy} />
                )}
              </div>
            </div>
            <div style={{ marginTop: 14 }}>
              <div className={rowClass}>
                <Dropdown label="Loại xe *" options={[{ key: '', text: '-- Chọn loại xe --' }, ...vehicleTypeOptions]}
                  selectedKey={form.vehicleType}
                  onChange={(_, o) => this._setField('vehicleType', String(o?.key || ''))}
                  errorMessage={getFieldError(errors, 'vehicleType')} disabled={busy} />
                <TextField label="Số hành khách *" type="number" value={String(form.numberOfPassengers)}
                  onChange={(_, v) => this._setField('numberOfPassengers', parseInt(v || '1', 10))}
                  errorMessage={getFieldError(errors, 'numberOfPassengers')} disabled={busy} />
              </div>
            </div>
            <div style={{ marginTop: 14 }}>
              <TextField label="Mục đích chuyến đi *" multiline rows={3}
                value={form.purpose}
                onChange={(_, v) => this._setField('purpose', v || '')}
                errorMessage={getFieldError(errors, 'purpose')} disabled={busy} />
            </div>
            <div style={{ marginTop: 14 }}>
              <TextField label="Yêu cầu đặc biệt" multiline rows={2}
                value={form.specialRequirement}
                onChange={(_, v) => this._setField('specialRequirement', v || '')} disabled={busy} />
            </div>
          </div>

          {/* Người phê duyệt */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Người phê duyệt</div>

            {/* Cảnh báo nếu chưa có manager nào trong hệ thống — hiện phía trên picker */}
            {allManagers.length === 0 && (
              <div className={styles.noApproverWarn} style={{ marginTop: 0, marginBottom: 12 }}>
                <FontIcon iconName="Warning" style={{ fontSize: 18, flexShrink: 0 }} />
                <span>
                  Chưa có trưởng phòng nào được cấu hình. Vui lòng liên hệ Admin để thêm tài khoản Manager.
                </span>
              </div>
            )}

            {/* People Picker tìm kiếm manager */}
            <div className={styles.pickerWrap}>
              <div className={styles.pickerLabel}>Chọn người phê duyệt *</div>
              <div className={styles.pickerBox}>
                <NormalPeoplePicker
                  onResolveSuggestions={this._onResolveSuggestions}
                  onChange={this._onPickerChange}
                  itemLimit={1}
                  disabled={busy}
                  selectedItems={selectedManager ? [this._managerToPersona(selectedManager)] : []}
                  pickerSuggestionsProps={{
                    suggestionsHeaderText: 'Gợi ý người phê duyệt',
                    noResultsFoundText: 'Không tìm thấy. Thử nhập tên, email hoặc phòng ban.',
                    loadingText: 'Đang tìm...',
                  }}
                  inputProps={{
                    placeholder: 'Nhập tên, email hoặc phòng ban để tìm...',
                    'aria-label': 'Chọn người phê duyệt',
                  }}
                  styles={{
                    root: { width: '100%' },
                    text: { border: 'none', minWidth: 0 },
                  }}
                />
              </div>
            </div>

            {/* Card hiển thị người phê duyệt đang được chọn */}
            {selectedManager && (
              <div className={styles.approverCard}>
                <div className={styles.approverIcon}>
                  <FontIcon iconName="Contact" style={{ fontSize: 20, color: '#fff' }} />
                </div>
                <div>
                  <div className={styles.approverName}>{selectedManager.UserName}</div>
                  <div className={styles.approverMeta}>
                    {selectedManager.Department && (
                      <span style={{ marginRight: 12 }}>
                        <FontIcon iconName="Group" style={{ fontSize: 11, marginRight: 4 }} />
                        {selectedManager.Department}
                      </span>
                    )}
                    <span>
                      <FontIcon iconName="Mail" style={{ fontSize: 11, marginRight: 4 }} />
                      {selectedManager.UserEmail}
                    </span>
                  </div>
                  {!hasSameDeptManagers && this.props.user.department && (
                    <div style={{ fontSize: 11, color: '#F59E0B', marginTop: 4 }}>
                      <FontIcon iconName="Info" style={{ fontSize: 11, marginRight: 4 }} />
                      Không có manager cùng phòng ban — đang dùng manager từ phòng ban khác
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Actions */}
          <div className={styles.actions}>
            <PrimaryButton
              text={isSubmitting ? 'Đang gửi...' : 'Gửi yêu cầu'}
              disabled={busy || !form.selectedManagerEmail}
              onClick={() => this._submit().catch(console.error)}
              onRenderIcon={isSubmitting ? () => <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginRight: 6 } }} /> : undefined}
              styles={{ root: { borderRadius: 8 } }}
            />
            <DefaultButton
              text={isSaving ? 'Đang lưu...' : 'Lưu nháp'}
              disabled={busy}
              onClick={() => this._saveDraft().catch(console.error)}
              styles={{ root: { borderRadius: 8 } }}
            />
            <DefaultButton text="Hủy" disabled={busy}
              onClick={() => this.props.onNavigate('my-requests')}
              styles={{ root: { borderRadius: 8 } }} />
          </div>
        </div>
      </div>
    );
  }
}
