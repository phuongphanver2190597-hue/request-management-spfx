import * as React from 'react';
import {
  Pivot, PivotItem, Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, TextField, Dropdown, IDropdownOption,
  Dialog, DialogType, DialogFooter, IconButton, mergeStyleSets, FontIcon, Text,
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAppUser, AppScreen } from '../../common/types/common';
import { VehicleMasterService } from '../../services/vehicleMasterService';
import { DriverMasterService } from '../../services/driverMasterService';
import { UserRoleService } from '../../services/userRoleService';
import { LocationMasterService } from '../../services/locationMasterService';
import { IVehicleMaster } from '../../models/VehicleMaster';
import { IDriverMaster } from '../../models/DriverMaster';
import { IUserRole } from '../../models/UserRole';
import { ILocationMaster } from '../../models/LocationMaster';
import { VEHICLE_TYPES, VEHICLE_STATUS, DRIVER_STATUS } from '../../common/constants/statuses';
import { ROLE, ROLE_LABEL } from '../../common/constants/roles';
import { extractErrorMessage } from '../../common/helpers/errorHelper';
import { PageHeader } from '../shared/PageHeader';
import { StatusBadge } from '../shared/StatusBadge';

const WHITE = '#FFFFFF';
const BORDER = '#E1E4E8';
const PRIMARY = '#0078D4';
const TEXT_MUTE = '#6A737D';

const styles = mergeStyleSets({
  tableWrap: { background: WHITE, borderRadius: 12, boxShadow: '0 1px 4px rgba(0,0,0,0.08)', overflow: 'hidden', marginTop: 14 },
  toolbar: { padding: '10px 14px', borderBottom: `1px solid ${BORDER}`, display: 'flex' as const, justifyContent: 'flex-end' as const },
  tableScroll: { overflowX: 'auto' as const },
  table: { width: '100%', borderCollapse: 'collapse' as const, minWidth: 480 },
  th: {
    padding: '10px 14px', textAlign: 'left' as const, fontSize: 11,
    fontWeight: 700, color: TEXT_MUTE, textTransform: 'uppercase' as const,
    background: '#FAFBFC', borderBottom: `1px solid ${BORDER}`, whiteSpace: 'nowrap' as const,
  },
  td: { padding: '10px 14px', fontSize: 13, borderBottom: `1px solid ${BORDER}` },
  empty: { padding: '32px 0', textAlign: 'center' as const, color: TEXT_MUTE },
  formRow2: { display: 'grid' as const, gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 },
  formRow1: { display: 'grid' as const, gridTemplateColumns: '1fr', gap: 14, marginBottom: 14 },
  mobileCard: {
    padding: '12px 14px', borderBottom: `1px solid ${BORDER}`,
    ':last-child': { borderBottom: 'none' },
  },
  mobileName: { fontWeight: 700, fontSize: 13, marginBottom: 3 },
  mobileLine: { fontSize: 12, color: TEXT_MUTE, marginBottom: 2 },
  mobileActions: { marginTop: 8, display: 'flex' as const, gap: 8, alignItems: 'center' as const },
});

const vehicleTypeOptions: IDropdownOption[] = VEHICLE_TYPES.map(v => ({ key: v, text: v }));
const vehicleStatusOptions: IDropdownOption[] = [
  VEHICLE_STATUS.AVAILABLE, VEHICLE_STATUS.IN_USE, VEHICLE_STATUS.MAINTENANCE, VEHICLE_STATUS.INACTIVE,
].map(v => ({ key: v, text: v }));
const driverStatusOptions: IDropdownOption[] = [
  DRIVER_STATUS.AVAILABLE, DRIVER_STATUS.ON_TRIP, DRIVER_STATUS.INACTIVE,
].map(v => ({ key: v, text: v }));
const roleOptions: IDropdownOption[] = [
  ROLE.REQUESTER, ROLE.MANAGER, ROLE.TRANSPORT_ADMIN, ROLE.DRIVER, ROLE.ADMIN,
].map(r => ({ key: r, text: ROLE_LABEL[r] || r }));

interface IAdminSettingsProps {
  context: WebPartContext;
  user: IAppUser;
  onNavigate: (screen: AppScreen, params?: unknown) => void;
}

interface IAdminSettingsState {
  vehicles: IVehicleMaster[];
  drivers: IDriverMaster[];
  users: IUserRole[];
  locations: ILocationMaster[];
  isLoading: boolean;
  error: string | null;
  successMsg: string | null;
  dialogType: 'vehicle' | 'driver' | 'user' | 'location' | null;
  isSaving: boolean;
  vForm: Partial<IVehicleMaster>;
  dForm: Partial<IDriverMaster>;
  uForm: Partial<IUserRole>;
  lForm: Partial<ILocationMaster>;
  isMobile: boolean;
}

export default class AdminSettings extends React.Component<IAdminSettingsProps, IAdminSettingsState> {
  private vehicleSvc: VehicleMasterService;
  private driverSvc: DriverMasterService;
  private userSvc: UserRoleService;
  private locationSvc: LocationMasterService;
  private _mql: MediaQueryList | null = null;

  constructor(props: IAdminSettingsProps) {
    super(props);
    this.state = {
      vehicles: [], drivers: [], users: [], locations: [],
      isLoading: false, error: null, successMsg: null,
      dialogType: null, isSaving: false,
      vForm: {}, dForm: {}, uForm: {}, lForm: {},
      isMobile: typeof window !== 'undefined' && window.innerWidth <= 600,
    };
    this.vehicleSvc = new VehicleMasterService(props.context);
    this.driverSvc  = new DriverMasterService(props.context);
    this.userSvc    = new UserRoleService(props.context);
    this.locationSvc= new LocationMasterService(props.context);
  }

  public componentDidMount(): void {
    this._loadAll().catch(console.error);
    if (typeof window !== 'undefined') {
      this._mql = window.matchMedia('(max-width: 600px)');
      this._mql.addListener(this._onMql);
    }
  }

  public componentWillUnmount(): void {
    if (this._mql) this._mql.removeListener(this._onMql);
  }

  private _onMql = (e: MediaQueryListEvent): void => { this.setState({ isMobile: e.matches }); }

  private async _loadAll(): Promise<void> {
    this.setState({ isLoading: true, error: null });
    try {
      const [vehicles, drivers, users, locations] = await Promise.all([
        this.vehicleSvc.getAll(),
        this.driverSvc.getAll(),
        this.userSvc.getAll(),
        this.locationSvc.getAllIncludingInactive(),
      ]);
      this.setState({ vehicles, drivers, users, locations, isLoading: false });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private async _saveVehicle(): Promise<void> {
    const { vForm } = this.state;
    if (!vForm.PlateNumber || !vForm.VehicleType) return;
    this.setState({ isSaving: true });
    try {
      await this.vehicleSvc.create({
        Title: vForm.PlateNumber,
        VehicleCode: `VH-${Date.now()}`,
        PlateNumber: vForm.PlateNumber,
        VehicleType: vForm.VehicleType,
        Brand: vForm.Brand || '',
        Model: vForm.Model || '',
        Capacity: vForm.Capacity || 4,
        Status: vForm.Status || VEHICLE_STATUS.AVAILABLE,
        CurrentOdometer: vForm.CurrentOdometer || 0,
        Note: vForm.Note || '',
        IsDeleted: false,
      });
      this.setState({ dialogType: null, vForm: {}, successMsg: 'Đã thêm xe thành công', isSaving: false });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  private async _saveDriver(): Promise<void> {
    const { dForm } = this.state;
    if (!dForm.DriverName || !dForm.DriverPhone) return;
    this.setState({ isSaving: true });
    try {
      await this.driverSvc.create({
        Title: dForm.DriverName,
        DriverCode: `DRV-${Date.now()}`,
        DriverName: dForm.DriverName,
        DriverPhone: dForm.DriverPhone,
        DriverEmail: dForm.DriverEmail || '',
        Status: dForm.Status || DRIVER_STATUS.AVAILABLE,
        LicenseNumber: dForm.LicenseNumber || '',
        Note: dForm.Note || '',
        IsDeleted: false,
      });
      this.setState({ dialogType: null, dForm: {}, successMsg: 'Đã thêm tài xế thành công', isSaving: false });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  private async _saveUser(): Promise<void> {
    const { uForm } = this.state;
    if (!uForm.UserEmail || !uForm.Role) return;
    this.setState({ isSaving: true });
    try {
      await this.userSvc.create({
        Title: uForm.UserName || uForm.UserEmail,
        UserId: uForm.UserEmail,
        UserName: uForm.UserName || '',
        UserEmail: uForm.UserEmail,
        Role: uForm.Role,
        Department: uForm.Department || '',
        IsActive: true,
      });
      this.setState({ dialogType: null, uForm: {}, successMsg: 'Đã thêm người dùng thành công', isSaving: false });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  private async _saveLocation(): Promise<void> {
    const { lForm } = this.state;
    if (!lForm.LocationName) return;
    this.setState({ isSaving: true });
    try {
      await this.locationSvc.create({
        Title: lForm.LocationName,
        LocationCode: `LOC-${Date.now()}`,
        LocationName: lForm.LocationName,
        Address: lForm.Address || '',
        Province: lForm.Province || '',
        Status: 'ACTIVE',
        IsDeleted: false,
      });
      this.setState({ dialogType: null, lForm: {}, successMsg: 'Đã thêm địa điểm thành công', isSaving: false });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isSaving: false });
    }
  }

  private async _deactivateUser(id: number): Promise<void> {
    try {
      await this.userSvc.deactivate(id);
      this.setState({ successMsg: 'Đã vô hiệu hóa người dùng' });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    }
  }

  private async _deleteVehicle(id: number): Promise<void> {
    try {
      await this.vehicleSvc.softDelete(id);
      this.setState({ successMsg: 'Đã xóa xe' });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    }
  }

  private async _deleteDriver(id: number): Promise<void> {
    try {
      await this.driverSvc.softDelete(id);
      this.setState({ successMsg: 'Đã xóa tài xế' });
      await this._loadAll();
    } catch (err) {
      this.setState({ error: extractErrorMessage(err) });
    }
  }

  public render(): React.ReactElement {
    const { vehicles, drivers, users, locations, isLoading, error, successMsg, dialogType, isSaving, vForm, dForm, uForm, lForm, isMobile } = this.state;
    const formRow = isMobile ? styles.formRow1 : styles.formRow2;

    return (
      <div>
        <PageHeader icon="Settings" title="Cài đặt hệ thống" subtitle="Quản lý danh mục xe, tài xế, người dùng, địa điểm" />

        {error && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 8, marginBottom: 12 } }} onDismiss={() => this.setState({ error: null })}>{error}</MessageBar>}
        {successMsg && <MessageBar messageBarType={MessageBarType.success} styles={{ root: { borderRadius: 8, marginBottom: 12 } }} onDismiss={() => this.setState({ successMsg: null })}>{successMsg}</MessageBar>}

        {isLoading ? (
          <Spinner size={SpinnerSize.large} label="Đang tải..." styles={{ root: { padding: '60px 0' } }} />
        ) : (
          <Pivot>
            {/* ── Vehicles ── */}
            <PivotItem headerText={`Xe (${vehicles.length})`} itemIcon="Car">
              <div className={styles.tableWrap}>
                <div className={styles.toolbar}>
                  <PrimaryButton iconProps={{ iconName: 'Add' }} text="Thêm xe"
                    onClick={() => this.setState({ dialogType: 'vehicle', vForm: {} })}
                    styles={{ root: { borderRadius: 8 } }} />
                </div>
                {isMobile ? (
                  <div>
                    {vehicles.map(v => (
                      <div key={v.ID} className={styles.mobileCard}>
                        <div className={styles.mobileName}>{v.PlateNumber}</div>
                        <div className={styles.mobileLine}>{v.VehicleType} · {v.Brand} {v.Model}</div>
                        <div className={styles.mobileLine}>{v.Capacity} chỗ · {v.CurrentOdometer?.toLocaleString('vi-VN')} km</div>
                        <div className={styles.mobileActions}>
                          <StatusBadge status={v.Status} />
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Xóa"
                            onClick={() => this._deleteVehicle(v.ID).catch(console.error)}
                            styles={{ root: { color: '#C53030' } }} />
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.tableScroll}>
                    <table className={styles.table}>
                      <thead><tr>
                        <th className={styles.th}>Biển số</th>
                        <th className={styles.th}>Loại xe</th>
                        <th className={styles.th}>Hãng / Mẫu</th>
                        <th className={styles.th}>Sức chứa</th>
                        <th className={styles.th}>ODO hiện tại</th>
                        <th className={styles.th}>Trạng thái</th>
                        <th className={styles.th}></th>
                      </tr></thead>
                      <tbody>
                        {vehicles.map(v => (
                          <tr key={v.ID}>
                            <td className={styles.td} style={{ fontWeight: 600 }}>{v.PlateNumber}</td>
                            <td className={styles.td}>{v.VehicleType}</td>
                            <td className={styles.td}>{v.Brand} {v.Model}</td>
                            <td className={styles.td}>{v.Capacity} chỗ</td>
                            <td className={styles.td}>{v.CurrentOdometer?.toLocaleString('vi-VN')} km</td>
                            <td className={styles.td}><StatusBadge status={v.Status} /></td>
                            <td className={styles.td}>
                              <IconButton iconProps={{ iconName: 'Delete' }} title="Xóa"
                                onClick={() => this._deleteVehicle(v.ID).catch(console.error)}
                                styles={{ root: { color: '#C53030' } }} />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </PivotItem>

            {/* ── Drivers ── */}
            <PivotItem headerText={`Tài xế (${drivers.length})`} itemIcon="Contact">
              <div className={styles.tableWrap}>
                <div className={styles.toolbar}>
                  <PrimaryButton iconProps={{ iconName: 'Add' }} text="Thêm tài xế"
                    onClick={() => this.setState({ dialogType: 'driver', dForm: {} })}
                    styles={{ root: { borderRadius: 8 } }} />
                </div>
                {isMobile ? (
                  <div>
                    {drivers.map(d => (
                      <div key={d.ID} className={styles.mobileCard}>
                        <div className={styles.mobileName}>{d.DriverName}</div>
                        <div className={styles.mobileLine}>{d.DriverPhone} · {d.DriverEmail}</div>
                        <div className={styles.mobileLine}>Bằng: {d.LicenseNumber || '-'}</div>
                        <div className={styles.mobileActions}>
                          <StatusBadge status={d.Status} />
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Xóa"
                            onClick={() => this._deleteDriver(d.ID).catch(console.error)}
                            styles={{ root: { color: '#C53030' } }} />
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.tableScroll}>
                    <table className={styles.table}>
                      <thead><tr>
                        <th className={styles.th}>Họ tên</th>
                        <th className={styles.th}>Điện thoại</th>
                        <th className={styles.th}>Email</th>
                        <th className={styles.th}>Bằng lái</th>
                        <th className={styles.th}>Trạng thái</th>
                        <th className={styles.th}></th>
                      </tr></thead>
                      <tbody>
                        {drivers.map(d => (
                          <tr key={d.ID}>
                            <td className={styles.td} style={{ fontWeight: 600 }}>{d.DriverName}</td>
                            <td className={styles.td}>{d.DriverPhone}</td>
                            <td className={styles.td}>{d.DriverEmail}</td>
                            <td className={styles.td}>{d.LicenseNumber}</td>
                            <td className={styles.td}><StatusBadge status={d.Status} /></td>
                            <td className={styles.td}>
                              <IconButton iconProps={{ iconName: 'Delete' }} title="Xóa"
                                onClick={() => this._deleteDriver(d.ID).catch(console.error)}
                                styles={{ root: { color: '#C53030' } }} />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </PivotItem>

            {/* ── Users ── */}
            <PivotItem headerText={`Người dùng (${users.length})`} itemIcon="People">
              <div className={styles.tableWrap}>
                <div className={styles.toolbar}>
                  <PrimaryButton iconProps={{ iconName: 'Add' }} text="Thêm người dùng"
                    onClick={() => this.setState({ dialogType: 'user', uForm: {} })}
                    styles={{ root: { borderRadius: 8 } }} />
                </div>
                {isMobile ? (
                  <div>
                    {users.map(u => (
                      <div key={u.ID} className={styles.mobileCard}>
                        <div className={styles.mobileName}>{u.UserName}</div>
                        <div className={styles.mobileLine}>{u.UserEmail}</div>
                        <div className={styles.mobileLine}>{ROLE_LABEL[u.Role] || u.Role} · {u.Department}</div>
                        <div className={styles.mobileActions}>
                          <span style={{ fontSize: 12, color: u.IsActive ? '#20BF6B' : '#C53030', fontWeight: 600 }}>
                            {u.IsActive ? 'Hoạt động' : 'Không HĐ'}
                          </span>
                          {u.IsActive && (
                            <IconButton iconProps={{ iconName: 'UserRemove' }} title="Vô hiệu hóa"
                              onClick={() => this._deactivateUser(u.ID).catch(console.error)}
                              styles={{ root: { color: '#C53030' } }} />
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.tableScroll}>
                    <table className={styles.table}>
                      <thead><tr>
                        <th className={styles.th}>Họ tên</th>
                        <th className={styles.th}>Email</th>
                        <th className={styles.th}>Vai trò</th>
                        <th className={styles.th}>Phòng ban</th>
                        <th className={styles.th}>Hoạt động</th>
                        <th className={styles.th}></th>
                      </tr></thead>
                      <tbody>
                        {users.map(u => (
                          <tr key={u.ID}>
                            <td className={styles.td} style={{ fontWeight: 600 }}>{u.UserName}</td>
                            <td className={styles.td}>{u.UserEmail}</td>
                            <td className={styles.td}>{ROLE_LABEL[u.Role] || u.Role}</td>
                            <td className={styles.td}>{u.Department}</td>
                            <td className={styles.td}>{u.IsActive ? <span style={{ color: '#20BF6B' }}>✓</span> : <span style={{ color: '#C53030' }}>✗</span>}</td>
                            <td className={styles.td}>
                              {u.IsActive && (
                                <IconButton iconProps={{ iconName: 'UserRemove' }} title="Vô hiệu hóa"
                                  onClick={() => this._deactivateUser(u.ID).catch(console.error)}
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
            </PivotItem>

            {/* ── Locations ── */}
            <PivotItem headerText={`Địa điểm (${locations.length})`} itemIcon="MapPin">
              <div className={styles.tableWrap}>
                <div className={styles.toolbar}>
                  <PrimaryButton iconProps={{ iconName: 'Add' }} text="Thêm địa điểm"
                    onClick={() => this.setState({ dialogType: 'location', lForm: {} })}
                    styles={{ root: { borderRadius: 8 } }} />
                </div>
                {isMobile ? (
                  <div>
                    {locations.map(l => (
                      <div key={l.ID} className={styles.mobileCard}>
                        <div className={styles.mobileName}>{l.LocationName}</div>
                        <div className={styles.mobileLine}>{l.Address}</div>
                        <div className={styles.mobileActions}>
                          <span style={{ fontSize: 12, color: TEXT_MUTE }}>{l.Province}</span>
                          <StatusBadge status={l.Status} />
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.tableScroll}>
                    <table className={styles.table}>
                      <thead><tr>
                        <th className={styles.th}>Tên địa điểm</th>
                        <th className={styles.th}>Địa chỉ</th>
                        <th className={styles.th}>Tỉnh/TP</th>
                        <th className={styles.th}>Trạng thái</th>
                      </tr></thead>
                      <tbody>
                        {locations.map(l => (
                          <tr key={l.ID}>
                            <td className={styles.td} style={{ fontWeight: 600 }}>{l.LocationName}</td>
                            <td className={styles.td}>{l.Address}</td>
                            <td className={styles.td}>{l.Province}</td>
                            <td className={styles.td}><StatusBadge status={l.Status} /></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </PivotItem>
          </Pivot>
        )}

        {/* Add Vehicle Dialog */}
        <Dialog hidden={dialogType !== 'vehicle'} dialogContentProps={{ type: DialogType.normal, title: 'Thêm xe mới' }}
          onDismiss={() => this.setState({ dialogType: null })}>
          <div className={formRow}>
            <TextField label="Biển số *" value={vForm.PlateNumber || ''} onChange={(_, v) => this.setState({ vForm: { ...vForm, PlateNumber: v || '' } })} />
            <Dropdown label="Loại xe *" options={vehicleTypeOptions} selectedKey={vForm.VehicleType || null}
              onChange={(_, o) => this.setState({ vForm: { ...vForm, VehicleType: String(o?.key || '') } })} />
            <TextField label="Hãng xe" value={vForm.Brand || ''} onChange={(_, v) => this.setState({ vForm: { ...vForm, Brand: v || '' } })} />
            <TextField label="Mẫu xe" value={vForm.Model || ''} onChange={(_, v) => this.setState({ vForm: { ...vForm, Model: v || '' } })} />
            <TextField label="Sức chứa" type="number" value={String(vForm.Capacity || 4)} onChange={(_, v) => this.setState({ vForm: { ...vForm, Capacity: parseInt(v || '4', 10) } })} />
            <TextField label="ODO hiện tại" type="number" value={String(vForm.CurrentOdometer || 0)} onChange={(_, v) => this.setState({ vForm: { ...vForm, CurrentOdometer: parseInt(v || '0', 10) } })} />
          </div>
          <DialogFooter>
            <PrimaryButton text={isSaving ? 'Đang lưu...' : 'Thêm'} disabled={isSaving} onClick={() => this._saveVehicle().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ dialogType: null })} />
          </DialogFooter>
        </Dialog>

        {/* Add Driver Dialog */}
        <Dialog hidden={dialogType !== 'driver'} dialogContentProps={{ type: DialogType.normal, title: 'Thêm tài xế mới' }}
          onDismiss={() => this.setState({ dialogType: null })}>
          <div className={formRow}>
            <TextField label="Họ tên *" value={dForm.DriverName || ''} onChange={(_, v) => this.setState({ dForm: { ...dForm, DriverName: v || '' } })} />
            <TextField label="Điện thoại *" value={dForm.DriverPhone || ''} onChange={(_, v) => this.setState({ dForm: { ...dForm, DriverPhone: v || '' } })} />
            <TextField label="Email" value={dForm.DriverEmail || ''} onChange={(_, v) => this.setState({ dForm: { ...dForm, DriverEmail: v || '' } })} />
            <TextField label="Số bằng lái" value={dForm.LicenseNumber || ''} onChange={(_, v) => this.setState({ dForm: { ...dForm, LicenseNumber: v || '' } })} />
          </div>
          <DialogFooter>
            <PrimaryButton text={isSaving ? 'Đang lưu...' : 'Thêm'} disabled={isSaving} onClick={() => this._saveDriver().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ dialogType: null })} />
          </DialogFooter>
        </Dialog>

        {/* Add User Dialog */}
        <Dialog hidden={dialogType !== 'user'} dialogContentProps={{ type: DialogType.normal, title: 'Thêm người dùng' }}
          onDismiss={() => this.setState({ dialogType: null })}>
          <div className={formRow}>
            <TextField label="Họ tên" value={uForm.UserName || ''} onChange={(_, v) => this.setState({ uForm: { ...uForm, UserName: v || '' } })} />
            <TextField label="Email *" value={uForm.UserEmail || ''} onChange={(_, v) => this.setState({ uForm: { ...uForm, UserEmail: v || '' } })} />
            <Dropdown label="Vai trò *" options={roleOptions} selectedKey={uForm.Role || null}
              onChange={(_, o) => this.setState({ uForm: { ...uForm, Role: String(o?.key || '') } })} />
            <TextField label="Phòng ban" value={uForm.Department || ''} onChange={(_, v) => this.setState({ uForm: { ...uForm, Department: v || '' } })} />
          </div>
          <DialogFooter>
            <PrimaryButton text={isSaving ? 'Đang lưu...' : 'Thêm'} disabled={isSaving} onClick={() => this._saveUser().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ dialogType: null })} />
          </DialogFooter>
        </Dialog>

        {/* Add Location Dialog */}
        <Dialog hidden={dialogType !== 'location'} dialogContentProps={{ type: DialogType.normal, title: 'Thêm địa điểm' }}
          onDismiss={() => this.setState({ dialogType: null })}>
          <div className={formRow}>
            <TextField label="Tên địa điểm *" value={lForm.LocationName || ''} onChange={(_, v) => this.setState({ lForm: { ...lForm, LocationName: v || '' } })} />
            <TextField label="Tỉnh/TP" value={lForm.Province || ''} onChange={(_, v) => this.setState({ lForm: { ...lForm, Province: v || '' } })} />
          </div>
          <TextField label="Địa chỉ đầy đủ" multiline rows={2} value={lForm.Address || ''} onChange={(_, v) => this.setState({ lForm: { ...lForm, Address: v || '' } })} />
          <DialogFooter>
            <PrimaryButton text={isSaving ? 'Đang lưu...' : 'Thêm'} disabled={isSaving} onClick={() => this._saveLocation().catch(console.error)} />
            <DefaultButton text="Hủy" onClick={() => this.setState({ dialogType: null })} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
