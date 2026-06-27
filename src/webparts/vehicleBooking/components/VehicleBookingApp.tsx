import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dropdown, IDropdownOption } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AppLayout } from '../../../components/layout/AppLayout';
import Dashboard from '../../../components/dashboard/Dashboard';
import MyRequests from '../../../components/requests/MyRequests';
import CreateRequestForm from '../../../components/requests/CreateRequestForm';
import RequestDetail from '../../../components/requests/RequestDetail';
import PendingMyApproval from '../../../components/approvals/PendingMyApproval';
import TransportAssignment from '../../../components/transport/TransportAssignment';
import VehicleScheduleCalendar from '../../../components/transport/VehicleScheduleCalendar';
import DriverTripList from '../../../components/driver/DriverTripList';
import CompletedTrips from '../../../components/driver/CompletedTrips';
import AdminSettings from '../../../components/admin/AdminSettings';
import UserRoleSetup from '../../../components/admin/UserRoleSetup';
import { UserRoleService } from '../../../services/userRoleService';
import { IUserRole } from '../../../models/UserRole';
import { IAppUser, AppScreen } from '../../../common/types/common';
import { ROLE } from '../../../common/constants/roles';
import { extractErrorMessage } from '../../../common/helpers/errorHelper';

// Bật DEV_MODE để hiện switcher giả lập user — tắt trước khi release
const DEV_MODE = true;

export interface IVehicleBookingAppProps {
  context: WebPartContext;
  apiBaseUrl?: string;
}

interface IVehicleBookingAppState {
  currentScreen: AppScreen;
  screenParams: unknown;
  user: IAppUser | null;
  realUser: IAppUser | null;
  allUsers: IUserRole[];
  isLoading: boolean;
  error: string | null;
  menuOpen: boolean;
}

export default class VehicleBookingApp extends React.Component<IVehicleBookingAppProps, IVehicleBookingAppState> {
  private userSvc: UserRoleService;

  constructor(props: IVehicleBookingAppProps) {
    super(props);
    this.state = {
      currentScreen: 'dashboard',
      screenParams: null,
      user: null,
      realUser: null,
      allUsers: [],
      isLoading: true,
      error: null,
      menuOpen: typeof window !== 'undefined' && window.innerWidth > 768,
    };
    this.userSvc = new UserRoleService(props.context);
  }

  public componentDidMount(): void {
    this._loadUser().catch(console.error);
  }

  private _getDeepLinkScreen(): { screen: AppScreen; params: unknown } | null {
    const params = new URLSearchParams(window.location.search);
    const requestId = params.get('requestId');
    if (requestId && !isNaN(Number(requestId))) {
      return { screen: 'request-detail', params: { id: Number(requestId) } };
    }
    return null;
  }

  private async _loadUser(): Promise<void> {
    try {
      const [user, allUsers] = await Promise.all([
        this.userSvc.getCurrentUserRole(this.props.context),
        DEV_MODE ? this.userSvc.getAllIncludingInactive() : Promise.resolve([]),
      ]);
      const deepLink = this._getDeepLinkScreen();
      this.setState({
        user, realUser: user, allUsers, isLoading: false,
        currentScreen: deepLink ? deepLink.screen : 'dashboard',
        screenParams:  deepLink ? deepLink.params  : null,
      });
    } catch (err) {
      this.setState({ error: extractErrorMessage(err), isLoading: false });
    }
  }

  private _switchUser = (email: string): void => {
    const { allUsers, realUser } = this.state;
    if (!email) {
      this.setState({ user: realUser, currentScreen: 'dashboard' });
      return;
    }
    const found = allUsers.find(u => u.UserEmail === email);
    if (!found) return;
    const spoofed: IAppUser = {
      userId:     found.UserEmail,
      userName:   found.UserName,
      userEmail:  found.UserEmail,
      role:       found.Role || ROLE.REQUESTER,
      department: found.Department,
      isActive:   found.IsActive,
    };
    this.setState({ user: spoofed, currentScreen: 'dashboard' });
  }

  private _navigate = (screen: AppScreen, params?: unknown): void => {
    this.setState({ currentScreen: screen, screenParams: params || null });
  }

  private _renderScreen(): React.ReactNode {
    const { currentScreen, screenParams, user } = this.state;
    const { context } = this.props;
    if (!user) return null;

    const params = screenParams as Record<string, unknown> | null;

    switch (currentScreen) {
      case 'dashboard':
        return <Dashboard context={context} user={user} onNavigate={this._navigate} />;
      case 'my-requests':
        return <MyRequests context={context} user={user} onNavigate={this._navigate} />;
      case 'create-request':
        return <CreateRequestForm context={context} user={user}
          editId={params?.editId as number | undefined}
          onNavigate={this._navigate} />;
      case 'request-detail':
        return <RequestDetail context={context} user={user}
          requestId={params?.id as number || 0} onNavigate={this._navigate} />;
      case 'pending-approval':
        return <PendingMyApproval context={context} user={user} onNavigate={this._navigate} />;
      case 'transport-assignment':
        return <TransportAssignment context={context} user={user} onNavigate={this._navigate} />;
      case 'vehicle-schedule':
        return <VehicleScheduleCalendar context={context} user={user} onNavigate={this._navigate} />;
      case 'driver-trips':
        return <DriverTripList context={context} user={user} onNavigate={this._navigate} />;
      case 'completed-trips':
        return <CompletedTrips context={context} user={user} onNavigate={this._navigate} />;
      case 'admin-settings':
        return <AdminSettings context={context} user={user} onNavigate={this._navigate} />;
      case 'user-role-setup':
        return <UserRoleSetup context={context} user={user} onNavigate={this._navigate} />;
      default:
        return <Dashboard context={context} user={user} onNavigate={this._navigate} />;
    }
  }

  private _renderDevSwitcher(): React.ReactElement | null {
    if (!DEV_MODE) return null;
    const { allUsers, user, realUser } = this.state;
    if (!allUsers.length) return null;

    const options: IDropdownOption[] = [
      { key: '', text: `— Tôi (${realUser ? realUser.userEmail : ''}) —` },
      ...allUsers.map(u => ({
        key: u.UserEmail,
        text: `${u.UserName} · ${u.Role} · ${u.Department || 'N/A'}`,
      })),
    ];

    const isSpoofing = user && realUser && user.userEmail !== realUser.userEmail;

    return (
      <div style={{
        background: isSpoofing ? '#7C3AED' : '#1F2937',
        padding: '6px 16px',
        display: 'flex',
        alignItems: 'center',
        gap: 12,
        flexShrink: 0,
        zIndex: 500,
      }}>
        <span style={{ color: '#FCD34D', fontSize: 11, fontWeight: 700, whiteSpace: 'nowrap' as const }}>
          🧪 DEV
        </span>
        <span style={{ color: '#9CA3AF', fontSize: 11, whiteSpace: 'nowrap' as const }}>Giả lập user:</span>
        <div style={{ flex: 1, maxWidth: 400 }}>
          <Dropdown
            options={options}
            selectedKey={isSpoofing ? user!.userEmail : ''}
            onChange={(_, o) => this._switchUser(String(o?.key || ''))}
            styles={{
              root: { margin: 0 },
              title: {
                background: 'rgba(255,255,255,0.1)', border: '1px solid rgba(255,255,255,0.2)',
                color: '#F9FAFB', fontSize: 12, height: 28, lineHeight: '28px',
              },
              caretDown: { color: '#F9FAFB' },
              dropdownItem: { fontSize: 12 },
              dropdownItemSelected: { fontSize: 12 },
            }}
          />
        </div>
        {isSpoofing && (
          <span style={{
            background: '#FCD34D', color: '#1F2937', fontSize: 10, fontWeight: 700,
            padding: '2px 8px', borderRadius: 10,
          }}>
            {user!.role.toUpperCase()}
          </span>
        )}
      </div>
    );
  }

  public render(): React.ReactElement {
    const { user, isLoading, error, currentScreen, menuOpen } = this.state;

    if (isLoading) {
      return (
        <div style={{ display: 'flex', height: '100vh', alignItems: 'center', justifyContent: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Đang khởi động hệ thống..." />
        </div>
      );
    }

    if (error) {
      return (
        <div style={{ padding: 24 }}>
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      );
    }

    if (!user) return null;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
        {this._renderDevSwitcher()}
        <div style={{ flex: 1, overflow: 'hidden' }}>
          <AppLayout
            user={user}
            currentScreen={currentScreen}
            menuOpen={menuOpen}
            onNavigate={this._navigate}
            onToggleMenu={() => this.setState(prev => ({ menuOpen: !prev.menuOpen }))}>
            {this._renderScreen()}
          </AppLayout>
        </div>
      </div>
    );
  }
}
