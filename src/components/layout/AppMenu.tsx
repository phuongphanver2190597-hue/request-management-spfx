import * as React from 'react';
import { mergeStyleSets, FontIcon } from '@fluentui/react';
import { AppScreen } from '../../common/types/common';
import { MENU_BY_ROLE } from '../../common/constants/roles';

const PRIMARY = '#0078D4';
const BG_ACTIVE = '#E6F2FB';

const styles = mergeStyleSets({
  nav: {
    width: 220,
    background: '#fff',
    borderRight: '1px solid #E1E4E8',
    overflowY: 'auto' as const,
    height: '100%',
    paddingTop: 8,
    boxSizing: 'border-box' as const,
  },
  navMobile: {
    width: '100%',
    background: '#fff',
    overflowY: 'auto' as const,
    height: '100%',
    paddingTop: 8,
    boxSizing: 'border-box' as const,
  },
  item: {
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 10,
    padding: '11px 16px',
    cursor: 'pointer',
    fontSize: 14,
    color: '#24292E',
    borderRadius: 6,
    margin: '1px 8px',
    transition: 'background 0.1s',
    ':hover': { background: '#F6F8FA' },
  },
  itemActive: {
    background: BG_ACTIVE,
    color: PRIMARY,
    fontWeight: 600,
  },
  section: {
    fontSize: 10,
    fontWeight: 700,
    color: '#6A737D',
    letterSpacing: '0.6px',
    textTransform: 'uppercase' as const,
    padding: '14px 16px 4px',
  },
  divider: {
    height: 1,
    background: '#E1E4E8',
    margin: '4px 12px',
  },
});

interface MenuItem {
  key: AppScreen;
  label: string;
  icon: string;
  section: string;
}

const ALL_MENU_ITEMS: MenuItem[] = [
  { key: 'dashboard',            label: 'Tổng quan',         icon: 'ViewDashboard', section: 'Chung' },
  { key: 'my-requests',          label: 'Yêu cầu của tôi',   icon: 'DocumentSet',   section: 'Yêu cầu' },
  { key: 'create-request',       label: 'Tạo yêu cầu mới',   icon: 'Add',           section: 'Yêu cầu' },
  { key: 'pending-approval',     label: 'Chờ duyệt',         icon: 'Clock',         section: 'Phê duyệt' },
  { key: 'transport-assignment', label: 'Phân công xe',       icon: 'Car',           section: 'Vận tải' },
  { key: 'vehicle-schedule',     label: 'Lịch xe',           icon: 'Calendar',      section: 'Vận tải' },
  { key: 'driver-trips',         label: 'Chuyến của tôi',    icon: 'Road',          section: 'Tài xế' },
  { key: 'completed-trips',      label: 'Chuyến hoàn thành', icon: 'CheckMark',     section: 'Báo cáo' },
  { key: 'admin-settings',       label: 'Cài đặt hệ thống',  icon: 'Settings',      section: 'Quản trị' },
  { key: 'user-role-setup',      label: 'Phân quyền người dùng', icon: 'PeopleAdd', section: 'Quản trị' },
];

interface IAppMenuProps {
  currentScreen: AppScreen;
  role: string;
  onNavigate: (screen: AppScreen) => void;
  isMobile?: boolean;
}

export const AppMenu: React.FC<IAppMenuProps> = ({ currentScreen, role, onNavigate, isMobile }) => {
  const allowed = MENU_BY_ROLE[role] || ['dashboard'];
  const visible = ALL_MENU_ITEMS.filter(m => (allowed as string[]).indexOf(m.key) !== -1);

  const sections: string[] = [];
  visible.forEach(m => { if (sections.indexOf(m.section) === -1) sections.push(m.section); });

  return (
    <nav className={isMobile ? styles.navMobile : styles.nav}>
      {sections.map((sec, si) => (
        <div key={sec}>
          {si > 0 && <div className={styles.divider} />}
          <div className={styles.section}>{sec}</div>
          {visible.filter(m => m.section === sec).map(m => {
            const isActive = currentScreen === m.key;
            return (
              <div
                key={m.key}
                className={`${styles.item}${isActive ? ' ' + styles.itemActive : ''}`}
                onClick={() => onNavigate(m.key)}
              >
                <FontIcon
                  iconName={m.icon}
                  style={{ fontSize: 16, color: isActive ? PRIMARY : '#6A737D', flexShrink: 0 }}
                />
                {m.label}
              </div>
            );
          })}
        </div>
      ))}
    </nav>
  );
};
