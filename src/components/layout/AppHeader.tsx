import * as React from 'react';
import { Persona, PersonaSize, IconButton, mergeStyleSets, FontIcon } from '@fluentui/react';
import { IAppUser } from '../../common/types/common';
import { ROLE_LABEL } from '../../common/constants/roles';

const styles = mergeStyleSets({
  header: {
    height: 56,
    background: '#0078D4',
    display: 'flex' as const,
    alignItems: 'center' as const,
    padding: '0 16px',
    justifyContent: 'space-between' as const,
    flexShrink: 0,
    boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
    boxSizing: 'border-box' as const,
    zIndex: 300,
    position: 'relative' as const,
  },
  brand: {
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 8,
    color: '#fff',
    fontSize: 15,
    fontWeight: 700,
    minWidth: 0,
  },
  brandText: {
    overflow: 'hidden' as const,
    textOverflow: 'ellipsis' as const,
    whiteSpace: 'nowrap' as const,
  },
  right: {
    display: 'flex' as const,
    alignItems: 'center' as const,
    gap: 8,
    flexShrink: 0,
  },
});

interface IAppHeaderProps {
  user: IAppUser;
  menuOpen: boolean;
  onToggleMenu: () => void;
}

export const AppHeader: React.FC<IAppHeaderProps> = ({ user, menuOpen, onToggleMenu }) => (
  <div className={styles.header}>
    <div className={styles.brand}>
      <IconButton
        iconProps={{ iconName: menuOpen ? 'ChromeClose' : 'GlobalNavButton' }}
        onClick={onToggleMenu}
        title={menuOpen ? 'Đóng menu' : 'Mở menu'}
        styles={{
          root: { color: '#fff', flexShrink: 0 },
          rootHovered: { background: 'rgba(255,255,255,0.15)', color: '#fff' },
          rootPressed: { background: 'rgba(255,255,255,0.25)', color: '#fff' },
          icon: { fontSize: menuOpen ? 12 : 16 },
        }}
      />
      <FontIcon iconName="Car" style={{ fontSize: 20, color: '#fff', flexShrink: 0 }} />
      <span className={styles.brandText}>Vehicle Booking</span>
    </div>
    <div className={styles.right}>
      <Persona
        text={user.userName}
        secondaryText={ROLE_LABEL[user.role] || user.role}
        size={PersonaSize.size32}
        styles={{
          root: { color: '#fff' },
          primaryText: { color: '#fff', fontSize: 13 },
          secondaryText: { color: 'rgba(255,255,255,0.75)', fontSize: 11 },
          details: { paddingRight: 0 },
        }}
      />
    </div>
  </div>
);
