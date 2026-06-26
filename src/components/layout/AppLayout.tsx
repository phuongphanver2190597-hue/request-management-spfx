import * as React from 'react';
import { mergeStyleSets } from '@fluentui/react';
import { AppHeader } from './AppHeader';
import { AppMenu } from './AppMenu';
import { IAppUser, AppScreen } from '../../common/types/common';

const styles = mergeStyleSets({
  shell: {
    display: 'flex' as const,
    flexDirection: 'column' as const,
    height: '100vh',
    overflow: 'hidden',
    fontFamily: `'Segoe UI', -apple-system, sans-serif`,
    background: '#F0F2F5',
    position: 'relative' as const,
  },
  header: {
    position: 'relative' as const,
    zIndex: 300,
    flexShrink: 0,
  },
  body: {
    display: 'flex' as const,
    flex: 1,
    overflow: 'hidden',
    position: 'relative' as const,
  },
  /* Mobile backdrop — absolute within body */
  backdrop: {
    position: 'absolute' as const,
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    background: 'rgba(0,0,0,0.45)',
    zIndex: 100,
  },
  /* Mobile drawer — absolute within body, left edge */
  drawerMobile: {
    position: 'absolute' as const,
    top: 0,
    left: 0,
    bottom: 0,
    width: 260,
    zIndex: 101,
    background: '#fff',
    boxShadow: '4px 0 20px rgba(0,0,0,0.2)',
    overflowY: 'auto' as const,
  },
  /* Desktop sidebar — inline flex item */
  sidebarDesktop: {
    flexShrink: 0,
  },
  main: {
    flex: 1,
    overflowY: 'auto' as const,
    padding: 24,
  },
  mainMobile: {
    flex: 1,
    overflowY: 'auto' as const,
    padding: 12,
  },
});

interface IAppLayoutProps {
  user: IAppUser;
  currentScreen: AppScreen;
  menuOpen: boolean;
  onNavigate: (screen: AppScreen) => void;
  onToggleMenu: () => void;
  children: React.ReactNode;
}

interface IAppLayoutState {
  isMobile: boolean;
}

export class AppLayout extends React.Component<IAppLayoutProps, IAppLayoutState> {
  private _mql: MediaQueryList | null = null;

  constructor(props: IAppLayoutProps) {
    super(props);
    this.state = { isMobile: typeof window !== 'undefined' && window.innerWidth <= 768 };
  }

  public componentDidMount(): void {
    if (typeof window !== 'undefined') {
      this._mql = window.matchMedia('(max-width: 768px)');
      this._mql.addListener(this._onMql);
      this.setState({ isMobile: this._mql.matches });
    }
  }

  public componentWillUnmount(): void {
    if (this._mql) this._mql.removeListener(this._onMql);
  }

  private _onMql = (e: MediaQueryListEvent): void => {
    this.setState({ isMobile: e.matches });
  }

  public render(): React.ReactElement {
    const { user, currentScreen, menuOpen, onNavigate, onToggleMenu, children } = this.props;
    const { isMobile } = this.state;

    const menu = (
      <AppMenu
        currentScreen={currentScreen}
        role={user.role}
        isMobile={isMobile}
        onNavigate={(s) => {
          onNavigate(s);
          if (isMobile) onToggleMenu();
        }}
      />
    );

    return (
      <div className={styles.shell}>
        {/* Header luôn nằm trên cùng */}
        <div className={styles.header}>
          <AppHeader user={user} menuOpen={menuOpen} onToggleMenu={onToggleMenu} />
        </div>

        <div className={styles.body}>
          {isMobile ? (
            <>
              {/* Main content luôn render đầy đủ phía sau */}
              <main className={styles.mainMobile}>{children}</main>

              {/* Drawer + backdrop nằm trên main */}
              {menuOpen && (
                <>
                  <div className={styles.backdrop} onClick={onToggleMenu} />
                  <div className={styles.drawerMobile}>{menu}</div>
                </>
              )}
            </>
          ) : (
            <>
              {menuOpen && <div className={styles.sidebarDesktop}>{menu}</div>}
              <main className={styles.main}>{children}</main>
            </>
          )}
        </div>
      </div>
    );
  }
}
