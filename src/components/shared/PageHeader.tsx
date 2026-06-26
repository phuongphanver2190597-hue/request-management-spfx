import * as React from 'react';
import { mergeStyleSets, FontIcon } from '@fluentui/react';

const styles = mergeStyleSets({
  wrap: {
    marginBottom: 20,
    display: 'flex' as const,
    alignItems: 'center' as const,
    justifyContent: 'space-between' as const,
  },
  left: { display: 'flex' as const, alignItems: 'center' as const, gap: 10 },
  icon: { fontSize: 22, color: '#0078D4' },
  title: { fontSize: 20, fontWeight: 700, color: '#1B1F23', lineHeight: '1.2' },
  sub: { fontSize: 13, color: '#6A737D', marginTop: 2 },
});

interface IPageHeaderProps {
  icon: string;
  title: string;
  subtitle?: string;
  actions?: React.ReactNode;
}

export const PageHeader: React.FC<IPageHeaderProps> = ({ icon, title, subtitle, actions }) => (
  <div className={styles.wrap}>
    <div className={styles.left}>
      <FontIcon iconName={icon} className={styles.icon} />
      <div>
        <div className={styles.title}>{title}</div>
        {subtitle && <div className={styles.sub}>{subtitle}</div>}
      </div>
    </div>
    {actions && <div>{actions}</div>}
  </div>
);
