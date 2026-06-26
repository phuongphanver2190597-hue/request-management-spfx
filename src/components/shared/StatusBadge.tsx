import * as React from 'react';
import { mergeStyleSets } from '@fluentui/react';
import { STATUS_COLOR, STATUS_LABEL } from '../../common/constants/statuses';

const styles = mergeStyleSets({
  badge: {
    display: 'inline-flex' as const,
    alignItems: 'center' as const,
    gap: 5,
    padding: '3px 10px',
    borderRadius: 20,
    fontSize: 11,
    fontWeight: 600,
    whiteSpace: 'nowrap' as const,
  },
  dot: {
    width: 6,
    height: 6,
    borderRadius: '50%',
    display: 'inline-block' as const,
    flexShrink: 0,
  },
});

interface IStatusBadgeProps {
  status: string;
}

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => {
  const s = STATUS_COLOR[status] || { bg: '#F3F4F6', color: '#6B7280', dot: '#9CA3AF' };
  return (
    <span className={styles.badge} style={{ background: s.bg, color: s.color }}>
      <span className={styles.dot} style={{ background: s.dot }} />
      {STATUS_LABEL[status] || status}
    </span>
  );
};
