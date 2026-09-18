import React from 'react';
import { getSemaforoStyles } from '../config/thresholds';
import { formatPercent } from '../utils/formatters';

interface StatusBadgeProps {
  percentage: number;
  showDot?: boolean;
  size?: 'sm' | 'md';
}

export const StatusBadge: React.FC<StatusBadgeProps> = ({
  percentage,
  showDot = true,
  size = 'md',
}) => {
  const styles = getSemaforoStyles(percentage);

  const sizeClasses =
    size === 'sm'
      ? 'px-2 py-0.5 text-xs font-semibold'
      : 'px-2.5 py-1 text-xs font-semibold';

  return (
    <span
      className={`inline-flex items-center gap-1.5 rounded-full border ${styles.bgColor} ${styles.textColor} ${styles.borderColor} ${sizeClasses} tabular-nums`}
      title={styles.label}
    >
      {showDot && (
        <span
          className={`h-2 w-2 rounded-full ${styles.dotColor} shrink-0`}
          aria-hidden="true"
        />
      )}
      <span>{formatPercent(percentage)}</span>
    </span>
  );
};
