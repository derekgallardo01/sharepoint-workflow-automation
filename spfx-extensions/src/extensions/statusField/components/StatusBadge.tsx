import * as React from 'react';
import styles from './StatusBadge.module.scss';

export interface IStatusBadgeProps {
  status: string;
}

type StatusVariant = 'notStarted' | 'inProgress' | 'underReview' | 'approved' | 'rejected' | 'default';

interface IStatusConfig {
  label: string;
  variant: StatusVariant;
  icon: string;
}

const STATUS_MAP: Record<string, IStatusConfig> = {
  'Not Started': {
    label: 'Not Started',
    variant: 'notStarted',
    icon: '\u25CB' // Circle outline
  },
  'In Progress': {
    label: 'In Progress',
    variant: 'inProgress',
    icon: '\u25D4' // Half circle
  },
  'Under Review': {
    label: 'Under Review',
    variant: 'underReview',
    icon: '\u25D0' // Three-quarter circle
  },
  'Approved': {
    label: 'Approved',
    variant: 'approved',
    icon: '\u2713' // Checkmark
  },
  'Rejected': {
    label: 'Rejected',
    variant: 'rejected',
    icon: '\u2717' // Cross
  }
};

const DEFAULT_CONFIG: IStatusConfig = {
  label: '',
  variant: 'default',
  icon: '\u25CF' // Filled circle
};

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => {
  if (!status) {
    return <span className={styles.statusBadge} />;
  }

  const config = STATUS_MAP[status] || { ...DEFAULT_CONFIG, label: status };

  return (
    <span
      className={`${styles.statusBadge} ${styles[config.variant]}`}
      title={config.label}
      role="status"
      aria-label={`Status: ${config.label}`}
    >
      <span className={styles.icon} aria-hidden="true">
        {config.icon}
      </span>
      <span className={styles.label}>
        {config.label}
      </span>
    </span>
  );
};
