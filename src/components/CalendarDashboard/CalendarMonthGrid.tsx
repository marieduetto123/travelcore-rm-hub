import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { CalendarDayCell } from './CalendarDayCell';
import { CalendarDay, MonthData } from './types';
import { calendarTokens } from './tokens';
import { HeatmapConfig } from './HeatmapModal';

const DAYS_OF_WEEK = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

const useStyles = makeStyles((theme) => ({
  // ── Normal (2-month side-by-side) ──
  root: {
    width: '100%',
    display: 'flex',
    padding: theme.spacing(2, 2.5),
    gap: 60,
    boxSizing: 'border-box',
  },
  monthColumn: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
  monthHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.75),
    height: 30,
  },
  monthLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 400,
    color: theme.palette.text.secondary,
    lineHeight: 'normal',
  },
  lockedLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.75),
  },
  lockIcon: {
    fontSize: '12px !important',
    color: theme.palette.text.secondary,
  },
  periodBadge: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.secondary,
  },
  dowRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 80.5px)',
    height: 15,
    marginBottom: 2,
  },
  dowCell: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    fontWeight: 400,
    color: theme.palette.text.disabled,
  },
  daysGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 80.5px)',
    gap: 1,
    backgroundColor: calendarTokens.border,
    border: `1px solid ${calendarTokens.border}`,
  },
  divider: {
    height: 21,
    borderRadius: 4,
    backgroundColor: calendarTokens.border,
    marginTop: theme.spacing(1),
  },

  // ── Compact (4-per-row, square tiles) ──
  rootCompact: {
    width: '100%',
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: theme.spacing(2),
    padding: theme.spacing(2, 2.5),
    boxSizing: 'border-box',
  },
  monthColumnCompact: {
    display: 'flex',
    flexDirection: 'column',
  },
  monthHeaderCompact: {
    display: 'flex',
    alignItems: 'center',
    height: 24,
    marginBottom: 4,
  },
  monthLabelCompact: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 600,
    color: theme.palette.text.secondary,
  },
  dowRowCompact: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 1fr)',
    marginBottom: 2,
  },
  dowCellCompact: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontFamily: 'Lato, sans-serif',
    fontSize: 9,
    fontWeight: 400,
    color: theme.palette.text.disabled,
    height: 14,
  },
  daysGridCompact: {
    display: 'grid',
    gridTemplateColumns: 'repeat(7, 1fr)',
  },
}));

type Props = {
  months: MonthData[];
  compact?: boolean;
  heatmapConfig?: HeatmapConfig | null;
  onDayClick?: (day: CalendarDay) => void;
};

export function CalendarMonthGrid({ months, compact, heatmapConfig, onDayClick }: Props) {
  const classes = useStyles();

  if (compact) {
    return (
      <div className={classes.rootCompact}>
        {months.map((month, mi) => (
          <div key={mi} className={classes.monthColumnCompact}>
            <div className={classes.monthHeaderCompact}>
              <span className={classes.monthLabelCompact}>{month.label}</span>
            </div>
            <div className={classes.dowRowCompact}>
              {DAYS_OF_WEEK.map((d, i) => (
                <div key={i} className={classes.dowCellCompact}>{d}</div>
              ))}
            </div>
            <div className={classes.daysGridCompact}>
              {month.days.map((day, di) => (
                <CalendarDayCell key={di} day={day} compact heatmapConfig={heatmapConfig} onClick={onDayClick} />
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  }

  return (
    <div className={classes.root}>
      {months.map((month, mi) => (
        <div key={mi} className={classes.monthColumn}>
          <div className={classes.monthHeader}>
            {month.isLocked ? (
              <div className={classes.lockedLabel}>
                <Icon className={classes.lockIcon}>lock</Icon>
                <span className={classes.periodBadge}>{month.label}</span>
              </div>
            ) : (
              <span className={classes.monthLabel}>{month.label}</span>
            )}
          </div>

          <div className={classes.dowRow}>
            {DAYS_OF_WEEK.map((d, i) => (
              <div key={i} className={classes.dowCell}>{d}</div>
            ))}
          </div>

          <div className={classes.daysGrid}>
            {month.days.map((day, di) => (
              <CalendarDayCell key={di} day={day} heatmapConfig={heatmapConfig} onClick={onDayClick} />
            ))}
          </div>

          <div className={classes.divider} />
        </div>
      ))}
    </div>
  );
}
