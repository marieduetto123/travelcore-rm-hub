import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { CalendarDay } from './types';
import { calendarTokens } from './tokens';

const useStyles = makeStyles((theme) => ({
  root: {
    width: 80.5,
    height: 130,
    border: `1px solid ${calendarTokens.border}`,
    padding: 4,
    display: 'flex',
    flexDirection: 'column',
    gap: 2,
    boxSizing: 'border-box',
    cursor: 'pointer',
    flexShrink: 0,
    backgroundColor: theme.palette.common.white,
    // reveal hover-only controls when this cell is hovered
    '&:hover $checkbox': { visibility: 'visible' },
    '&:hover $rightIcon': { visibility: 'visible' },
  },
  closed: {
    backgroundColor: calendarTokens.cellBackgroundClosed,
  },
  empty: {
    visibility: 'hidden',
    cursor: 'default',
  },

  // Top row: checkbox | date | icon
  topRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    width: '100%',
    flexShrink: 0,
  },
  checkbox: {
    width: 18,
    height: 18,
    backgroundColor: theme.palette.common.white,
    border: `2px solid ${calendarTokens.checkboxBorder}`,
    borderRadius: 2,
    boxSizing: 'border-box',
    flexShrink: 0,
    visibility: 'hidden',
  },
  dateContainer: {
    width: 24,
    height: 24,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  dayNumber: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 10,
    lineHeight: 'normal',
    color: '#252525',
    whiteSpace: 'nowrap',
  },
  dayNumberDimmed: {
    opacity: 0.5,
  },
  rightIcon: {
    fontSize: '18px !important',
    color: theme.palette.text.secondary,
    visibility: 'hidden',
  },

  // Metric rows
  metricRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    width: '100%',
    padding: '2px 4px',
    borderRadius: 2,
    flexShrink: 0,
    overflow: 'hidden',
    boxSizing: 'border-box',
    backgroundColor: calendarTokens.cellBackground,
  },
  metricRowCompare: {
    backgroundColor: calendarTokens.compareRowBackground,
  },
  metricText: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 8,
    lineHeight: 1.25,
    color: '#000000',
    whiteSpace: 'nowrap',
  },

  // Bottom spacer — today icon floats bottom-right
  bottomSpacer: {
    flex: '1 0 0',
    display: 'flex',
    alignItems: 'flex-end',
    justifyContent: 'flex-end',
    minHeight: 1,
  },
  todayIcon: {
    fontSize: '14px !important',
    color: theme.palette.primary.main,
  },

  // ── Compact mode (square tile, no metrics) ──
  rootCompact: {
    width: '100%',
    height: 32,
    padding: 0,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    border: `1px solid ${calendarTokens.border}`,
    boxSizing: 'border-box',
    cursor: 'pointer',
    backgroundColor: theme.palette.common.white,
  },
  closedCompact: {
    backgroundColor: calendarTokens.cellBackgroundClosed,
  },
  emptyCompact: {
    visibility: 'hidden',
    cursor: 'default',
  },
  dayNumberCompact: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 10,
    color: '#252525',
  },
}));

type Props = {
  day: CalendarDay;
  compact?: boolean;
  onClick?: (day: CalendarDay) => void;
};

export function CalendarDayCell({ day, compact, onClick }: Props) {
  const classes = useStyles();

  if (compact) {
    if (!day.isInMonth) return <div className={clsx(classes.rootCompact, classes.emptyCompact)} />;
    return (
      <div
        className={clsx(classes.rootCompact, day.isClosed && classes.closedCompact)}
        onClick={() => onClick?.(day)}
      >
        <span className={classes.dayNumberCompact}>{day.dayNumber}</span>
      </div>
    );
  }

  if (!day.isInMonth) {
    return <div className={clsx(classes.root, classes.empty)} />;
  }

  return (
    <div
      className={clsx(classes.root, day.isClosed && classes.closed)}
      onClick={() => onClick?.(day)}
    >
      {/* Top row: checkbox | day number | icon */}
      <div className={classes.topRow}>
        <div className={classes.checkbox} />

        <div className={classes.dateContainer}>
          <span className={classes.dayNumber}>{day.dayNumber}</span>
        </div>

        <Icon className={classes.rightIcon}>
          {day.isClosed ? 'lock' : 'visibility'}
        </Icon>
      </div>

      {/* Metric rows */}
      {day.metrics.map((metric, index) => (
        <div
          key={index}
          className={clsx(classes.metricRow, metric.isCompare && classes.metricRowCompare)}
        >
          <span className={classes.metricText}>{metric.label}</span>
          <span className={classes.metricText}>{metric.value}</span>
        </div>
      ))}

      {/* Spacer — today icon sits at bottom-right */}
      <div className={classes.bottomSpacer}>
        {day.isToday && (
          <Icon className={classes.todayIcon}>calendar_today</Icon>
        )}
      </div>
    </div>
  );
}
