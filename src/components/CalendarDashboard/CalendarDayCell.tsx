import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { CalendarDay } from './types';
import { calendarTokens } from './tokens';

const useStyles = makeStyles((theme) => ({
  root: {
    position: 'relative',
    width: 80.5,
    height: 130,
    border: `1px solid ${calendarTokens.border}`,
    padding: theme.spacing(1.125),
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    cursor: 'pointer',
    flexShrink: 0,
    '&:hover $visibilityBtn': {
      opacity: 1,
    },
  },
  normal: {
    backgroundColor: calendarTokens.cellBackground,
  },
  closed: {
    backgroundColor: calendarTokens.cellBackgroundClosed,
  },
  empty: {
    visibility: 'hidden',
    cursor: 'default',
  },
  topSection: {
    position: 'relative',
    width: '100%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    paddingBottom: 2,
    flexShrink: 0,
  },
  checkbox: {
    position: 'absolute',
    left: 2,
    top: 1,
    width: 12,
    height: 12,
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${calendarTokens.checkboxBorder}`,
    borderRadius: 2.5,
    boxSizing: 'border-box',
    flexShrink: 0,
  },
  dayNumber: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 11,
    lineHeight: 'normal',
    color: theme.palette.text.primary,
    whiteSpace: 'nowrap',
  },
  lockIcon: {
    position: 'absolute',
    left: 1,
    top: '50%',
    transform: 'translateY(-50%)',
    fontSize: '14px !important',
    color: theme.palette.error.main,
    lineHeight: '14px',
  },
  visibilityBtn: {
    position: 'absolute',
    right: 2,
    top: '50%',
    transform: 'translateY(-50%)',
    width: 20,
    height: 20,
    borderRadius: 4,
    backgroundColor: calendarTokens.primaryHover,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    opacity: 0,
    transition: theme.transitions.create('opacity'),
    '& .MuiIcon-root': {
      fontSize: '16px !important',
      color: theme.palette.primary.main,
    },
  },
  metricsSection: {
    flex: '1 0 0',
    width: '100%',
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'flex-end',
    paddingBottom: 16,
    minHeight: 58.78,
  },
  metricsContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: 2,
    width: '100%',
  },
  metricRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    overflow: 'hidden',
    width: '100%',
    height: 13.2,
  },
  metricLabel: {
    fontFamily: 'Inter, sans-serif',
    fontWeight: 500,
    fontSize: 12,
    lineHeight: '13.2px',
    color: theme.palette.text.primary,
    overflow: 'hidden',
    whiteSpace: 'nowrap',
  },
  metricValue: {
    fontFamily: 'Inter, sans-serif',
    fontWeight: 600,
    fontSize: 12,
    lineHeight: '13.2px',
    color: theme.palette.text.primary,
    textAlign: 'right',
    whiteSpace: 'nowrap',
  },
  compareText: {
    color: calendarTokens.benchmarkColor,
  },
  todayCorner: {
    position: 'absolute',
    bottom: 4,
    right: 4,
    display: 'flex',
    '& .MuiIcon-root': {
      fontSize: '16px !important',
      color: theme.palette.primary.main,
      lineHeight: '16px',
    },
  },
}));

type Props = {
  day: CalendarDay;
  onClick?: (day: CalendarDay) => void;
};

export function CalendarDayCell({ day, onClick }: Props) {
  const classes = useStyles();

  if (!day.isInMonth) {
    return <div className={clsx(classes.root, classes.normal, classes.empty)} />;
  }

  return (
    <div
      className={clsx(classes.root, day.isClosed ? classes.closed : classes.normal)}
      onClick={() => onClick?.(day)}
    >
      <div className={classes.topSection}>
        <div className={classes.checkbox} />
        <span className={classes.dayNumber}>{day.dayNumber}</span>
        {day.isClosed && (
          <Icon className={classes.lockIcon}>lock</Icon>
        )}
        <div className={classes.visibilityBtn}>
          <Icon>visibility</Icon>
        </div>
      </div>

      <div className={classes.metricsSection}>
        <div className={classes.metricsContainer}>
          {day.metrics.map((metric, index) => (
            <div key={index} className={classes.metricRow}>
              <span className={clsx(classes.metricLabel, metric.isCompare && classes.compareText)}>
                {metric.label}
              </span>
              <span className={clsx(classes.metricValue, metric.isCompare && classes.compareText)}>
                {metric.value}
              </span>
            </div>
          ))}
        </div>
      </div>

      {day.isToday && (
        <div className={classes.todayCorner}>
          <Icon>today</Icon>
        </div>
      )}
    </div>
  );
}
