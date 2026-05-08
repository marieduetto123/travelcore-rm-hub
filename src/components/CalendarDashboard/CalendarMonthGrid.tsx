import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { CalendarDayCell } from './CalendarDayCell';
import { CalendarDay, MonthData } from './types';
import { calendarTokens } from './tokens';

const DAYS_OF_WEEK = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

const useStyles = makeStyles((theme) => ({
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
  },
  divider: {
    height: 21,
    borderRadius: 4,
    backgroundColor: calendarTokens.border,
    marginTop: theme.spacing(1),
  },
}));

type Props = {
  months: [MonthData, MonthData];
  onDayClick?: (day: CalendarDay) => void;
};

export function CalendarMonthGrid({ months, onDayClick }: Props) {
  const classes = useStyles();

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
              <CalendarDayCell key={di} day={day} onClick={onDayClick} />
            ))}
          </div>

          <div className={classes.divider} />
        </div>
      ))}
    </div>
  );
}
