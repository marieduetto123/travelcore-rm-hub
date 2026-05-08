import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { CalendarDay } from './types';
import { calendarTokens } from './tokens';
import { HeatmapConfig, HeatmapType } from './HeatmapModal';

function resolveHeatmapBg(day: CalendarDay, cfg: HeatmapConfig): string | null {
  if (!cfg.type || !day.isInMonth) return null;

  // Condition gate
  if (cfg.conditionEnabled) {
    const condVal = extractValue(day, cfg.conditionMetric as HeatmapType);
    if (!evalCondition(condVal, cfg.conditionOp, cfg.conditionValue)) return null;
  }

  if (cfg.type === 'stop_sales') {
    return day.isClosed ? 'rgba(220,38,38,0.28)' : null;
  }

  const val = extractValue(day, cfg.type);
  const { grey, green, blue } = cfg.thresholds;
  if (val <= grey)  return 'rgba(156,163,175,0.32)';
  if (val <= green) return 'rgba(22,163,74,0.24)';
  if (val <= blue)  return 'rgba(37,99,235,0.22)';
  return 'rgba(220,38,38,0.24)';
}

function extractValue(day: CalendarDay, type: HeatmapType | string): number {
  if (type === 'hotel_occ') {
    const row = day.metrics.find(m => m.label === 'Occ' && !m.isCompare);
    return row ? parseFloat(row.value.replace('%','')) || 0 : 0;
  }
  // Deterministic mock for other types seeded by day number
  const seed = day.dayNumber || 1;
  const offsets: Record<string, number> = { remaining_rooms: 17, meal_plan: 37, to_forecast: 53 };
  return ((seed * 13 + (offsets[type] ?? 7)) % 101);
}

function evalCondition(val: number, op: string, threshold: number): boolean {
  if (op === '>')  return val > threshold;
  if (op === '>=') return val >= threshold;
  if (op === '<')  return val < threshold;
  if (op === '<=') return val <= threshold;
  return true;
}

const useStyles = makeStyles((theme) => ({
  root: {
    width: 80.5,
    height: 130,
    border: 'none',
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
  heatmapConfig?: HeatmapConfig | null;
  onClick?: (day: CalendarDay) => void;
};

export function CalendarDayCell({ day, compact, heatmapConfig, onClick }: Props) {
  const classes = useStyles();
  const heatmapBg = heatmapConfig ? resolveHeatmapBg(day, heatmapConfig) : null;

  if (compact) {
    if (!day.isInMonth) return <div className={clsx(classes.rootCompact, classes.emptyCompact)} />;
    return (
      <div
        className={clsx(classes.rootCompact, day.isClosed && classes.closedCompact)}
        style={heatmapBg ? { backgroundColor: heatmapBg } : undefined}
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
      style={heatmapBg ? { backgroundColor: heatmapBg } : undefined}
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
