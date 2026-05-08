import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Paper from '@material-ui/core/Paper';
import Typography from '@material-ui/core/Typography';
import Icon from '@material-ui/core/Icon';
import Divider from '@material-ui/core/Divider';
import Tab from '@material-ui/core/Tab';
import Tabs from '@material-ui/core/Tabs';
import { CalendarHeader } from './CalendarHeader';
import { CalendarMonthGrid } from './CalendarMonthGrid';
import { DayDetailPopup } from './DayDetailPopup';
import { WeekView } from './WeekView';
import { CalendarDay, DayDetailGroup, MonthData } from './types';
import { buildMonthData } from './calendarUtils';
import { DateRange, MonthRef } from './DateRangePickerDropdown';
import { HeatmapConfig } from './HeatmapModal';
import { calendarTokens } from './tokens';

const useStyles = makeStyles((theme) => ({
  root: {
    padding: theme.spacing(3),
    display: 'flex',
    flexDirection: 'column',
    gap: theme.spacing(2.5),
  },

  // Trends section (placeholder)
  trendsCard: {
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 4,
  },
  trendsHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  trendsTitle: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 16,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  trendsPlaceholder: {
    height: 160,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: theme.palette.text.disabled,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    gap: theme.spacing(1),
  },

  // Calendar card
  calendarCard: {
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 4,
    position: 'relative',
    overflow: 'visible',
  },

  // Tab bar
  tabBar: {
    borderBottom: `1px solid ${calendarTokens.border}`,
    minHeight: 52,
    padding: theme.spacing(0, 2),
    '& .MuiTab-root': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 14,
      textTransform: 'none',
      minHeight: 52,
      color: theme.palette.text.secondary,
    },
    '& .Mui-selected': {
      color: theme.palette.primary.main,
      fontWeight: 700,
    },
    '& .MuiTabs-indicator': {
      backgroundColor: theme.palette.primary.main,
    },
  },

  // Legend row
  legendRow: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(2),
    padding: theme.spacing(1.25, 2.5),
    borderBottom: `1px solid ${calendarTokens.border}`,
    flexWrap: 'wrap',
  },
  legendItem: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.5),
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
  },
  legendSwatch: {
    width: 10,
    height: 10,
    borderRadius: 2,
    flexShrink: 0,
  },
  legendSwatchHotel: {
    backgroundColor: calendarTokens.cellBackground,
  },
  legendSwatchOperator: {
    backgroundColor: calendarTokens.compareRowBackground,
  },
  legendIcon: {
    fontSize: '16px !important',
    color: theme.palette.text.secondary,
  },

  // Monthly metrics
  metricsSection: {
    padding: theme.spacing(0, 2.5, 3),
  },
  metricsMonthHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 0),
    borderBottom: `1px solid ${calendarTokens.border}`,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  metricsRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(1, 0),
    borderBottom: `1px solid ${calendarTokens.border}`,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: calendarTokens.cellBackground,
    },
  },
  metricsRowLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.5),
  },
  metricsRowValues: {
    display: 'flex',
    gap: theme.spacing(4),
  },
  metricsValue: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 600,
    color: theme.palette.text.primary,
    textAlign: 'right',
    minWidth: 80,
  },
  progressBarWrapper: {
    flex: 1,
    maxWidth: 200,
    height: 8,
    borderRadius: 4,
    backgroundColor: calendarTokens.border,
    overflow: 'hidden',
    margin: theme.spacing(0, 2),
  },
  progressBarFill: {
    height: '100%',
    borderRadius: 4,
    background: `linear-gradient(90deg, ${theme.palette.primary.main}, #00a39f)`,
  },
}));

const SAMPLE_METRICS_ROWS = [
  { label: 'Total Guests', values: ['$686k', '$710k'], pct: 0.78 },
  { label: 'Occupancy', values: ['91%', '86%'], pct: 0.91 },
  { label: 'ADR', values: ['$285', '$301'], pct: 0.65 },
  { label: 'RevPAR', values: ['$259', '$259'], pct: 0.59 },
  { label: 'Available Rooms', values: ['1,240', '1,240'], pct: 1 },
];

const SAMPLE_DAY_DETAIL: DayDetailGroup[] = [
  {
    title: 'Revenue',
    isExpanded: true,
    items: [
      { label: 'AI', percentage: 65, seats: 286, value: '$18,590' },
      { label: 'BB', percentage: 28, seats: 124, value: '$8,060' },
      { label: 'HB', percentage: 12, seats: 54, value: '$3,510' },
      { label: 'RO', percentage: -5, seats: -22, value: '-$1,430', isNegative: true },
    ],
  },
  {
    title: 'Operators',
    isExpanded: true,
    items: [
      { label: 'TO', percentage: 35, value: '$9,940' },
      { label: 'D', percentage: 30, value: '$8,520' },
      { label: 'OTA', percentage: 20, value: '$5,680' },
      { label: 'Oth', percentage: 15, value: '$4,260' },
    ],
  },
  { title: 'Contract Rates', isExpanded: false, items: [{ label: 'Operator A', value: '$285' }] },
  { title: 'Summary', isExpanded: false, items: [] },
];

export function CalendarDashboardPage() {
  const classes = useStyles();
  const [activeTab, setActiveTab] = useState(0);
  const [selectedDay, setSelectedDay] = useState<CalendarDay | null>(null);
  const [detailGroups, setDetailGroups] = useState<DayDetailGroup[]>(SAMPLE_DAY_DETAIL);
  const [weekViewDay, setWeekViewDay] = useState<Date | null>(null);

  const [calRange, setCalRange] = useState<DateRange>({
    start: { year: 2026, month: 0 },
    end:   { year: 2026, month: 1 },
  });

  const monthCount =
    (calRange.end.year - calRange.start.year) * 12 +
    (calRange.end.month - calRange.start.month) + 1;
  const compact = monthCount >= 3;

  const months: MonthData[] = (() => {
    const result: MonthData[] = [];
    for (let i = 0; i < monthCount; i++) {
      const total = calRange.start.year * 12 + calRange.start.month + i;
      const y = Math.floor(total / 12);
      const m = total % 12;
      result.push(buildMonthData(y, m, i === monthCount - 1 ? { isLocked: true } : undefined));
    }
    return result;
  })();

  // Monday-align the week that contains the clicked day
  const getWeekStart = (date: Date): Date => {
    const d = new Date(date);
    const day = d.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diff);
    d.setHours(0, 0, 0, 0);
    return d;
  };

  const [weekStart, setWeekStart] = useState<Date>(getWeekStart(new Date()));
  const [heatmapConfig, setHeatmapConfig] = useState<HeatmapConfig | null>(null);

  const handleDayClick = (day: CalendarDay) => {
    if (activeTab === 0) {
      // Monthly → switch to Weekly tab and show week view
      setActiveTab(1);
      setWeekStart(getWeekStart(day.date));
      setWeekViewDay(day.date);
      setSelectedDay(null);
    } else {
      setSelectedDay(day);
      setDetailGroups(SAMPLE_DAY_DETAIL);
    }
  };

  const handleToggleGroup = (index: number) => {
    setDetailGroups((prev) =>
      prev.map((g, i) => (i === index ? { ...g, isExpanded: !g.isExpanded } : g))
    );
  };

  const shiftWeek = (dir: number) => {
    setWeekStart((d) => {
      const next = new Date(d);
      next.setDate(d.getDate() + dir * 7);
      return next;
    });
  };

  return (
    <div className={classes.root}>
      {/* Trends card (placeholder) */}
      <Paper className={classes.trendsCard} elevation={0}>
        <div className={classes.trendsHeader}>
          <Typography className={classes.trendsTitle}>Trends</Typography>
          <Icon style={{ fontSize: 18, color: '#9ca3af' }}>more_horiz</Icon>
        </div>
        <div className={classes.trendsPlaceholder}>
          <Icon style={{ fontSize: 32, color: '#d1d5db' }}>show_chart</Icon>
          <span>Chart area — connect a charting library</span>
        </div>
      </Paper>

      {/* Calendar card */}
      <Paper className={classes.calendarCard} elevation={0}>
        <CalendarHeader
          onRangeChange={setCalRange}
          onHeatmapApply={(cfg) => setHeatmapConfig(cfg.type ? cfg : null)}
        />

        {/* Tab bar */}
        <Tabs
          value={activeTab}
          onChange={(_, v) => { setActiveTab(v); setWeekViewDay(null); setSelectedDay(null); }}
          className={classes.tabBar}
        >
          <Tab label="Monthly" />
          <Tab label="Weekly" />
        </Tabs>

        {/* Week view */}
        {weekViewDay ? (
          <div style={{ padding: 16 }}>
            <WeekView
              weekStart={weekStart}
              onBack={() => setWeekViewDay(null)}
              onPrevWeek={() => shiftWeek(-1)}
              onNextWeek={() => shiftWeek(1)}
            />
          </div>
        ) : (
          <>
            {/* Legend */}
            <div className={classes.legendRow}>
              <div className={classes.legendItem}>
                <Icon className={classes.legendIcon}>visibility</Icon>
                Visible
              </div>
              <div className={classes.legendItem}>
                <Icon className={classes.legendIcon}>lock</Icon>
                Closed
              </div>
              <div className={classes.legendItem}>
                <Icon className={classes.legendIcon}>today</Icon>
                Today
              </div>
              <div className={classes.legendItem}>
                <div className={`${classes.legendSwatch} ${classes.legendSwatchHotel}`} />
                Hotel
              </div>
              <div className={classes.legendItem}>
                <div className={`${classes.legendSwatch} ${classes.legendSwatchOperator}`} />
                Operator
              </div>
            </div>

            {/* Month grid */}
            <CalendarMonthGrid months={months} compact={compact} heatmapConfig={heatmapConfig} onDayClick={handleDayClick} />

            {/* Monthly metrics */}
            <Divider />
            <div className={classes.metricsSection}>
              <div className={classes.metricsMonthHeader}>
                {months.map(m => m.label).join(' · ')}
              </div>

              {SAMPLE_METRICS_ROWS.map((row) => (
                <div key={row.label} className={classes.metricsRow}>
                  <Typography className={classes.metricsRowLabel}>{row.label}</Typography>
                  <div className={classes.progressBarWrapper}>
                    <div className={classes.progressBarFill} style={{ width: `${row.pct * 100}%` }} />
                  </div>
                  <div className={classes.metricsRowValues}>
                    {row.values.map((v, vi) => (
                      <Typography key={vi} className={classes.metricsValue}>{v}</Typography>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {/* Day detail popup */}
        {selectedDay && (
          <DayDetailPopup
            date={selectedDay.date.toLocaleDateString('en-US', {
              weekday: 'short',
              month: 'short',
              day: 'numeric',
              year: 'numeric',
            })}
            groups={detailGroups}
            onClose={() => setSelectedDay(null)}
            onToggleGroup={handleToggleGroup}
            style={{ top: 800, right: 24 }}
          />
        )}
      </Paper>
    </div>
  );
}
