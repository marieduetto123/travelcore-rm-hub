import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import Checkbox from '@material-ui/core/Checkbox';
import clsx from 'clsx';
import { calendarTokens } from './tokens';
import { CloseOutModal } from './CloseOutModal';

type Props = {
  weekStart: Date;
  onBack: () => void;
  onPrevWeek: () => void;
  onNextWeek: () => void;
};

const METRICS = [
  { key: 'occ', label: 'Occupancy', format: (v: number) => `${v}%` },
  { key: 'adr', label: 'ADR', format: (v: number) => `$${v}` },
  { key: 'revenue', label: 'Revenue', format: (v: number) => `$${(v / 1000).toFixed(0)}k` },
  { key: 'rnSold', label: 'RN Sold', format: (v: number) => `${v}` },
  { key: 'availRooms', label: 'Avail Rooms', format: (v: number) => `${v}` },
  { key: 'revpar', label: 'RevPAR', format: (v: number) => `$${v}` },
];

const COMPARE_OPTIONS = [
  { value: 'none', label: 'No Compare' },
  { value: 'ly', label: 'vs LY' },
  { value: 'stly', label: 'vs STLY' },
  { value: 'forecast', label: 'vs Forecast' },
];

function generateDayData(date: Date) {
  const seed = date.getDate() * 7 + date.getMonth();
  return {
    occ: 65 + (seed % 30),
    adr: 240 + (seed % 80),
    revenue: 18000 + seed * 400,
    rnSold: 180 + (seed % 60),
    availRooms: 50 + (seed % 30),
    revpar: 156 + (seed % 60),
  };
}

const useStyles = makeStyles((theme) => ({
  root: {
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 4,
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${calendarTokens.border}`,
    flexWrap: 'wrap',
  },
  backBtn: {
    color: theme.palette.text.secondary,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    padding: theme.spacing(0.5, 1),
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  weekLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 700,
    color: theme.palette.text.primary,
    marginLeft: theme.spacing(1),
    marginRight: theme.spacing(1),
  },
  navBtn: {
    color: theme.palette.text.secondary,
    padding: 4,
    '& .MuiIcon-root': { fontSize: '18px !important' },
  },
  spacer: { flex: 1 },
  closeReopenBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    height: 32,
    padding: theme.spacing(0, 1.5),
    borderRadius: 4,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '& .MuiIcon-root': { fontSize: '14px !important', marginRight: theme.spacing(0.5) },
  },
  ghostBtn: {
    height: 32,
    padding: theme.spacing(0, 1.5),
    color: theme.palette.text.secondary,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    '& .MuiIcon-root': { fontSize: '14px !important', marginRight: theme.spacing(0.5) },
  },
  compareSelect: {
    height: 32,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    '& .MuiSelect-select': {
      padding: theme.spacing(0.5, 3.5, 0.5, 1.25),
      fontFamily: 'Lato, sans-serif',
    },
    '& .MuiOutlinedInput-notchedOutline': { borderColor: calendarTokens.border },
  },

  // Grid
  gridWrap: {
    overflowX: 'auto',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: '120px repeat(7, 1fr)',
    minWidth: 700,
  },
  rowLabel: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(0, 1.5),
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    minHeight: 40,
    backgroundColor: calendarTokens.cellBackground,
  },
  dayHeader: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: theme.spacing(1, 0.5),
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    minHeight: 52,
    cursor: 'pointer',
    '&:hover': { backgroundColor: calendarTokens.primaryHover },
    '&:last-child': { borderRight: 'none' },
  },
  dayHeaderLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.secondary,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
  },
  dayHeaderDate: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 16,
    fontWeight: 700,
    color: theme.palette.text.primary,
    lineHeight: 1.2,
  },
  dayHeaderToday: {
    color: theme.palette.primary.main,
  },
  checkboxCell: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: theme.spacing(0.25),
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    '&:last-child': { borderRight: 'none' },
  },
  dataCell: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-end',
    justifyContent: 'center',
    padding: theme.spacing(0.75, 1.25),
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    minHeight: 40,
    '&:last-child': { borderRight: 'none' },
  },
  dataCellValue: {
    fontFamily: 'Inter, sans-serif',
    fontSize: 12,
    fontWeight: 600,
    color: theme.palette.text.primary,
  },
  dataCellCompare: {
    fontFamily: 'Inter, sans-serif',
    fontSize: 10,
    color: theme.palette.text.secondary,
  },
  dataCellPositive: { color: '#15803d' },
  dataCellNegative: { color: theme.palette.error.main },
  metricRowLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
    fontWeight: 600,
  },
  cornerCell: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(0, 1.5),
    backgroundColor: calendarTokens.cellBackground,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  bulkBanner: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    padding: theme.spacing(1, 2),
    backgroundColor: calendarTokens.primaryHover,
    borderTop: `1px solid ${calendarTokens.border}`,
  },
  bulkBannerText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.primary.main,
    fontWeight: 600,
    flex: 1,
  },
  bulkCloseBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    textTransform: 'none',
    padding: theme.spacing(0.5, 1.5),
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
}));

export function WeekView({ weekStart, onBack, onPrevWeek, onNextWeek }: Props) {
  const classes = useStyles();
  const [compareMode, setCompareMode] = useState('none');
  const [selectedDays, setSelectedDays] = useState<Set<number>>(new Set());
  const [closeOutOpen, setCloseOutOpen] = useState(false);
  const [closeOutDates, setCloseOutDates] = useState<{ start: string; end: string } | undefined>();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const days: Date[] = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(weekStart);
    d.setDate(weekStart.getDate() + i);
    return d;
  });

  const dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

  const weekLabel = `${days[0].toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${days[6].toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;

  const toggleDay = (idx: number) =>
    setSelectedDays((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });

  const openCloseOut = () => {
    if (selectedDays.size > 0) {
      const sorted = [...selectedDays].sort((a, b) => a - b);
      setCloseOutDates({
        start: days[sorted[0]].toISOString().split('T')[0],
        end: days[sorted[sorted.length - 1]].toISOString().split('T')[0],
      });
    } else {
      setCloseOutDates({
        start: days[0].toISOString().split('T')[0],
        end: days[6].toISOString().split('T')[0],
      });
    }
    setCloseOutOpen(true);
  };

  return (
    <>
      <div className={classes.root}>
        {/* Header bar */}
        <div className={classes.header}>
          <Button className={classes.backBtn} onClick={onBack}>
            <Icon>arrow_back</Icon>
            Month View
          </Button>

          <IconButton className={classes.navBtn} size="small" onClick={onPrevWeek}>
            <Icon>chevron_left</Icon>
          </IconButton>
          <Typography className={classes.weekLabel}>{weekLabel}</Typography>
          <IconButton className={classes.navBtn} size="small" onClick={onNextWeek}>
            <Icon>chevron_right</Icon>
          </IconButton>

          <span className={classes.spacer} />

          <Button className={classes.closeReopenBtn} onClick={openCloseOut} disableElevation>
            <Icon>lock</Icon>
            Close/Re-Open
          </Button>

          <Button className={classes.ghostBtn}>
            <Icon>filter_list</Icon>
            Filters
          </Button>

          <Select
            value={compareMode}
            onChange={(e) => setCompareMode(e.target.value as string)}
            variant="outlined"
            className={classes.compareSelect}
          >
            {COMPARE_OPTIONS.map((o) => (
              <MenuItem key={o.value} value={o.value} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>
                {o.label}
              </MenuItem>
            ))}
          </Select>
        </div>

        {/* Grid */}
        <div className={classes.gridWrap}>
          <div className={classes.grid}>
            {/* Corner header */}
            <div className={classes.cornerCell} />

            {/* Day headers */}
            {days.map((day, i) => {
              const isToday = day.getTime() === today.getTime();
              return (
                <div key={i} className={classes.dayHeader}>
                  <Typography className={classes.dayHeaderLabel}>{dayNames[i]}</Typography>
                  <Typography className={clsx(classes.dayHeaderDate, isToday && classes.dayHeaderToday)}>
                    {day.getDate()}
                  </Typography>
                </div>
              );
            })}

            {/* Checkbox row */}
            <div className={classes.rowLabel}>
              <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 11, color: 'inherit' }}>
                Select
              </Typography>
            </div>
            {days.map((_, i) => (
              <div key={i} className={classes.checkboxCell}>
                <Checkbox
                  checked={selectedDays.has(i)}
                  onChange={() => toggleDay(i)}
                  size="small"
                  style={{ color: selectedDays.has(i) ? undefined : calendarTokens.checkboxBorder }}
                />
              </div>
            ))}

            {/* Metric rows */}
            {METRICS.map(({ key, label, format }) => (
              <React.Fragment key={key}>
                <div className={classes.rowLabel}>
                  <Typography className={classes.metricRowLabel}>{label}</Typography>
                </div>
                {days.map((day, i) => {
                  const data = generateDayData(day);
                  const val = data[key as keyof typeof data] as number;
                  const compareSeed = (day.getDate() * 3 + 5) % 20;
                  const comparePct = compareMode !== 'none' ? compareSeed - 10 : null;
                  return (
                    <div key={i} className={classes.dataCell}>
                      <Typography className={classes.dataCellValue}>{format(val)}</Typography>
                      {comparePct !== null && (
                        <Typography
                          className={clsx(
                            classes.dataCellCompare,
                            comparePct > 0 && classes.dataCellPositive,
                            comparePct < 0 && classes.dataCellNegative,
                          )}
                        >
                          {comparePct > 0 ? '+' : ''}{comparePct}%
                        </Typography>
                      )}
                    </div>
                  );
                })}
              </React.Fragment>
            ))}
          </div>
        </div>

        {/* Bulk close banner */}
        {selectedDays.size > 0 && (
          <div className={classes.bulkBanner}>
            <Typography className={classes.bulkBannerText}>
              {selectedDays.size} day{selectedDays.size > 1 ? 's' : ''} selected
            </Typography>
            <Button
              className={classes.bulkCloseBtn}
              variant="contained"
              disableElevation
              onClick={openCloseOut}
            >
              Close/Re-Open Selected
            </Button>
            <Button
              style={{ fontFamily: 'Lato, sans-serif', fontSize: 12, textTransform: 'none' }}
              onClick={() => setSelectedDays(new Set())}
            >
              Clear
            </Button>
          </div>
        )}
      </div>

      <CloseOutModal
        open={closeOutOpen}
        onClose={() => setCloseOutOpen(false)}
        initialDates={closeOutDates}
      />
    </>
  );
}
