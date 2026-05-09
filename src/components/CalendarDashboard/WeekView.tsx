import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import clsx from 'clsx';
import { calendarTokens } from './tokens';
import { CloseOutModal } from './CloseOutModal';

type Props = {
  weekStart: Date;
  onBack: () => void;
  onPrevWeek: () => void;
  onNextWeek: () => void;
};

type RowDef = {
  key: string;
  label: string;
  dot: string;
  format: (v: number) => string;
  showBar?: boolean;
  indent?: boolean;
};

type SectionDef = {
  key: string;
  label: string;
  accent: string;
  rows: RowDef[];
};

const pct  = (v: number) => `${v}%`;
const dollar = (v: number) => `$${v}`;
const money  = (v: number) => `$${(v / 1000).toFixed(0)}k`;
const num    = (v: number) => `${v}`;

const SECTIONS: SectionDef[] = [
  {
    key: 'body',
    label: 'Body Metrics',
    accent: '#004948',
    rows: [
      { key: 'total_occ', label: 'Total Hotel Occupancy', dot: '#004948', format: pct, showBar: true },
      { key: 'td_hub',    label: 'Travel Distribution Hubs', dot: '#7c6fe0', format: pct, indent: true },
      { key: 'online',    label: 'Online + Office',          dot: '#f59e0b', format: pct, indent: true },
      { key: 'mice',      label: 'MICE',                     dot: '#3b82f6', format: pct, indent: true },
      { key: 'b2b',       label: 'B2B',                      dot: '#10b981', format: pct, indent: true },
      { key: 'other_occ', label: 'Other',                    dot: '#f97316', format: pct, indent: true },
      { key: 'total_rev', label: 'Total Hotel Revenue',      dot: '#004948', format: money },
      { key: 'adr',       label: 'ADR',                      dot: '#7c6fe0', format: dollar },
      { key: 'revpar',    label: 'RevPAR',                   dot: '#3b82f6', format: dollar },
      { key: 'rn_sold',   label: 'RN Sold',                  dot: '#f59e0b', format: num },
      { key: 'avg_los',   label: 'Avg LOS',                  dot: '#10b981', format: (v) => `${(v / 10).toFixed(1)}` },
    ],
  },
  {
    key: 'meal',
    label: 'Meal Plans',
    accent: '#f59e0b',
    rows: [
      { key: 'mp_ai', label: 'All Inclusive',   dot: '#004948', format: pct },
      { key: 'mp_bb', label: 'Bed & Breakfast', dot: '#7c6fe0', format: pct },
      { key: 'mp_hb', label: 'Half Board',      dot: '#f59e0b', format: pct },
      { key: 'mp_ro', label: 'Room Only',       dot: '#3b82f6', format: pct },
    ],
  },
  {
    key: 'mix',
    label: 'Business Mix',
    accent: '#3b82f6',
    rows: [
      { key: 'bm_direct', label: 'Direct',        dot: '#004948', format: pct },
      { key: 'bm_ota',    label: 'OTA',           dot: '#7c6fe0', format: pct },
      { key: 'bm_to',     label: 'Tour Operator', dot: '#f59e0b', format: pct },
      { key: 'bm_other',  label: 'Other',         dot: '#3b82f6', format: pct },
    ],
  },
  {
    key: 'avail',
    label: 'Room Availability',
    accent: '#64748b',
    rows: [
      { key: 'av_std', label: 'Standard', dot: '#004948', format: (v) => `${v} rm` },
      { key: 'av_sup', label: 'Superior', dot: '#7c6fe0', format: (v) => `${v} rm` },
      { key: 'av_dlx', label: 'Deluxe',  dot: '#f59e0b', format: (v) => `${v} rm` },
      { key: 'av_ste', label: 'Suite',   dot: '#3b82f6', format: (v) => `${v} rm` },
    ],
  },
  {
    key: 'rates',
    label: 'Travel Co. Rates',
    accent: '#7c6fe0',
    rows: [
      { key: 'rate_bar',      label: 'Best Available Rate', dot: '#004948', format: dollar },
      { key: 'rate_promo',    label: 'Promotional Rate',    dot: '#7c6fe0', format: dollar },
      { key: 'rate_contract', label: 'Contract Rate',       dot: '#f59e0b', format: dollar },
    ],
  },
];

const COMPARE_OPTIONS = [
  { value: 'none',     label: 'No Compare' },
  { value: 'ly',       label: 'vs LY' },
  { value: 'stly',     label: 'vs STLY' },
  { value: 'forecast', label: 'vs Forecast' },
];

const ROW_RANGES: Record<string, [number, number]> = {
  total_occ: [55, 93], td_hub: [10, 38], online: [15, 48], mice: [5, 22],
  b2b: [8, 28], other_occ: [2, 14],
  total_rev: [15000, 34000], adr: [188, 318], revpar: [128, 248], rn_sold: [118, 218], avg_los: [22, 58],
  mp_ai: [24, 58], mp_bb: [14, 44], mp_hb: [9, 32], mp_ro: [4, 24],
  bm_direct: [19, 44], bm_ota: [14, 40], bm_to: [18, 48], bm_other: [4, 20],
  av_std: [8, 38], av_sup: [4, 22], av_dlx: [2, 14], av_ste: [1, 8],
  rate_bar: [182, 318], rate_promo: [142, 262], rate_contract: [158, 288],
};

function hashSeed(a: number, b: number): number {
  return (((a * 1664525 + 1013904223) ^ (b * 22695477 + 1)) >>> 0);
}

function genValue(dayIdx: number, rowKey: string): number {
  let h = dayIdx;
  for (let i = 0; i < rowKey.length; i++) h = hashSeed(h, rowKey.charCodeAt(i));
  const [lo, hi] = ROW_RANGES[rowKey] ?? [10, 100];
  return lo + (h % (hi - lo));
}

function genDelta(dayIdx: number, rowKey: string): number {
  let h = dayIdx + 37;
  for (let i = 0; i < rowKey.length; i++) h = hashSeed(h, rowKey.charCodeAt(i) + 50);
  return ((h % 201) - 100) / 10;
}

const useStyles = makeStyles((theme) => ({
  root: {
    backgroundColor: theme.palette.common.white,
    borderRadius: 4,
    overflow: 'hidden',
  },

  // ── Header ──────────────────────────────────────────────────────────────────
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
    minWidth: 0,
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  weekLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 700,
    color: theme.palette.text.primary,
    margin: theme.spacing(0, 0.5),
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
    '& .MuiSelect-select': { padding: theme.spacing(0.5, 3.5, 0.5, 1.25), fontFamily: 'Lato, sans-serif' },
    '& .MuiOutlinedInput-notchedOutline': { borderColor: calendarTokens.border },
  },

  // ── Grid ────────────────────────────────────────────────────────────────────
  gridWrap: { overflowX: 'auto' },
  grid: {
    display: 'grid',
    gridTemplateColumns: '200px repeat(7, minmax(110px, 1fr))',
    minWidth: 970,
  },

  // ── Day header row ───────────────────────────────────────────────────────────
  cornerCell: {
    height: 47,
    backgroundColor: calendarTokens.cellBackground,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  dayHeader: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: 47,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    backgroundColor: theme.palette.common.white,
    '&:last-child': { borderRight: 'none' },
  },
  dayName: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 10,
    fontWeight: 700,
    textTransform: 'uppercase',
    letterSpacing: 0.8,
    color: theme.palette.text.secondary,
  },
  dateNum: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 18,
    fontWeight: 700,
    lineHeight: 1,
    color: theme.palette.text.primary,
  },
  dateNumToday: { color: theme.palette.primary.main },
  dateMonth: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 10,
    color: theme.palette.text.disabled,
    lineHeight: 1.4,
  },

  // ── Close Out row ───────────────────────────────────────────────────────────
  closeOutLabelCell: {
    display: 'flex',
    alignItems: 'center',
    gap: 5,
    padding: theme.spacing(0, 1.5),
    height: 40,
    backgroundColor: calendarTokens.cellBackground,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    cursor: 'pointer',
    '&:hover': { backgroundColor: '#ebebeb' },
  },
  closeOutLabelText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    color: theme.palette.primary.main,
    textDecoration: 'underline',
    textDecorationColor: 'transparent',
    transition: 'text-decoration-color 0.1s',
    '$closeOutLabelCell:hover &': { textDecorationColor: theme.palette.primary.main },
  },
  closeOutCell: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    height: 40,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    cursor: 'pointer',
    '&:hover': { backgroundColor: calendarTokens.primaryHover },
    '&:last-child': { borderRight: 'none' },
  },
  lockIcon:  { fontSize: '16px !important', color: '#e53935' },
  visIcon:   { fontSize: '16px !important', color: theme.palette.primary.main },

  // ── Section header (spans all columns) ───────────────────────────────────
  sectionHeader: {
    gridColumn: '1 / -1',
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    padding: theme.spacing(0, 2),
    height: 45,
    backgroundColor: calendarTokens.cellBackground,
    borderBottom: `1px solid ${calendarTokens.border}`,
    cursor: 'pointer',
    userSelect: 'none',
    '&:hover': { backgroundColor: '#ebebeb' },
  },
  sectionAccent: {
    width: 3,
    height: 16,
    borderRadius: 2,
    flexShrink: 0,
  },
  sectionTitle: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    color: theme.palette.text.primary,
    flex: 1,
    letterSpacing: 0.2,
  },
  chevron: {
    fontSize: '18px !important',
    color: theme.palette.text.secondary,
    transition: 'transform 0.15s',
  },
  chevronCollapsed: { transform: 'rotate(-90deg)' },

  // ── Label cell ───────────────────────────────────────────────────────────
  labelCell: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    padding: theme.spacing(0, 1.5),
    height: 33,
    backgroundColor: calendarTokens.cellBackground,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  labelCellBar: {
    height: 52,
    alignItems: 'flex-start',
    paddingTop: 10,
  },
  labelIndent: { paddingLeft: theme.spacing(3.5) },
  dot: {
    width: 7,
    height: 7,
    borderRadius: '50%',
    flexShrink: 0,
    marginTop: 1,
  },
  labelText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.secondary,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    lineHeight: 1.3,
  },

  // ── Data cell ────────────────────────────────────────────────────────────
  dataCell: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-end',
    justifyContent: 'center',
    padding: theme.spacing(0.5, 1.25),
    height: 33,
    borderRight: `1px solid ${calendarTokens.border}`,
    borderBottom: `1px solid ${calendarTokens.border}`,
    '&:last-child': { borderRight: 'none' },
  },
  dataCellBar: {
    height: 52,
    justifyContent: 'flex-start',
    paddingTop: 8,
  },
  value: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    color: theme.palette.text.primary,
    lineHeight: 1,
  },

  // ── Delta badge ──────────────────────────────────────────────────────────
  badge: {
    display: 'inline-flex',
    alignItems: 'center',
    padding: '1px 4px',
    borderRadius: 3,
    fontFamily: 'Lato, sans-serif',
    fontSize: 9,
    fontWeight: 700,
    lineHeight: 1.5,
    marginTop: 3,
    whiteSpace: 'nowrap',
  },
  badgePos:     { backgroundColor: '#d1fae5', color: '#065f46' },
  badgeNeg:     { backgroundColor: '#fee2e2', color: '#991b1b' },
  badgeNeutral: { backgroundColor: '#f3f4f6', color: '#6b7280' },

  // ── Occupancy bar ────────────────────────────────────────────────────────
  barTrack: {
    width: '100%',
    height: 4,
    borderRadius: 2,
    backgroundColor: calendarTokens.border,
    overflow: 'hidden',
    marginTop: 7,
  },
  barFill: {
    height: '100%',
    borderRadius: 2,
    backgroundColor: theme.palette.primary.main,
  },

  // ── Sticky footer ─────────────────────────────────────────────────────────
  stickyFooter: {
    position: 'fixed',
    bottom: 0,
    left: 0,
    right: 0,
    zIndex: 1200,
    height: 60,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(0, 3),
    backgroundColor: theme.palette.common.white,
    borderTop: `2px solid ${theme.palette.primary.main}`,
    boxShadow: '0 -4px 20px rgba(0,0,0,0.1)',
  },
  footerLeft: { display: 'flex', alignItems: 'center', gap: theme.spacing(1) },
  footerDot: {
    width: 8,
    height: 8,
    borderRadius: '50%',
    backgroundColor: theme.palette.primary.main,
    flexShrink: 0,
  },
  footerCount: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 600,
    color: theme.palette.text.primary,
  },
  footerRight: { display: 'flex', alignItems: 'center', gap: theme.spacing(1.5) },
  footerClearBtn: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.text.secondary,
    height: 36,
    padding: theme.spacing(0, 1.5),
    '&:hover': { backgroundColor: 'transparent', color: theme.palette.text.primary },
  },
  footerCloseBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 600,
    textTransform: 'none',
    height: 36,
    padding: theme.spacing(0, 2),
    borderRadius: 4,
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.75) },
  },
}));

const DAY_NAMES   = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const MONTH_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

export function WeekView({ weekStart, onBack, onPrevWeek, onNextWeek }: Props) {
  const classes = useStyles();
  const [compareMode, setCompareMode]   = useState('none');
  const [closedDays, setClosedDays]     = useState<Set<number>>(new Set());
  const [collapsed, setCollapsed]       = useState<Set<string>>(new Set());
  const [closeOutOpen, setCloseOutOpen] = useState(false);
  const [closeOutDates, setCloseOutDates] = useState<{ start: string; end: string } | undefined>();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const days: Date[] = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(weekStart);
    d.setDate(weekStart.getDate() + i);
    return d;
  });

  const weekLabel = `${days[0].toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${days[6].toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;

  const toggleClosed = (i: number) =>
    setClosedDays((prev) => { const n = new Set(prev); n.has(i) ? n.delete(i) : n.add(i); return n; });

  const toggleSection = (key: string) =>
    setCollapsed((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });

  const openCloseOut = () => {
    setCloseOutDates({
      start: days[0].toISOString().split('T')[0],
      end:   days[6].toISOString().split('T')[0],
    });
    setCloseOutOpen(true);
  };

  const showCompare = compareMode !== 'none';
  const closedCount = closedDays.size;

  return (
    <>
      <div className={classes.root}>
        {/* Header */}
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

            {/* ── Day header row ── */}
            <div className={classes.cornerCell} />
            {days.map((day, i) => {
              const isToday = day.getTime() === today.getTime();
              return (
                <div key={i} className={classes.dayHeader}>
                  <Typography className={classes.dayName}>{DAY_NAMES[i]}</Typography>
                  <Typography className={clsx(classes.dateNum, isToday && classes.dateNumToday)}>
                    {day.getDate()}
                  </Typography>
                  <Typography className={classes.dateMonth}>{MONTH_SHORT[day.getMonth()]}</Typography>
                </div>
              );
            })}

            {/* ── Close Out row ── */}
            <div className={classes.closeOutLabelCell} onClick={openCloseOut}>
              <Icon style={{ fontSize: 13, color: '#006461' }}>lock_outline</Icon>
              <Typography className={classes.closeOutLabelText}>Close Out</Typography>
            </div>
            {days.map((_, i) => (
              <div key={i} className={classes.closeOutCell} onClick={() => toggleClosed(i)}>
                <Icon className={closedDays.has(i) ? classes.lockIcon : classes.visIcon}>
                  {closedDays.has(i) ? 'lock' : 'visibility'}
                </Icon>
              </div>
            ))}

            {/* ── Metric sections ── */}
            {SECTIONS.map((section) => {
              const isCollapsed = collapsed.has(section.key);
              return (
                <React.Fragment key={section.key}>
                  <div className={classes.sectionHeader} onClick={() => toggleSection(section.key)}>
                    <div className={classes.sectionAccent} style={{ backgroundColor: section.accent }} />
                    <Typography className={classes.sectionTitle}>{section.label}</Typography>
                    <Icon className={clsx(classes.chevron, isCollapsed && classes.chevronCollapsed)}>
                      expand_more
                    </Icon>
                  </div>

                  {!isCollapsed && section.rows.map((row) => {
                    const isBar = !!row.showBar;
                    return (
                      <React.Fragment key={row.key}>
                        {/* Label */}
                        <div className={clsx(classes.labelCell, isBar && classes.labelCellBar, row.indent && classes.labelIndent)}>
                          <div className={classes.dot} style={{ backgroundColor: row.dot }} />
                          <Typography className={classes.labelText}>{row.label}</Typography>
                        </div>

                        {/* Day data cells */}
                        {days.map((day, i) => {
                          const dayIdx = day.getDate() * 7 + day.getMonth();
                          const val    = genValue(dayIdx, row.key);
                          const delta  = showCompare ? genDelta(dayIdx, row.key) : null;
                          const isPos  = delta !== null && delta > 0;
                          const isNeg  = delta !== null && delta < 0;
                          return (
                            <div key={i} className={clsx(classes.dataCell, isBar && classes.dataCellBar)}>
                              <Typography className={classes.value}>{row.format(val)}</Typography>
                              {delta !== null && (
                                <span className={clsx(
                                  classes.badge,
                                  isPos && classes.badgePos,
                                  isNeg && classes.badgeNeg,
                                  !isPos && !isNeg && classes.badgeNeutral,
                                )}>
                                  {isPos ? '▲' : isNeg ? '▼' : '─'}{Math.abs(delta).toFixed(1)}%
                                </span>
                              )}
                              {isBar && (
                                <div className={classes.barTrack}>
                                  <div className={classes.barFill} style={{ width: `${Math.min(val, 100)}%` }} />
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </React.Fragment>
              );
            })}

          </div>
        </div>
      </div>

      {/* Sticky footer — shown when days are closed out */}
      {closedCount > 0 && (
        <div className={classes.stickyFooter}>
          <div className={classes.footerLeft}>
            <div className={classes.footerDot} />
            <Typography className={classes.footerCount}>
              {closedCount} day{closedCount > 1 ? 's' : ''} closed
            </Typography>
          </div>
          <div className={classes.footerRight}>
            <Button className={classes.footerClearBtn} disableRipple onClick={() => setClosedDays(new Set())}>
              Clear
            </Button>
            <Button className={classes.footerCloseBtn} variant="contained" disableElevation onClick={openCloseOut}>
              <Icon>lock</Icon>
              Save Close Out
            </Button>
          </div>
        </div>
      )}

      <CloseOutModal
        open={closeOutOpen}
        onClose={() => setCloseOutOpen(false)}
        initialDates={closeOutDates}
      />
    </>
  );
}
