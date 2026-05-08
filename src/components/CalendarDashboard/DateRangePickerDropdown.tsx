import React, { useState, useCallback } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Popover from '@material-ui/core/Popover';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';

export interface MonthRef { year: number; month: number }
export interface DateRange { start: MonthRef; end: MonthRef }

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function addMonths(ref: MonthRef, n: number): MonthRef {
  const total = ref.year * 12 + ref.month + n;
  const y = Math.floor(total / 12);
  const m = ((total % 12) + 12) % 12;
  return { year: y, month: m };
}

function cmpMonth(a: MonthRef, b: MonthRef): number {
  return (a.year * 12 + a.month) - (b.year * 12 + b.month);
}

function fmtMonth(ref: MonthRef): string {
  return `${MONTHS[ref.month]} ${ref.year}`;
}

const QUICK_GROUPS: { groupLabel: string | null; items: { label: string; getDates: (c: MonthRef) => DateRange }[] }[] = [
  {
    groupLabel: null,
    items: [
      { label: 'This Month',  getDates: (c) => ({ start: c, end: c }) },
      { label: 'Next Month',  getDates: (c) => { const n = addMonths(c, 1); return { start: n, end: n }; } },
    ],
  },
  {
    groupLabel: 'Current',
    items: [
      { label: 'Quarter', getDates: (c) => { const q = Math.floor(c.month / 3) * 3; const s = { year: c.year, month: q }; return { start: s, end: addMonths(s, 2) }; } },
      { label: 'Year',    getDates: (c) => ({ start: { year: c.year, month: 0 }, end: { year: c.year, month: 11 } }) },
    ],
  },
  {
    groupLabel: 'Next',
    items: [
      { label: '3 Months',  getDates: (c) => ({ start: c, end: addMonths(c, 2) }) },
      { label: '6 Months',  getDates: (c) => ({ start: c, end: addMonths(c, 5) }) },
      { label: '12 Months', getDates: (c) => ({ start: c, end: addMonths(c, 11) }) },
    ],
  },
];

const useStyles = makeStyles((theme) => ({
  paper: {
    width: 660,
    borderRadius: 8,
    overflow: 'hidden',
    boxShadow: '0 8px 24px rgba(0,0,0,0.14)',
    display: 'flex',
    flexDirection: 'column',
  },
  body: {
    display: 'flex',
  },
  calendars: {
    flex: 1,
    display: 'flex',
    padding: '20px 20px 16px',
    gap: 0,
  },
  panel: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
  panelDivider: {
    width: 1,
    backgroundColor: '#dde1e2',
    margin: '0 14px',
    alignSelf: 'stretch',
  },
  panelHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 10,
    height: 28,
  },
  navBtn: {
    width: 24,
    height: 24,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    borderRadius: 3,
    '&:hover': { backgroundColor: '#f5f5f5' },
    '& .MuiIcon-root': { fontSize: '18px !important', color: '#585858' },
  },
  navPlaceholder: {
    width: 24,
  },
  yearLabel: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 14,
    color: theme.palette.text.primary,
  },
  monthsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    rowGap: 2,
  },
  // Cell wrapper — clip the range bg so it doesn't bleed outside the cell
  monthCell: {
    position: 'relative',
    height: 34,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    userSelect: 'none',
    overflow: 'hidden',
  },
  // Full-width range background bar (z:0)
  rangeBg: {
    position: 'absolute',
    top: 3,
    bottom: 3,
    zIndex: 0,
    backgroundColor: 'rgba(0,100,97,0.13)',
  },
  // Circular selected cell (z:1)
  circle: {
    position: 'absolute',
    width: 42,
    height: 28,
    borderRadius: 14,
    zIndex: 1,
  },
  circleSelected: {
    backgroundColor: theme.palette.primary.main,
  },
  circleHover: {
    backgroundColor: 'rgba(0,100,97,0.15)',
  },
  // Month text (z:2)
  monthText: {
    position: 'relative',
    zIndex: 2,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 400,
    color: theme.palette.text.primary,
    lineHeight: 1,
  },
  monthTextSelected: {
    color: theme.palette.common.white,
    fontWeight: 600,
  },

  // ── Sidebar ──
  sidebar: {
    width: 106,
    borderLeft: `1px solid #dde1e2`,
    padding: '16px 10px',
    display: 'flex',
    flexDirection: 'column',
    overflowY: 'auto',
  },
  sidebarDivider: {
    height: 1,
    backgroundColor: '#dde1e2',
    margin: '6px 0',
  },
  sidebarGroupLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 10,
    fontWeight: 700,
    color: '#9ca3af',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    padding: '4px 4px 2px',
  },
  sidebarItem: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.primary.main,
    padding: '5px 8px',
    borderRadius: 4,
    cursor: 'pointer',
    '&:hover': { backgroundColor: 'rgba(0,100,97,0.08)' },
  },
  sidebarItemActive: {
    fontWeight: 700,
    backgroundColor: 'rgba(0,100,97,0.08)',
  },

  // ── Footer ──
  footer: {
    borderTop: `1px solid #dde1e2`,
    height: 52,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '0 20px',
    flexShrink: 0,
  },
  rangeText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
  },
  footerBtns: {
    display: 'flex',
    gap: 8,
  },
  cancelBtn: {
    height: 32,
    padding: '0 16px',
    borderRadius: 4,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.text.secondary,
    border: '1px solid #dde1e2',
    minWidth: 'auto',
  },
  applyBtn: {
    height: 32,
    padding: '0 16px',
    borderRadius: 4,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    minWidth: 'auto',
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '&:disabled': { opacity: 0.5, color: theme.palette.common.white },
  },
}));

type Props = {
  anchorEl: HTMLElement | null;
  value: DateRange;
  onApply: (range: DateRange) => void;
  onClose: () => void;
};

export function DateRangePickerDropdown({ anchorEl, value, onApply, onClose }: Props) {
  const classes = useStyles();
  const today: MonthRef = { year: new Date().getFullYear(), month: new Date().getMonth() };

  const [leftYear, setLeftYear]   = useState(() => value.start.year);
  const [rightYear, setRightYear] = useState(() =>
    value.end.year > value.start.year ? value.end.year : value.start.year + 1
  );
  const [pendingStart, setPendingStart] = useState<MonthRef | null>(value.start);
  const [pendingEnd, setPendingEnd]     = useState<MonthRef | null>(value.end);
  const [hoverMonth, setHoverMonth]     = useState<MonthRef | null>(null);
  const [activeQuick, setActiveQuick]   = useState<string | null>(null);

  const shiftYears = (dir: number) => {
    setLeftYear(y => y + dir);
    setRightYear(y => y + dir);
  };

  const handleMonthClick = useCallback((ref: MonthRef) => {
    setActiveQuick(null);
    if (!pendingStart || (pendingStart && pendingEnd)) {
      setPendingStart(ref);
      setPendingEnd(null);
    } else {
      if (cmpMonth(ref, pendingStart) === 0) {
        setPendingStart(null);
      } else if (cmpMonth(ref, pendingStart) < 0) {
        setPendingEnd(pendingStart);
        setPendingStart(ref);
      } else {
        setPendingEnd(ref);
      }
    }
  }, [pendingStart, pendingEnd]);

  const getCellState = useCallback((year: number, month: number) => {
    const val = year * 12 + month;
    const startVal = pendingStart ? pendingStart.year * 12 + pendingStart.month : null;
    const effectiveEnd = pendingEnd ?? (pendingStart && hoverMonth ? hoverMonth : null);
    const endVal = effectiveEnd ? effectiveEnd.year * 12 + effectiveEnd.month : null;

    const isStart = startVal !== null && val === startVal;
    const isEnd   = endVal   !== null && val === endVal;
    const hasRange = startVal !== null && endVal !== null && startVal !== endVal;

    const lo = hasRange ? Math.min(startVal!, endVal!) : null;
    const hi = hasRange ? Math.max(startVal!, endVal!) : null;
    const isInRange = lo !== null && hi !== null && val > lo && val < hi;

    // Hover preview (when start set, no end, hovering)
    const isHoverPreview = !pendingEnd && pendingStart && hoverMonth
      ? (() => {
          const sv = startVal!;
          const hv = hoverMonth.year * 12 + hoverMonth.month;
          return val > Math.min(sv, hv) && val < Math.max(sv, hv);
        })()
      : false;

    return { isStart, isEnd, hasRange, isInRange: isInRange || isHoverPreview };
  }, [pendingStart, pendingEnd, hoverMonth]);

  const handleQuickSelect = (label: string, getDates: (c: MonthRef) => DateRange) => {
    const range = getDates(today);
    setPendingStart(range.start);
    setPendingEnd(range.end);
    setActiveQuick(label);
    setLeftYear(range.start.year);
    setRightYear(range.end.year > range.start.year ? range.end.year : range.start.year + 1);
  };

  const rangeLabel = (() => {
    if (pendingStart && pendingEnd) return `${fmtMonth(pendingStart)} – ${fmtMonth(pendingEnd)}`;
    if (pendingStart) return `${fmtMonth(pendingStart)} – …`;
    return '–';
  })();

  const renderPanel = (year: number, isLeft: boolean) => (
    <div className={classes.panel}>
      <div className={classes.panelHeader}>
        {isLeft ? (
          <>
            <div className={classes.navBtn} onClick={() => shiftYears(-1)}>
              <Icon>chevron_left</Icon>
            </div>
            <span className={classes.yearLabel}>{year}</span>
            <div className={classes.navPlaceholder} />
          </>
        ) : (
          <>
            <div className={classes.navPlaceholder} />
            <span className={classes.yearLabel}>{year}</span>
            <div className={classes.navBtn} onClick={() => shiftYears(1)}>
              <Icon>chevron_right</Icon>
            </div>
          </>
        )}
      </div>

      <div className={classes.monthsGrid}>
        {MONTHS.map((name, mi) => {
          const { isStart, isEnd, hasRange, isInRange } = getCellState(year, mi);
          const isSelected = isStart || isEnd;

          const isHovering = !pendingEnd && hoverMonth?.year === year && hoverMonth?.month === mi;

          return (
            <div
              key={mi}
              className={classes.monthCell}
              onClick={() => handleMonthClick({ year, month: mi })}
              onMouseEnter={() => { if (pendingStart && !pendingEnd) setHoverMonth({ year, month: mi }); }}
              onMouseLeave={() => setHoverMonth(null)}
            >
              {/* Range background bar */}
              {(isInRange || (isStart && hasRange) || (isEnd && hasRange)) && (
                <div
                  className={classes.rangeBg}
                  style={{
                    left:  isStart ? '50%' : 0,
                    right: isEnd   ? '50%' : 0,
                  }}
                />
              )}

              {/* Selected circle */}
              {isSelected && <div className={`${classes.circle} ${classes.circleSelected}`} />}

              {/* Hover indicator (non-selected) */}
              {!isSelected && isHovering && <div className={`${classes.circle} ${classes.circleHover}`} />}

              <span className={`${classes.monthText}${isSelected ? ` ${classes.monthTextSelected}` : ''}`}>
                {name}
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );

  return (
    <Popover
      open={Boolean(anchorEl)}
      anchorEl={anchorEl}
      onClose={onClose}
      anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}
      transformOrigin={{ vertical: 'top', horizontal: 'right' }}
      PaperProps={{ className: classes.paper }}
      style={{ marginTop: 4 }}
    >
      <div className={classes.body}>
        <div className={classes.calendars}>
          {renderPanel(leftYear, true)}
          <div className={classes.panelDivider} />
          {renderPanel(rightYear, false)}
        </div>

        {/* Quick-select sidebar */}
        <div className={classes.sidebar}>
          {QUICK_GROUPS.map((group, gi) => (
            <React.Fragment key={gi}>
              {gi > 0 && <div className={classes.sidebarDivider} />}
              {group.groupLabel && (
                <div className={classes.sidebarGroupLabel}>{group.groupLabel}</div>
              )}
              {group.items.map((item) => (
                <div
                  key={item.label}
                  className={`${classes.sidebarItem}${activeQuick === item.label ? ` ${classes.sidebarItemActive}` : ''}`}
                  onClick={() => handleQuickSelect(item.label, item.getDates)}
                >
                  {item.label}
                </div>
              ))}
            </React.Fragment>
          ))}
        </div>
      </div>

      {/* Footer */}
      <div className={classes.footer}>
        <span className={classes.rangeText}>{rangeLabel}</span>
        <div className={classes.footerBtns}>
          <Button className={classes.cancelBtn} onClick={onClose}>Cancel</Button>
          <Button
            className={classes.applyBtn}
            disableElevation
            disabled={!pendingStart || !pendingEnd}
            onClick={() => {
              if (pendingStart && pendingEnd) {
                onApply({ start: pendingStart, end: pendingEnd });
                onClose();
              }
            }}
          >
            Apply
          </Button>
        </div>
      </div>
    </Popover>
  );
}
