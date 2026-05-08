import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import Typography from '@material-ui/core/Typography';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import Badge from '@material-ui/core/Badge';
import { calendarTokens } from './tokens';
import { CellMetricsPopup } from './CellMetricsPopup';
import { FiltersDropdown, CalendarFilters, DEFAULT_FILTERS } from './FiltersDropdown';
import { HeatmapModal } from './HeatmapModal';
import { CloseOutModal } from './CloseOutModal';

const COMPARE_OPTIONS = [
  { value: 'none', label: 'No Compare' },
  { value: 'ly', label: 'vs LY' },
  { value: 'stly', label: 'vs STLY' },
  { value: 'forecast', label: 'vs Forecast' },
  { value: 'budget', label: 'vs Budget' },
];

const useStyles = makeStyles((theme) => ({
  root: {
    backgroundColor: theme.palette.common.white,
    borderBottom: `1px solid ${calendarTokens.border}`,
    boxShadow: calendarTokens.headerShadow,
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    paddingTop: theme.spacing(1.5),
    paddingBottom: theme.spacing(1.875),
    paddingLeft: theme.spacing(2),
    paddingRight: theme.spacing(2),
    minHeight: 63,
    boxSizing: 'border-box',
    flexWrap: 'wrap',
  },
  titleContainer: {
    flex: '1 0 0',
    minWidth: 0,
  },
  title: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 16,
    fontWeight: 400,
    color: theme.palette.text.secondary,
    lineHeight: 'normal',
  },
  controls: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: theme.spacing(1),
    justifyContent: 'flex-end',
  },
  primaryBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    height: 36,
    padding: theme.spacing(0, 2),
    borderRadius: 4,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 400,
    textTransform: 'none',
    minWidth: 'auto',
    '&:hover': {
      backgroundColor: theme.palette.primary.dark,
    },
    '& .MuiIcon-root': {
      fontSize: '16px !important',
      marginRight: theme.spacing(0.5),
    },
  },
  ghostBtn: {
    height: 36,
    padding: theme.spacing(0, 2),
    borderRadius: 4,
    color: theme.palette.text.secondary,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 400,
    textTransform: 'none',
    minWidth: 'auto',
    '& .MuiIcon-root': {
      fontSize: '16px !important',
    },
    '& .MuiButton-label': {
      gap: theme.spacing(0.75),
    },
  },
  ghostBtnActive: {
    color: theme.palette.primary.main,
    backgroundColor: calendarTokens.primaryHover,
  },
  compareSelect: {
    height: 36,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    '& .MuiSelect-select': {
      padding: theme.spacing(0.5, 3.5, 0.5, 1.25),
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
    },
    '& .MuiOutlinedInput-notchedOutline': {
      borderColor: calendarTokens.border,
    },
  },
  dateRangeBtn: {
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${calendarTokens.border}`,
    height: 36,
    padding: theme.spacing(0, 1.875),
    borderRadius: 4,
    color: theme.palette.text.primary,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 400,
    textTransform: 'none',
    minWidth: 'auto',
    '&:hover': {
      backgroundColor: calendarTokens.cellBackground,
    },
    '& .MuiButton-label': {
      gap: theme.spacing(0.75),
    },
    '& .MuiIcon-root': {
      fontSize: '16px !important',
    },
  },
  expandIcon: {
    fontSize: '14px !important',
  },
  filtersBadge: {
    '& .MuiBadge-badge': {
      backgroundColor: theme.palette.primary.main,
      color: theme.palette.common.white,
      fontSize: 10,
      height: 16,
      minWidth: 16,
      top: 4,
      right: 4,
    },
  },
}));

type Props = {
  title?: string;
  dateRangeLabel?: string;
  onDateRange?: () => void;
};

function countActiveFilters(f: CalendarFilters): number {
  let n = 0;
  if (f.operator !== 'all') n++;
  if (f.roomType !== 'all') n++;
  if (f.mealPlan !== 'all') n++;
  if (f.sourceGeo !== 'all') n++;
  if (f.pickup !== 7) n++;
  return n;
}

export function CalendarHeader({
  title = 'Calendar',
  dateRangeLabel = 'January 2026 – February 2026',
  onDateRange,
}: Props) {
  const classes = useStyles();

  // CellMetrics
  const [cellMetricsAnchor, setCellMetricsAnchor] = useState<HTMLElement | null>(null);

  // Filters
  const [filtersAnchor, setFiltersAnchor] = useState<HTMLElement | null>(null);
  const [filters, setFilters] = useState<CalendarFilters>(DEFAULT_FILTERS);

  // Heatmap
  const [heatmapOpen, setHeatmapOpen] = useState(false);

  // Close/Re-Open
  const [closeOutOpen, setCloseOutOpen] = useState(false);

  // Compare
  const [compareMode, setCompareMode] = useState('none');

  const activeFilterCount = countActiveFilters(filters);

  return (
    <div className={classes.root}>
      <div className={classes.titleContainer}>
        <Typography className={classes.title}>{title}</Typography>
      </div>

      <div className={classes.controls}>
        {/* Close/Re-Open */}
        <Button
          className={classes.primaryBtn}
          onClick={() => setCloseOutOpen(true)}
          disableElevation
        >
          <Icon>lock</Icon>
          Close/Re-Open
        </Button>

        {/* Filters */}
        <Badge
          badgeContent={activeFilterCount || undefined}
          className={classes.filtersBadge}
        >
          <Button
            className={classes.ghostBtn}
            onClick={(e) => setFiltersAnchor(e.currentTarget)}
          >
            <Icon>filter_list</Icon>
            Filters
            <Icon className={classes.expandIcon}>expand_more</Icon>
          </Button>
        </Badge>

        {/* Heatmap */}
        <Button
          className={`${classes.ghostBtn}${heatmapOpen ? ` ${classes.ghostBtnActive}` : ''}`}
          onClick={() => setHeatmapOpen(true)}
        >
          <Icon>grid_view</Icon>
          Heatmap
        </Button>

        {/* Cell Metrics */}
        <Button
          className={classes.ghostBtn}
          onClick={(e) => setCellMetricsAnchor(e.currentTarget)}
        >
          <Icon>tune</Icon>
          Cell Metrics
          <Icon className={classes.expandIcon}>expand_more</Icon>
        </Button>

        {/* Compare */}
        <Select
          value={compareMode}
          onChange={(e) => setCompareMode(e.target.value as string)}
          variant="outlined"
          className={classes.compareSelect}
        >
          {COMPARE_OPTIONS.map((o) => (
            <MenuItem
              key={o.value}
              value={o.value}
              style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}
            >
              {o.label}
            </MenuItem>
          ))}
        </Select>

        {/* Date Range */}
        <Button className={classes.dateRangeBtn} onClick={onDateRange}>
          {dateRangeLabel}
          <Icon>calendar_today</Icon>
        </Button>
      </div>

      {/* Popups */}
      <CellMetricsPopup
        anchorEl={cellMetricsAnchor}
        onClose={() => setCellMetricsAnchor(null)}
        onApply={() => setCellMetricsAnchor(null)}
      />

      <FiltersDropdown
        anchorEl={filtersAnchor}
        onClose={() => setFiltersAnchor(null)}
        filters={filters}
        onApply={(f) => setFilters(f)}
      />

      <HeatmapModal
        open={heatmapOpen}
        onClose={() => setHeatmapOpen(false)}
      />

      <CloseOutModal
        open={closeOutOpen}
        onClose={() => setCloseOutOpen(false)}
      />
    </div>
  );
}
