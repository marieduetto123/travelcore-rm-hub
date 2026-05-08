import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import Typography from '@material-ui/core/Typography';
import { calendarTokens } from './tokens';
import { CellMetricsPopup } from './CellMetricsPopup';

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
  compareDisabled: {
    backgroundColor: calendarTokens.dropdownBackground,
    border: `1px solid ${calendarTokens.border}`,
    height: 36,
    minWidth: 120,
    padding: theme.spacing(0, 1.375),
    borderRadius: 4,
    display: 'flex',
    alignItems: 'center',
    boxShadow: '0px 1px 1.5px rgba(0,0,0,0.08)',
    opacity: 0.4,
    cursor: 'default',
  },
  compareLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.disabled,
    lineHeight: '16px',
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
}));

type Props = {
  title?: string;
  dateRangeLabel?: string;
  onCloseReopen?: () => void;
  onFilters?: () => void;
  onHeatmap?: () => void;
  onDateRange?: () => void;
};

export function CalendarHeader({
  title = 'Calendar',
  dateRangeLabel = 'January 2026 – February 2026',
  onCloseReopen,
  onFilters,
  onHeatmap,
  onDateRange,
}: Props) {
  const classes = useStyles();
  const [cellMetricsAnchor, setCellMetricsAnchor] = useState<HTMLElement | null>(null);

  return (
    <div className={classes.root}>
      <div className={classes.titleContainer}>
        <Typography className={classes.title}>{title}</Typography>
      </div>

      <div className={classes.controls}>
        <Button className={classes.primaryBtn} onClick={onCloseReopen} disableElevation>
          <Icon>lock</Icon>
          Close/Re-Open
        </Button>

        <Button className={classes.ghostBtn} onClick={onFilters}>
          <Icon>filter_list</Icon>
          Filters
          <Icon className={classes.expandIcon}>expand_more</Icon>
        </Button>

        <Button className={classes.ghostBtn} onClick={onHeatmap}>
          <Icon>grid_view</Icon>
          Heatmap
        </Button>

        <Button
          className={classes.ghostBtn}
          onClick={(e) => setCellMetricsAnchor(e.currentTarget)}
        >
          <Icon>tune</Icon>
          Cell Metrics
          <Icon className={classes.expandIcon}>expand_more</Icon>
        </Button>

        <div className={classes.compareDisabled}>
          <Typography className={classes.compareLabel}>No Compare</Typography>
        </div>

        <Button className={classes.dateRangeBtn} onClick={onDateRange}>
          {dateRangeLabel}
          <Icon>calendar_today</Icon>
        </Button>
      </div>

      <CellMetricsPopup
        anchorEl={cellMetricsAnchor}
        onClose={() => setCellMetricsAnchor(null)}
        onApply={() => setCellMetricsAnchor(null)}
      />
    </div>
  );
}
