import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Popover from '@material-ui/core/Popover';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { calendarTokens } from './tokens';

// ─── Design tokens not in standard MUI palette ───────────────────────────────
const popupTokens = {
  rowBg: '#f8f9fa',
  rowBorder: '#e5e7eb',
  headerBorder: '#f3f4f6',
  sectionBorder: '#eaeeef',
  checkboxUncheckedBorder: '#4f5b60',
  checkboxDisabledBg: '#eaeeef',
  checkboxDisabledBorder: '#aeb4ba',
  unavailableBoxBg: '#e5e7eb',
  unavailableBoxBorder: '#dde1e2',
  infoBlue: '#1b4dc0',
  infoBlueBorder: '#00298c',
  errorBg: '#ffebee',
  errorText: '#991f1f',
  segmentCyan: '#0891b2',
  segmentViolet: '#7c3aed',
  segmentAmber: '#f59e0b',
} as const;

// ─── Types ────────────────────────────────────────────────────────────────────
type CellType = 'interactive' | 'checkboxDisabled' | 'unavailable' | 'absent';

type MetricRow = {
  id: string;
  label: string;
  hotel: CellType;
  operator: CellType;
};

type MetricGroup = {
  id: string;
  name: string;
  rows: MetricRow[];
};

type SegmentItem = {
  id: string;
  label: string;
  color?: string;
  bold?: boolean;
};

// ─── Static data ─────────────────────────────────────────────────────────────
const SEGMENTS: SegmentItem[] = [
  { id: 'all', label: 'All', bold: true },
  { id: 'static_fit', label: 'Static FIT Rates', color: popupTokens.segmentCyan },
  { id: 'operator_dynamic', label: 'Operator Dynamic', color: popupTokens.segmentViolet },
  { id: 'tour_series', label: 'Tour Series', color: popupTokens.segmentAmber },
];

const METRIC_GROUPS: MetricGroup[] = [
  {
    id: 'occupancy', name: 'Occupancy',
    rows: [
      { id: 'occ_act_1', label: 'Actual', hotel: 'interactive', operator: 'checkboxDisabled' },
      { id: 'occ_act_2', label: 'Actual', hotel: 'interactive', operator: 'checkboxDisabled' },
      { id: 'occ_act_3', label: 'Actual', hotel: 'interactive', operator: 'checkboxDisabled' },
      { id: 'occ_fcst', label: 'Fcst', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'adr', name: 'ADR',
    rows: [
      { id: 'adr_act_1', label: 'Actual', hotel: 'interactive', operator: 'checkboxDisabled' },
      { id: 'adr_act_2', label: 'Actual', hotel: 'interactive', operator: 'checkboxDisabled' },
      { id: 'adr_act_3', label: 'Actual', hotel: 'checkboxDisabled', operator: 'checkboxDisabled' },
      { id: 'adr_act_4', label: 'Actual', hotel: 'checkboxDisabled', operator: 'checkboxDisabled' },
      { id: 'adr_act_5', label: 'Actual', hotel: 'checkboxDisabled', operator: 'checkboxDisabled' },
      { id: 'adr_fcst', label: 'Fcst', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'revenue', name: 'Revenue',
    rows: [
      { id: 'rev_act', label: 'Actual', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rev_ly', label: 'LY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rev_stly', label: 'STLY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rev_fcst', label: 'Fcst', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'rn_sold', name: 'RN Sold',
    rows: [
      { id: 'rn_act', label: 'Actual', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rn_ly', label: 'LY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rn_stly', label: 'STLY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rn_fcst', label: 'Fcst', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'revpar', name: 'RevPAR',
    rows: [
      { id: 'rp_act', label: 'Actual', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rp_ly', label: 'LY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rp_stly', label: 'STLY', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'rp_fcst', label: 'Fcst', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'other', name: 'Other Metrics',
    rows: [
      { id: 'oth_pickup', label: 'Pickup', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'oth_los', label: 'Avg LOS', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'oth_lead', label: 'ALT', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'oth_adults', label: 'AD', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'oth_children', label: 'CHD', hotel: 'unavailable', operator: 'unavailable' },
      { id: 'oth_guests', label: 'PAX', hotel: 'unavailable', operator: 'unavailable' },
    ],
  },
  {
    id: 'availability', name: 'Availability',
    rows: [
      { id: 'av_rooms', label: 'AR', hotel: 'interactive', operator: 'absent' },
      { id: 'av_guar', label: 'Avail Guar.', hotel: 'absent', operator: 'checkboxDisabled' },
    ],
  },
  {
    id: 'biz_mix', name: 'Business Mix',
    rows: [
      { id: 'bm_to', label: 'TO Mix %', hotel: 'unavailable', operator: 'absent' },
      { id: 'bm_direct', label: 'Direct Mix %', hotel: 'unavailable', operator: 'absent' },
      { id: 'bm_ota', label: 'OTA Mix %', hotel: 'unavailable', operator: 'absent' },
    ],
  },
  {
    id: 'selling_rates', name: 'Selling Rates',
    rows: [
      { id: 'sr_contract', label: 'TO Contract Rate', hotel: 'unavailable', operator: 'absent' },
      { id: 'sr_promo', label: 'Promotion %', hotel: 'unavailable', operator: 'absent' },
      { id: 'sr_base', label: 'Base Segment Rate', hotel: 'unavailable', operator: 'absent' },
    ],
  },
];

const DEFAULT_SELECTIONS: Record<string, boolean> = {};

// ─── Styles ───────────────────────────────────────────────────────────────────
const useStyles = makeStyles((theme) => ({
  paper: {
    width: 370,
    height: 564,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    borderRadius: 8,
    border: `1px solid ${popupTokens.rowBorder}`,
    boxShadow: '0px 8px 24px rgba(0,0,0,0.15)',
  },

  // Header
  header: {
    flexShrink: 0,
    borderBottom: `1px solid ${popupTokens.headerBorder}`,
    padding: '8px 12px 7px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  headerTitle: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 10,
    color: calendarTokens.checkboxBorder,
    textTransform: 'uppercase',
    letterSpacing: 0.4,
    lineHeight: 'normal',
  },
  headerCounter: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 10,
    color: theme.palette.error.main,
    lineHeight: 'normal',
  },

  // Segments area
  segmentsArea: {
    flexShrink: 0,
    padding: theme.spacing(1, 1.5),
    paddingBottom: theme.spacing(0),
  },
  segmentsLabel: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 9,
    color: theme.palette.text.disabled,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    lineHeight: 'normal',
    marginBottom: 0,
  },
  segmentRow: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    paddingTop: 3,
    paddingBottom: 3,
    '&:first-child': {
      paddingTop: 9,
    },
  },
  segmentSwatch: {
    width: 10,
    height: 10,
    borderRadius: 2,
    flexShrink: 0,
  },
  segmentLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    lineHeight: 'normal',
    color: theme.palette.text.secondary,
  },
  segmentLabelBold: {
    fontWeight: 700,
    color: theme.palette.text.primary,
  },

  // Info snackbar (always shown, blue)
  infoSnackbar: {
    margin: theme.spacing(1, 1.5),
    backgroundColor: theme.palette.common.white,
    border: `1px solid ${popupTokens.infoBlueBorder}`,
    borderRadius: 4,
    display: 'flex',
    alignItems: 'flex-start',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 1),
  },
  infoIcon: {
    fontSize: '24px !important',
    color: popupTokens.infoBlue,
    flexShrink: 0,
    lineHeight: '24px',
  },
  infoText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    color: popupTokens.infoBlue,
    lineHeight: 'normal',
    alignSelf: 'center',
  },

  // Table divider
  tableDivider: {
    flexShrink: 0,
    borderTop: `1px solid ${popupTokens.sectionBorder}`,
    margin: theme.spacing(0, 1.5, 0, 1.5),
  },

  // Metrics table
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    minHeight: 0,
  },

  // Table header row
  tableHeaderRow: {
    display: 'flex',
    justifyContent: 'flex-end',
  },
  tableHeaderCell: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 9,
    color: calendarTokens.checkboxBorder,
    textTransform: 'uppercase',
    letterSpacing: 0.4,
    textAlign: 'center',
    padding: '5px 8px',
  },
  tableHeaderHotel: { width: 123 },
  tableHeaderOperator: { width: 103 },

  // Group header row
  groupRow: {
    backgroundColor: popupTokens.rowBg,
    borderTop: `1px solid ${popupTokens.rowBorder}`,
    borderBottom: `1px solid ${popupTokens.rowBorder}`,
    height: 31,
    display: 'flex',
    alignItems: 'center',
    paddingLeft: 12,
    paddingRight: 12,
    paddingTop: 5.5,
    paddingBottom: 2.5,
    boxSizing: 'border-box',
  },
  groupName: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 14,
    color: theme.palette.text.primary,
    lineHeight: 'normal',
    whiteSpace: 'nowrap',
  },

  // Data row
  dataRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '100%',
  },
  dataLabelCell: {
    width: 70,
    paddingLeft: 12,
    paddingRight: 8,
    paddingTop: 7,
    paddingBottom: 6.5,
    flexShrink: 0,
  },
  dataLabel: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 400,
    fontSize: 11,
    color: theme.palette.text.secondary,
    lineHeight: 'normal',
    whiteSpace: 'nowrap',
  },
  dataCheckCell: {
    flex: 1,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    paddingTop: 7,
    paddingBottom: 4,
    paddingLeft: 4,
    paddingRight: 4,
    minWidth: 0,
  },

  // Custom checkbox states
  checkbox: {
    width: 18,
    height: 18,
    borderRadius: 2,
    boxSizing: 'border-box',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
    transition: theme.transitions.create(['background-color', 'border-color']),
  },
  checkboxChecked: {
    backgroundColor: theme.palette.primary.main,
    border: `1px solid ${theme.palette.primary.main}`,
    '& $checkboxIcon': { opacity: 1 },
  },
  checkboxUnchecked: {
    backgroundColor: theme.palette.common.white,
    border: `2px solid ${popupTokens.checkboxUncheckedBorder}`,
  },
  checkboxDisabledState: {
    backgroundColor: popupTokens.checkboxDisabledBg,
    border: `2px solid ${popupTokens.checkboxDisabledBorder}`,
    cursor: 'default',
  },
  checkboxIcon: {
    fontSize: '14px !important',
    color: theme.palette.common.white,
    opacity: 0,
  },

  // Unavailable slot (16×16 grey box, not a checkbox)
  unavailableBox: {
    width: 16,
    height: 16,
    borderRadius: 2,
    backgroundColor: popupTokens.unavailableBoxBg,
    border: `1px solid ${popupTokens.unavailableBoxBorder}`,
    flexShrink: 0,
    boxSizing: 'border-box',
  },

  // Footer
  footer: {
    flexShrink: 0,
    borderTop: `1px solid ${popupTokens.rowBorder}`,
    height: 50.5,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
    padding: '11px 12px 10px',
    boxSizing: 'border-box',
    position: 'relative',
  },
  resetBtn: {
    height: 32,
    padding: theme.spacing(0, 1.5),
    borderRadius: 4,
    color: theme.palette.primary.main,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 400,
    textTransform: 'none',
    minWidth: 'auto',
  },
  applyBtn: {
    height: 32,
    padding: theme.spacing(0, 1.5),
    borderRadius: 4,
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 400,
    textTransform: 'none',
    minWidth: 'auto',
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '&:disabled': { backgroundColor: calendarTokens.checkboxBorder, color: theme.palette.common.white },
  },

  // Error snackbar (floating above footer when >4 selected)
  errorSnackbar: {
    position: 'absolute',
    bottom: 50.5 + 8,
    left: 15,
    width: 289,
    backgroundColor: popupTokens.errorBg,
    borderRadius: 4,
    boxShadow: '0px 1px 0.5px rgba(0,0,0,0.14), 0px 2px 0.5px rgba(0,0,0,0.12), 0px 1px 1.5px rgba(0,0,0,0.2)',
    display: 'flex',
    alignItems: 'flex-start',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 1),
  },
  errorIcon: {
    fontSize: '24px !important',
    color: popupTokens.errorText,
    flexShrink: 0,
    lineHeight: '24px',
  },
  errorText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    color: popupTokens.errorText,
    lineHeight: 'normal',
    flex: 1,
    alignSelf: 'center',
  },
  errorClose: {
    fontSize: '24px !important',
    color: popupTokens.errorText,
    cursor: 'pointer',
    flexShrink: 0,
    lineHeight: '24px',
  },
}));

// ─── Sub-components ───────────────────────────────────────────────────────────
type CheckboxCellProps = {
  type: CellType;
  checked?: boolean;
  onToggle?: () => void;
  isAtMax?: boolean;
  classes: ReturnType<typeof useStyles>;
};

function CheckboxCell({ type, checked, onToggle, isAtMax, classes }: CheckboxCellProps) {
  if (type === 'absent') return <div style={{ flex: 1, minWidth: 0 }} />;

  // All non-absent types render as selectable checkboxes
  const showDisabled = isAtMax && !checked;

  return (
    <div className={classes.dataCheckCell}>
      <div
        className={clsx(
          classes.checkbox,
          checked
            ? classes.checkboxChecked
            : showDisabled
            ? classes.checkboxDisabledState
            : classes.checkboxUnchecked,
        )}
        onClick={onToggle}
      >
        <Icon className={classes.checkboxIcon}>check</Icon>
      </div>
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────
type Props = {
  anchorEl: HTMLElement | null;
  onClose: () => void;
  onApply?: (selections: Record<string, boolean>) => void;
};

const MAX_SELECTIONS = 4;

export function CellMetricsPopup({ anchorEl, onClose, onApply }: Props) {
  const classes = useStyles();
  const [selections, setSelections] = useState<Record<string, boolean>>(DEFAULT_SELECTIONS);
  const [showError, setShowError] = useState(false);

  const selectedCount = Object.values(selections).filter(Boolean).length;

  const handleToggle = (id: string) => {
    const isChecked = selections[id];
    if (!isChecked && selectedCount >= MAX_SELECTIONS) {
      setShowError(true);
      return;
    }
    setSelections((prev) => ({ ...prev, [id]: !isChecked }));
    setShowError(false);
  };

  const handleReset = () => {
    setSelections({});
    setShowError(false);
  };

  const handleApply = () => {
    if (selectedCount <= MAX_SELECTIONS) {
      onApply?.(selections);
      onClose();
    }
  };

  return (
    <Popover
      open={Boolean(anchorEl)}
      anchorEl={anchorEl}
      onClose={onClose}
      anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
      transformOrigin={{ vertical: 'top', horizontal: 'left' }}
      PaperProps={{ className: classes.paper, elevation: 0 }}
      marginThreshold={8}
    >
      {/* Header */}
      <div className={classes.header}>
        <Typography className={classes.headerTitle}>Cell Metrics</Typography>
        <Typography className={classes.headerCounter}>
          {selectedCount} / {MAX_SELECTIONS} rows
        </Typography>
      </div>

      {/* Segments to show */}
      <div className={classes.segmentsArea}>
        <Typography className={classes.segmentsLabel}>Segments to show</Typography>
        {SEGMENTS.map((seg) => (
          <div key={seg.id} className={classes.segmentRow}>
            <div className={clsx(classes.checkbox, classes.checkboxChecked)}>
              <Icon className={classes.checkboxIcon}>check</Icon>
            </div>
            {seg.color && (
              <div className={classes.segmentSwatch} style={{ backgroundColor: seg.color }} />
            )}
            <span className={clsx(classes.segmentLabel, seg.bold && classes.segmentLabelBold)}>
              {seg.label}
            </span>
          </div>
        ))}
      </div>

      {/* Info snackbar */}
      <div className={classes.infoSnackbar}>
        <Icon className={classes.infoIcon}>info</Icon>
        <Typography className={classes.infoText}>Only select up to 4 metrics</Typography>
      </div>

      {/* Table divider */}
      <div className={classes.tableDivider} />

      {/* Metrics table */}
      <div className={classes.tableWrapper}>
        {/* Column headers */}
        <div className={classes.tableHeaderRow}>
          <Typography className={clsx(classes.tableHeaderCell, classes.tableHeaderHotel)}>
            Hotel
          </Typography>
          <Typography className={clsx(classes.tableHeaderCell, classes.tableHeaderOperator)}>
            Operator
          </Typography>
        </div>

        {/* Groups */}
        {METRIC_GROUPS.map((group) => (
          <React.Fragment key={group.id}>
            <div className={classes.groupRow}>
              <Typography className={classes.groupName}>{group.name}</Typography>
            </div>
            {group.rows.map((row) => (
              <div key={row.id} className={classes.dataRow}>
                <div className={classes.dataLabelCell}>
                  <Typography className={classes.dataLabel}>{row.label}</Typography>
                </div>
                <CheckboxCell
                  type={row.hotel}
                  checked={!!selections[`${row.id}_hotel`]}
                  onToggle={() => handleToggle(`${row.id}_hotel`)}
                  isAtMax={selectedCount >= MAX_SELECTIONS}
                  classes={classes}
                />
                <CheckboxCell
                  type={row.operator}
                  checked={!!selections[`${row.id}_op`]}
                  onToggle={() => handleToggle(`${row.id}_op`)}
                  isAtMax={selectedCount >= MAX_SELECTIONS}
                  classes={classes}
                />
              </div>
            ))}
          </React.Fragment>
        ))}
      </div>

      {/* Footer */}
      <div className={classes.footer}>
        {/* Error snackbar (floating above footer) */}
        {showError && (
          <div className={classes.errorSnackbar}>
            <Icon className={classes.errorIcon}>error_outline</Icon>
            <Typography className={classes.errorText}>Only select up to 4 metrics</Typography>
            <Icon className={classes.errorClose} onClick={() => setShowError(false)}>close</Icon>
          </div>
        )}

        <Button className={classes.resetBtn} onClick={handleReset}>
          Reset
        </Button>
        <Button
          className={classes.applyBtn}
          onClick={handleApply}
          disabled={selectedCount > MAX_SELECTIONS}
          disableElevation
        >
          Apply
        </Button>
      </div>
    </Popover>
  );
}
