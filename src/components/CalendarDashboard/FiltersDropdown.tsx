import React, { useState, useEffect } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Popover from '@material-ui/core/Popover';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import FormControlLabel from '@material-ui/core/FormControlLabel';
import Radio from '@material-ui/core/Radio';
import RadioGroup from '@material-ui/core/RadioGroup';
import Divider from '@material-ui/core/Divider';
import { calendarTokens } from './tokens';

export type CalendarFilters = {
  operator: string;
  roomType: string;
  mealPlan: string;
  sourceGeo: string;
  pickup: number;
};

export const DEFAULT_FILTERS: CalendarFilters = {
  operator: 'all',
  roomType: 'all',
  mealPlan: 'all',
  sourceGeo: 'all',
  pickup: 7,
};

type Props = {
  anchorEl: HTMLElement | null;
  onClose: () => void;
  filters: CalendarFilters;
  onApply: (filters: CalendarFilters) => void;
};

const OPERATOR_OPTIONS = ['All', 'Sunwing', 'TUI Group', 'Thomas Cook', 'Club Med', 'Jet2holidays'];
const ROOM_TYPE_OPTIONS = ['All', 'Standard', 'Superior', 'Deluxe', 'Suite'];
const MEAL_PLAN_OPTIONS = ['All', 'All Inclusive', 'Half Board', 'B&B', 'Room Only'];
const SOURCE_GEO_OPTIONS = ['All', 'UK', 'Spain', 'US', 'Mexico'];
const PICKUP_OPTIONS = [1, 3, 7];

const useStyles = makeStyles((theme) => ({
  paper: {
    width: 280,
    padding: 0,
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  headerTitle: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  body: {
    maxHeight: 400,
    overflowY: 'auto',
    padding: theme.spacing(0, 2),
  },
  section: {
    paddingTop: theme.spacing(1.5),
    paddingBottom: theme.spacing(1),
  },
  sectionLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    fontWeight: 700,
    color: theme.palette.text.disabled,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: theme.spacing(0.5),
  },
  radioGroup: {
    gap: 0,
  },
  radioLabel: {
    margin: 0,
    '& .MuiFormControlLabel-label': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      color: theme.palette.text.primary,
    },
    '& .MuiRadio-root': {
      padding: theme.spacing(0.5, 1, 0.5, 0),
      color: calendarTokens.checkboxBorder,
    },
    '& .Mui-checked': {
      color: `${theme.palette.primary.main} !important`,
    },
  },
  pickupRow: {
    display: 'flex',
    gap: theme.spacing(1),
    marginTop: theme.spacing(0.5),
  },
  pickupBtn: {
    flex: 1,
    height: 30,
    borderRadius: 4,
    border: `1px solid ${calendarTokens.border}`,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    minWidth: 'auto',
    backgroundColor: theme.palette.common.white,
    color: theme.palette.text.secondary,
  },
  pickupBtnActive: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    borderColor: theme.palette.primary.main,
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
  footer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 2),
    borderTop: `1px solid ${calendarTokens.border}`,
  },
  resetBtn: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.text.secondary,
  },
  applyBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
}));

function FilterSection({ label, children }: { label: string; children: React.ReactNode }) {
  const classes = useStyles();
  return (
    <>
      <div className={classes.section}>
        <Typography className={classes.sectionLabel}>{label}</Typography>
        {children}
      </div>
      <Divider />
    </>
  );
}

export function FiltersDropdown({ anchorEl, onClose, filters, onApply }: Props) {
  const classes = useStyles();
  const [draft, setDraft] = useState<CalendarFilters>(filters);

  useEffect(() => {
    if (anchorEl) setDraft(filters);
  }, [anchorEl, filters]);

  const handleReset = () => setDraft(DEFAULT_FILTERS);

  const handleApply = () => {
    onApply(draft);
    onClose();
  };

  return (
    <Popover
      open={Boolean(anchorEl)}
      anchorEl={anchorEl}
      onClose={onClose}
      anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
      transformOrigin={{ vertical: 'top', horizontal: 'left' }}
      PaperProps={{ className: classes.paper }}
    >
      <div className={classes.header}>
        <Typography className={classes.headerTitle}>Filters</Typography>
      </div>

      <div className={classes.body}>
        <FilterSection label="Tour Operator">
          <RadioGroup
            value={draft.operator}
            onChange={(_, v) => setDraft((d) => ({ ...d, operator: v }))}
            className={classes.radioGroup}
          >
            {OPERATOR_OPTIONS.map((opt) => (
              <FormControlLabel
                key={opt}
                value={opt.toLowerCase() === 'all' ? 'all' : opt}
                control={<Radio size="small" />}
                label={opt}
                className={classes.radioLabel}
              />
            ))}
          </RadioGroup>
        </FilterSection>

        <FilterSection label="Room Type">
          <RadioGroup
            value={draft.roomType}
            onChange={(_, v) => setDraft((d) => ({ ...d, roomType: v }))}
            className={classes.radioGroup}
          >
            {ROOM_TYPE_OPTIONS.map((opt) => (
              <FormControlLabel
                key={opt}
                value={opt.toLowerCase() === 'all' ? 'all' : opt}
                control={<Radio size="small" />}
                label={opt}
                className={classes.radioLabel}
              />
            ))}
          </RadioGroup>
        </FilterSection>

        <FilterSection label="Meal Plan">
          <RadioGroup
            value={draft.mealPlan}
            onChange={(_, v) => setDraft((d) => ({ ...d, mealPlan: v }))}
            className={classes.radioGroup}
          >
            {MEAL_PLAN_OPTIONS.map((opt) => (
              <FormControlLabel
                key={opt}
                value={opt.toLowerCase() === 'all' ? 'all' : opt}
                control={<Radio size="small" />}
                label={opt}
                className={classes.radioLabel}
              />
            ))}
          </RadioGroup>
        </FilterSection>

        <FilterSection label="Source Geography">
          <RadioGroup
            value={draft.sourceGeo}
            onChange={(_, v) => setDraft((d) => ({ ...d, sourceGeo: v }))}
            className={classes.radioGroup}
          >
            {SOURCE_GEO_OPTIONS.map((opt) => (
              <FormControlLabel
                key={opt}
                value={opt.toLowerCase() === 'all' ? 'all' : opt}
                control={<Radio size="small" />}
                label={opt}
                className={classes.radioLabel}
              />
            ))}
          </RadioGroup>
        </FilterSection>

        <div className={classes.section}>
          <Typography className={classes.sectionLabel}>Customize Pickup</Typography>
          <div className={classes.pickupRow}>
            {PICKUP_OPTIONS.map((d) => (
              <Button
                key={d}
                className={`${classes.pickupBtn}${draft.pickup === d ? ` ${classes.pickupBtnActive}` : ''}`}
                onClick={() => setDraft((prev) => ({ ...prev, pickup: d }))}
                variant="outlined"
                disableElevation
              >
                {d}d
              </Button>
            ))}
          </div>
        </div>
      </div>

      <div className={classes.footer}>
        <Button className={classes.resetBtn} onClick={handleReset}>
          Reset All
        </Button>
        <Button className={classes.applyBtn} variant="contained" disableElevation onClick={handleApply}>
          Apply
        </Button>
      </div>
    </Popover>
  );
}
