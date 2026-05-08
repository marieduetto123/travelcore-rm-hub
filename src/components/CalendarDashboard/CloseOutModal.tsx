import React, { useState, useEffect } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Dialog from '@material-ui/core/Dialog';
import DialogTitle from '@material-ui/core/DialogTitle';
import DialogContent from '@material-ui/core/DialogContent';
import DialogActions from '@material-ui/core/DialogActions';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import TextField from '@material-ui/core/TextField';
import FormControlLabel from '@material-ui/core/FormControlLabel';
import RadioGroup from '@material-ui/core/RadioGroup';
import Radio from '@material-ui/core/Radio';
import Divider from '@material-ui/core/Divider';
import Snackbar from '@material-ui/core/Snackbar';
import clsx from 'clsx';
import { calendarTokens } from './tokens';

type CloseType = 'full' | 'los' | 'reopen';

type DateRange = { id: number; start: string; end: string };

type StrategyRule = {
  id: number;
  operators: string[];
  roomTypes: string[];
  boardTypes: string[];
};

type Props = {
  open: boolean;
  onClose: () => void;
  initialDates?: { start: string; end: string };
  onConfirm?: (dates: DateRange[]) => void;
};

const OPERATORS = ['All', 'TUI Group', 'Thomas Cook', 'Sunwing', 'Club Med', 'Jet2'];
const ROOM_TYPES = ['All', 'Standard', 'Superior', 'Deluxe', 'Suite'];
const BOARD_TYPES = ['All', 'All Inclusive', 'Half Board', 'B&B', 'Room Only'];

let nextId = 1;

const useStyles = makeStyles((theme) => ({
  dialog: {
    '& .MuiDialog-paper': {
      width: 620,
      maxWidth: '95vw',
      maxHeight: '90vh',
      borderRadius: 6,
    },
  },
  titleBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(2, 2.5),
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  titleText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 16,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  content: {
    padding: theme.spacing(2.5),
    overflowY: 'auto',
  },
  sectionLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    color: theme.palette.text.disabled,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: theme.spacing(1),
  },
  typeCards: {
    display: 'flex',
    gap: theme.spacing(1.5),
    marginBottom: theme.spacing(2.5),
  },
  typeCard: {
    flex: 1,
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 6,
    padding: theme.spacing(1.5),
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    transition: 'border-color 0.15s, background-color 0.15s',
    '&:hover': {
      borderColor: theme.palette.primary.main,
      backgroundColor: calendarTokens.primaryHover,
    },
    '& .MuiIcon-root': { fontSize: '18px !important', color: theme.palette.text.secondary },
  },
  typeCardActive: {
    borderColor: theme.palette.primary.main,
    backgroundColor: calendarTokens.primaryHover,
    '& .MuiIcon-root': { color: theme.palette.primary.main },
  },
  typeCardLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 600,
    color: theme.palette.text.primary,
  },
  minNightsRow: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    marginBottom: theme.spacing(2.5),
  },
  minNightsLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.primary,
  },
  minNightsInput: {
    width: 72,
    '& .MuiInputBase-input': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.75, 1),
    },
    '& .MuiOutlinedInput-root fieldset': { borderColor: calendarTokens.border },
  },
  dateRangeSection: {
    marginBottom: theme.spacing(2.5),
  },
  dateRangeRow: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    marginBottom: theme.spacing(1),
  },
  dateInput: {
    flex: 1,
    '& .MuiInputBase-input': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.75, 1),
    },
    '& .MuiOutlinedInput-root fieldset': { borderColor: calendarTokens.border },
  },
  dateRangeSep: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
  },
  removeDateBtn: {
    color: theme.palette.text.disabled,
    padding: 4,
    '&:hover': { color: theme.palette.error.main },
    '& .MuiIcon-root': { fontSize: '16px !important' },
  },
  addBtn: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.primary.main,
    padding: theme.spacing(0.5, 0),
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  ruleSection: {
    marginBottom: theme.spacing(2.5),
  },
  ruleCard: {
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 6,
    padding: theme.spacing(1.5),
    marginBottom: theme.spacing(1),
  },
  ruleHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: theme.spacing(1),
  },
  ruleLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    color: theme.palette.text.disabled,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
  },
  chipRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: theme.spacing(0.5),
    marginBottom: theme.spacing(0.75),
  },
  chip: {
    padding: theme.spacing(0.375, 1),
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 14,
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
    '&:hover': { borderColor: theme.palette.primary.main, color: theme.palette.primary.main },
  },
  chipActive: {
    borderColor: theme.palette.primary.main,
    backgroundColor: calendarTokens.primaryHover,
    color: theme.palette.primary.main,
    fontWeight: 600,
  },
  formSection: {
    marginBottom: theme.spacing(2),
  },
  formLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.primary,
    marginBottom: theme.spacing(0.5),
    display: 'block',
  },
  textField: {
    '& .MuiInputBase-input, & .MuiInputBase-inputMultiline': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.875, 1.25),
    },
    '& .MuiOutlinedInput-root fieldset': { borderColor: calendarTokens.border },
  },
  radioLabel: {
    margin: 0,
    '& .MuiFormControlLabel-label': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
    },
    '& .MuiRadio-root': {
      padding: theme.spacing(0.5, 1, 0.5, 0),
      color: calendarTokens.checkboxBorder,
    },
    '& .Mui-checked': { color: `${theme.palette.primary.main} !important` },
  },
  actions: {
    borderTop: `1px solid ${calendarTokens.border}`,
    padding: theme.spacing(1.5, 2.5),
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
  },
  cancelBtn: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.text.secondary,
  },
  confirmBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
}));

function ChipGroup({
  options,
  selected,
  onChange,
}: {
  options: string[];
  selected: string[];
  onChange: (val: string[]) => void;
}) {
  const classes = useStyles();
  const toggle = (opt: string) => {
    if (opt === 'All') { onChange(['All']); return; }
    const without = selected.filter((s) => s !== 'All');
    onChange(without.includes(opt) ? without.filter((s) => s !== opt) : [...without, opt]);
  };
  return (
    <div className={classes.chipRow}>
      {options.map((opt) => (
        <span
          key={opt}
          className={clsx(classes.chip, selected.includes(opt) && classes.chipActive)}
          onClick={() => toggle(opt)}
        >
          {opt}
        </span>
      ))}
    </div>
  );
}

export function CloseOutModal({ open, onClose, initialDates, onConfirm }: Props) {
  const classes = useStyles();
  const [closeType, setCloseType] = useState<CloseType>('full');
  const [minNights, setMinNights] = useState(2);
  const [dateRanges, setDateRanges] = useState<DateRange[]>([]);
  const [rules, setRules] = useState<StrategyRule[]>([]);
  const [email, setEmail] = useState('');
  const [message, setMessage] = useState('');
  const [action, setAction] = useState('email');
  const [snackbar, setSnackbar] = useState(false);

  useEffect(() => {
    if (open) {
      const today = new Date().toISOString().split('T')[0];
      const end = initialDates?.end ?? today;
      const start = initialDates?.start ?? today;
      setDateRanges([{ id: nextId++, start, end }]);
      setRules([]);
      setCloseType('full');
      setMessage('');
    }
  }, [open, initialDates]);

  const addDateRange = () => {
    const today = new Date().toISOString().split('T')[0];
    setDateRanges((d) => [...d, { id: nextId++, start: today, end: today }]);
  };

  const removeDateRange = (id: number) => setDateRanges((d) => d.filter((r) => r.id !== id));

  const updateDate = (id: number, field: 'start' | 'end', val: string) =>
    setDateRanges((d) => d.map((r) => (r.id === id ? { ...r, [field]: val } : r)));

  const addRule = () =>
    setRules((r) => [
      ...r,
      { id: nextId++, operators: ['All'], roomTypes: ['All'], boardTypes: ['All'] },
    ]);

  const removeRule = (id: number) => setRules((r) => r.filter((rule) => rule.id !== id));

  const updateRule = (id: number, field: keyof Omit<StrategyRule, 'id'>, val: string[]) =>
    setRules((r) => r.map((rule) => (rule.id === id ? { ...rule, [field]: val } : rule)));

  const handleConfirm = () => {
    onConfirm?.(dateRanges);
    setSnackbar(true);
    onClose();
  };

  return (
    <>
      <Dialog open={open} onClose={onClose} className={classes.dialog} maxWidth={false}>
        <div className={classes.titleBar}>
          <Typography className={classes.titleText}>Close / Re-Open</Typography>
          <IconButton size="small" onClick={onClose}>
            <Icon style={{ fontSize: 18 }}>close</Icon>
          </IconButton>
        </div>

        <DialogContent className={classes.content}>
          {/* Type cards */}
          <Typography className={classes.sectionLabel}>Action Type</Typography>
          <div className={classes.typeCards}>
            {([
              { id: 'full' as CloseType, label: 'Close All Day', icon: 'block' },
              { id: 'los' as CloseType, label: 'Min Length of Stay', icon: 'night_shelter' },
              { id: 'reopen' as CloseType, label: 'Re-Open', icon: 'lock_open' },
            ]).map(({ id, label, icon }) => (
              <div
                key={id}
                className={clsx(classes.typeCard, closeType === id && classes.typeCardActive)}
                onClick={() => setCloseType(id)}
              >
                <Icon>{icon}</Icon>
                <Typography className={classes.typeCardLabel}>{label}</Typography>
              </div>
            ))}
          </div>

          {closeType === 'los' && (
            <div className={classes.minNightsRow}>
              <Typography className={classes.minNightsLabel}>Minimum nights:</Typography>
              <TextField
                value={minNights}
                onChange={(e) => setMinNights(Math.max(1, Number(e.target.value)))}
                type="number"
                variant="outlined"
                size="small"
                className={classes.minNightsInput}
                inputProps={{ min: 1, max: 30 }}
              />
            </div>
          )}

          <Divider style={{ marginBottom: 20 }} />

          {/* Date ranges */}
          <div className={classes.dateRangeSection}>
            <Typography className={classes.sectionLabel}>Date Ranges</Typography>
            {dateRanges.map((range) => (
              <div key={range.id} className={classes.dateRangeRow}>
                <TextField
                  type="date"
                  value={range.start}
                  onChange={(e) => updateDate(range.id, 'start', e.target.value)}
                  variant="outlined"
                  size="small"
                  className={classes.dateInput}
                />
                <Typography className={classes.dateRangeSep}>–</Typography>
                <TextField
                  type="date"
                  value={range.end}
                  onChange={(e) => updateDate(range.id, 'end', e.target.value)}
                  variant="outlined"
                  size="small"
                  className={classes.dateInput}
                />
                {dateRanges.length > 1 && (
                  <IconButton
                    className={classes.removeDateBtn}
                    size="small"
                    onClick={() => removeDateRange(range.id)}
                  >
                    <Icon>close</Icon>
                  </IconButton>
                )}
              </div>
            ))}
            <Button className={classes.addBtn} onClick={addDateRange} startIcon={<Icon>add</Icon>}>
              Add Date Range
            </Button>
          </div>

          <Divider style={{ marginBottom: 20 }} />

          {/* Strategy rules */}
          <div className={classes.ruleSection}>
            <Typography className={classes.sectionLabel}>Strategy Rules</Typography>
            {rules.map((rule) => (
              <div key={rule.id} className={classes.ruleCard}>
                <div className={classes.ruleHeader}>
                  <Typography className={classes.ruleLabel}>Rule</Typography>
                  <IconButton
                    className={classes.removeDateBtn}
                    size="small"
                    onClick={() => removeRule(rule.id)}
                  >
                    <Icon>close</Icon>
                  </IconButton>
                </div>
                <Typography className={classes.formLabel}>Tour Operator</Typography>
                <ChipGroup
                  options={OPERATORS}
                  selected={rule.operators}
                  onChange={(v) => updateRule(rule.id, 'operators', v)}
                />
                <Typography className={classes.formLabel}>Room Type</Typography>
                <ChipGroup
                  options={ROOM_TYPES}
                  selected={rule.roomTypes}
                  onChange={(v) => updateRule(rule.id, 'roomTypes', v)}
                />
                <Typography className={classes.formLabel}>Board Type</Typography>
                <ChipGroup
                  options={BOARD_TYPES}
                  selected={rule.boardTypes}
                  onChange={(v) => updateRule(rule.id, 'boardTypes', v)}
                />
              </div>
            ))}
            <Button className={classes.addBtn} onClick={addRule} startIcon={<Icon>add</Icon>}>
              Add Strategy
            </Button>
          </div>

          <Divider style={{ marginBottom: 20 }} />

          {/* Notification */}
          <div className={classes.formSection}>
            <Typography className={classes.sectionLabel}>Notification</Typography>
            <label className={classes.formLabel}>Email addresses</label>
            <TextField
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              variant="outlined"
              size="small"
              fullWidth
              placeholder="operator@example.com"
              className={classes.textField}
            />
          </div>
          <div className={classes.formSection}>
            <label className={classes.formLabel}>Sales message</label>
            <TextField
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              variant="outlined"
              multiline
              rows={3}
              fullWidth
              placeholder="Optional message to include..."
              className={classes.textField}
            />
          </div>
          <div className={classes.formSection}>
            <Typography className={classes.sectionLabel}>Send Action</Typography>
            <RadioGroup
              value={action}
              onChange={(_, v) => setAction(v)}
              row
            >
              {[
                { value: 'email', label: 'Email Operators' },
                { value: 'note', label: 'Internal Note' },
                { value: 'both', label: 'Both' },
              ].map(({ value, label }) => (
                <FormControlLabel
                  key={value}
                  value={value}
                  control={<Radio size="small" />}
                  label={label}
                  className={classes.radioLabel}
                />
              ))}
            </RadioGroup>
          </div>
        </DialogContent>

        <DialogActions className={classes.actions}>
          <Button className={classes.cancelBtn} onClick={onClose}>
            Cancel
          </Button>
          <Button
            className={classes.confirmBtn}
            variant="contained"
            disableElevation
            onClick={handleConfirm}
          >
            {closeType === 'reopen' ? 'Re-Open Days' : 'Close Out Days'}
          </Button>
        </DialogActions>
      </Dialog>

      <Snackbar
        open={snackbar}
        autoHideDuration={3500}
        onClose={() => setSnackbar(false)}
        message={closeType === 'reopen' ? 'Days re-opened successfully' : 'Days closed out successfully'}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      />
    </>
  );
}
