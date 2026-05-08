import React, { useState } from 'react';
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
import Checkbox from '@material-ui/core/Checkbox';
import FormControlLabel from '@material-ui/core/FormControlLabel';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import clsx from 'clsx';
import { calendarTokens } from './tokens';

export type HeatmapType = 'stop_sales' | 'hotel_occ' | 'remaining_rooms' | 'meal_plan' | 'to_forecast';

export type HeatmapConfig = {
  type: HeatmapType | null;
  thresholds: { grey: number; green: number; blue: number };
  roomTypes: string[];
  conditionEnabled: boolean;
  conditionMetric: string;
  conditionOp: string;
  conditionValue: number;
};

export const DEFAULT_HEATMAP: HeatmapConfig = {
  type: null,
  thresholds: { grey: 50, green: 75, blue: 90 },
  roomTypes: [],
  conditionEnabled: false,
  conditionMetric: 'hotel_occ',
  conditionOp: '>',
  conditionValue: 70,
};

type Props = {
  open: boolean;
  onClose: () => void;
  config?: HeatmapConfig;
  onApply?: (config: HeatmapConfig) => void;
};

const TYPE_CARDS: { id: HeatmapType; label: string; icon: string; description: string }[] = [
  { id: 'stop_sales', label: 'Stop Sales', icon: 'block', description: 'Show closed room types per day' },
  { id: 'hotel_occ', label: 'Hotel Occupancy', icon: 'hotel', description: 'Colour by occupancy %' },
  { id: 'remaining_rooms', label: 'Remaining Rooms', icon: 'bed', description: 'Colour by rooms remaining' },
  { id: 'meal_plan', label: 'Meal Plan Guests', icon: 'restaurant', description: 'Colour by meal plan volume' },
  { id: 'to_forecast', label: 'TO Forecast', icon: 'trending_up', description: 'Colour by tour operator forecast' },
];

const ROOM_TYPES = ['Standard', 'Superior', 'Deluxe', 'Suite', 'Family'];
const CONDITION_METRICS = [
  { value: 'hotel_occ', label: 'Hotel Occupancy' },
  { value: 'remaining_rooms', label: 'Remaining Rooms' },
  { value: 'meal_plan', label: 'Meal Plan Guests' },
  { value: 'to_otb', label: 'TO OTB' },
];
const CONDITION_OPS = ['>', '>=', '<', '<='];

const useStyles = makeStyles((theme) => ({
  dialog: {
    '& .MuiDialog-paper': {
      width: 560,
      maxWidth: '95vw',
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
    padding: theme.spacing(2, 2.5),
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
  typeGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: theme.spacing(1),
    marginBottom: theme.spacing(2.5),
  },
  typeCard: {
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 6,
    padding: theme.spacing(1.5),
    cursor: 'pointer',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    gap: theme.spacing(0.5),
    transition: 'border-color 0.15s, background-color 0.15s',
    '&:hover': {
      borderColor: theme.palette.primary.main,
      backgroundColor: calendarTokens.primaryHover,
    },
    '& .MuiIcon-root': {
      fontSize: '20px !important',
      color: theme.palette.text.secondary,
    },
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
  typeCardDesc: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.secondary,
    lineHeight: 1.3,
  },
  thresholdSection: {
    marginBottom: theme.spacing(2.5),
  },
  thresholdRow: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    marginBottom: theme.spacing(1),
  },
  colorDot: {
    width: 12,
    height: 12,
    borderRadius: '50%',
    flexShrink: 0,
  },
  thresholdLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
    minWidth: 60,
  },
  thresholdInput: {
    width: 80,
    '& .MuiInputBase-input': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.75, 1),
    },
    '& .MuiOutlinedInput-root': {
      '& fieldset': { borderColor: calendarTokens.border },
      '&:hover fieldset': { borderColor: theme.palette.text.secondary },
    },
  },
  roomTypeSection: {
    marginBottom: theme.spacing(2.5),
  },
  roomTypeChips: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: theme.spacing(0.75),
  },
  roomTypeChip: {
    padding: theme.spacing(0.5, 1.25),
    border: `1px solid ${calendarTokens.border}`,
    borderRadius: 16,
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
    '&:hover': { borderColor: theme.palette.primary.main, color: theme.palette.primary.main },
  },
  roomTypeChipActive: {
    borderColor: theme.palette.primary.main,
    backgroundColor: calendarTokens.primaryHover,
    color: theme.palette.primary.main,
    fontWeight: 600,
  },
  conditionSection: {
    borderTop: `1px solid ${calendarTokens.border}`,
    paddingTop: theme.spacing(2),
  },
  conditionControls: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    marginTop: theme.spacing(1),
    flexWrap: 'wrap',
  },
  conditionSelect: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    '& .MuiSelect-select': { padding: theme.spacing(0.75, 1) },
    '& fieldset': { borderColor: calendarTokens.border },
  },
  conditionValueInput: {
    width: 72,
    '& .MuiInputBase-input': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.75, 1),
    },
    '& .MuiOutlinedInput-root': {
      '& fieldset': { borderColor: calendarTokens.border },
    },
  },
  conditionCheckLabel: {
    '& .MuiFormControlLabel-label': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      color: theme.palette.text.primary,
    },
    '& .MuiCheckbox-root': { padding: theme.spacing(0.5) },
    '& .Mui-checked': { color: `${theme.palette.primary.main} !important` },
  },
  actions: {
    borderTop: `1px solid ${calendarTokens.border}`,
    padding: theme.spacing(1.5, 2.5),
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
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

export function HeatmapModal({ open, onClose, config = DEFAULT_HEATMAP, onApply }: Props) {
  const classes = useStyles();
  const [draft, setDraft] = useState<HeatmapConfig>(config);

  const handleReset = () => setDraft(DEFAULT_HEATMAP);

  const handleApply = () => {
    onApply?.(draft);
    onClose();
  };

  const toggleRoomType = (rt: string) => {
    setDraft((d) => ({
      ...d,
      roomTypes: d.roomTypes.includes(rt) ? d.roomTypes.filter((r) => r !== rt) : [...d.roomTypes, rt],
    }));
  };

  return (
    <Dialog open={open} onClose={onClose} className={classes.dialog} maxWidth={false}>
      <div className={classes.titleBar}>
        <Typography className={classes.titleText}>Heatmap</Typography>
        <IconButton size="small" onClick={onClose}>
          <Icon style={{ fontSize: 18 }}>close</Icon>
        </IconButton>
      </div>

      <DialogContent className={classes.content}>
        {/* Type selection */}
        <Typography className={classes.sectionLabel}>Select Type</Typography>
        <div className={classes.typeGrid}>
          {TYPE_CARDS.map((card) => (
            <div
              key={card.id}
              className={clsx(classes.typeCard, draft.type === card.id && classes.typeCardActive)}
              onClick={() => setDraft((d) => ({ ...d, type: card.id }))}
            >
              <Icon>{card.icon}</Icon>
              <Typography className={classes.typeCardLabel}>{card.label}</Typography>
              <Typography className={classes.typeCardDesc}>{card.description}</Typography>
            </div>
          ))}
        </div>

        {/* Thresholds (visible when type is selected) */}
        {draft.type && (
          <div className={classes.thresholdSection}>
            <Typography className={classes.sectionLabel}>Colour Thresholds</Typography>
            {(
              [
                { key: 'grey' as const, color: '#9ca3af', label: 'Grey zone ≤' },
                { key: 'green' as const, color: '#16a34a', label: 'Green zone ≤' },
                { key: 'blue' as const, color: '#2563eb', label: 'Blue zone ≤' },
              ] as const
            ).map(({ key, color, label }) => (
              <div key={key} className={classes.thresholdRow}>
                <div className={classes.colorDot} style={{ backgroundColor: color }} />
                <Typography className={classes.thresholdLabel}>{label}</Typography>
                <TextField
                  value={draft.thresholds[key]}
                  onChange={(e) =>
                    setDraft((d) => ({
                      ...d,
                      thresholds: { ...d.thresholds, [key]: Number(e.target.value) },
                    }))
                  }
                  type="number"
                  variant="outlined"
                  size="small"
                  className={classes.thresholdInput}
                  inputProps={{ min: 0, max: 100 }}
                />
                <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, color: '#6b7280' }}>
                  %
                </Typography>
              </div>
            ))}
          </div>
        )}

        {/* Room types (Stop Sales only) */}
        {draft.type === 'stop_sales' && (
          <div className={classes.roomTypeSection}>
            <Typography className={classes.sectionLabel}>Room Types</Typography>
            <div className={classes.roomTypeChips}>
              {ROOM_TYPES.map((rt) => (
                <span
                  key={rt}
                  className={clsx(classes.roomTypeChip, draft.roomTypes.includes(rt) && classes.roomTypeChipActive)}
                  onClick={() => toggleRoomType(rt)}
                >
                  {rt}
                </span>
              ))}
            </div>
          </div>
        )}

        {/* Condition */}
        <div className={classes.conditionSection}>
          <FormControlLabel
            control={
              <Checkbox
                checked={draft.conditionEnabled}
                onChange={(_, v) => setDraft((d) => ({ ...d, conditionEnabled: v }))}
                size="small"
              />
            }
            label="Add condition"
            className={classes.conditionCheckLabel}
          />
          {draft.conditionEnabled && (
            <div className={classes.conditionControls}>
              <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>When</Typography>
              <Select
                value={draft.conditionMetric}
                onChange={(e) => setDraft((d) => ({ ...d, conditionMetric: e.target.value as string }))}
                variant="outlined"
                className={classes.conditionSelect}
              >
                {CONDITION_METRICS.map((m) => (
                  <MenuItem key={m.value} value={m.value} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>
                    {m.label}
                  </MenuItem>
                ))}
              </Select>
              <Select
                value={draft.conditionOp}
                onChange={(e) => setDraft((d) => ({ ...d, conditionOp: e.target.value as string }))}
                variant="outlined"
                className={classes.conditionSelect}
                style={{ minWidth: 64 }}
              >
                {CONDITION_OPS.map((op) => (
                  <MenuItem key={op} value={op} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>
                    {op}
                  </MenuItem>
                ))}
              </Select>
              <TextField
                value={draft.conditionValue}
                onChange={(e) => setDraft((d) => ({ ...d, conditionValue: Number(e.target.value) }))}
                type="number"
                variant="outlined"
                size="small"
                className={classes.conditionValueInput}
                inputProps={{ min: 0 }}
              />
            </div>
          )}
        </div>
      </DialogContent>

      <DialogActions className={classes.actions}>
        <Button className={classes.resetBtn} onClick={handleReset}>
          Reset
        </Button>
        <Button className={classes.applyBtn} variant="contained" disableElevation onClick={handleApply}>
          Apply Heatmap
        </Button>
      </DialogActions>
    </Dialog>
  );
}
