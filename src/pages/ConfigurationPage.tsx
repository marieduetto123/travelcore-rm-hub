import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Tabs from '@material-ui/core/Tabs';
import Tab from '@material-ui/core/Tab';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Paper from '@material-ui/core/Paper';
import TextField from '@material-ui/core/TextField';
import FormControlLabel from '@material-ui/core/FormControlLabel';
import RadioGroup from '@material-ui/core/RadioGroup';
import Radio from '@material-ui/core/Radio';
import Checkbox from '@material-ui/core/Checkbox';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import Snackbar from '@material-ui/core/Snackbar';
import Divider from '@material-ui/core/Divider';
import Table from '@material-ui/core/Table';
import TableHead from '@material-ui/core/TableHead';
import TableBody from '@material-ui/core/TableBody';
import TableRow from '@material-ui/core/TableRow';
import TableCell from '@material-ui/core/TableCell';
import Chip from '@material-ui/core/Chip';
import clsx from 'clsx';

const tokens = { border: '#dde1e2', sectionBg: '#f8f9fa', rowHover: 'rgba(0,100,97,0.04)' };

type Strategy = {
  id: number;
  name: string;
  stayFrom: string;
  stayTo: string;
  dba: string;
  dbaOp: string;
  demOp: string;
  demVal: number;
  roomTypes: string[];
  mealPlans: string[];
  active: boolean;
};

const INITIAL_STRATEGIES: Strategy[] = [
  {
    id: 1,
    name: 'High Season Close — TUI',
    stayFrom: '2026-07-01',
    stayTo: '2026-08-31',
    dba: '14',
    dbaOp: '<=',
    demOp: '>=',
    demVal: 88,
    roomTypes: ['Standard', 'Deluxe'],
    mealPlans: ['All Inclusive'],
    active: true,
  },
  {
    id: 2,
    name: 'Winter Min LOS',
    stayFrom: '2026-12-20',
    stayTo: '2027-01-05',
    dba: '30',
    dbaOp: '<=',
    demOp: '>',
    demVal: 75,
    roomTypes: ['All'],
    mealPlans: ['All'],
    active: false,
  },
];

const MEAL_PLAN_ROWS = [
  { id: 1, name: 'All Inclusive', code: 'AI' },
  { id: 2, name: 'Full Board', code: 'FB' },
  { id: 3, name: 'Half Board', code: 'HB' },
  { id: 4, name: 'Bed & Breakfast', code: 'BB' },
  { id: 5, name: 'Room Only', code: 'RO' },
];

let strategyIdCounter = 10;

const useStyles = makeStyles((theme) => ({
  root: { padding: theme.spacing(3), display: 'flex', flexDirection: 'column', gap: theme.spacing(2.5) },
  pageTitle: { fontFamily: 'Lato, sans-serif', fontSize: 20, fontWeight: 700, color: theme.palette.text.primary },
  card: { border: `1px solid ${tokens.border}`, borderRadius: 4, backgroundColor: theme.palette.common.white, overflow: 'hidden' },
  tabBar: {
    borderBottom: `1px solid ${tokens.border}`,
    minHeight: 44,
    '& .MuiTab-root': { fontFamily: 'Lato, sans-serif', fontSize: 13, textTransform: 'none', minHeight: 44, color: theme.palette.text.secondary },
    '& .Mui-selected': { color: theme.palette.primary.main, fontWeight: 700 },
    '& .MuiTabs-indicator': { backgroundColor: theme.palette.primary.main },
  },
  tabContent: { padding: theme.spacing(3) },
  sectionTitle: { fontFamily: 'Lato, sans-serif', fontSize: 14, fontWeight: 700, color: theme.palette.text.primary, marginBottom: theme.spacing(2) },
  fieldGroup: { marginBottom: theme.spacing(2.5) },
  fieldLabel: { fontFamily: 'Lato, sans-serif', fontSize: 12, fontWeight: 700, color: theme.palette.text.disabled, textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: theme.spacing(0.75), display: 'block' },
  radioLabel: {
    margin: 0,
    '& .MuiFormControlLabel-label': { fontFamily: 'Lato, sans-serif', fontSize: 13 },
    '& .MuiRadio-root': { padding: theme.spacing(0.5, 1, 0.5, 0) },
    '& .Mui-checked': { color: `${theme.palette.primary.main} !important` },
  },
  checkLabel: {
    margin: 0,
    '& .MuiFormControlLabel-label': { fontFamily: 'Lato, sans-serif', fontSize: 13 },
    '& .MuiCheckbox-root': { padding: theme.spacing(0.5, 1, 0.5, 0) },
    '& .Mui-checked': { color: `${theme.palette.primary.main} !important` },
  },
  table: {
    '& .MuiTableCell-head': {
      fontFamily: 'Lato, sans-serif', fontSize: 12, fontWeight: 700, color: theme.palette.text.secondary,
      borderBottom: `1px solid ${tokens.border}`, padding: theme.spacing(1.25, 2), backgroundColor: tokens.sectionBg,
    },
    '& .MuiTableCell-body': {
      fontFamily: 'Lato, sans-serif', fontSize: 13, color: theme.palette.text.primary,
      borderBottom: `1px solid ${tokens.border}`, padding: theme.spacing(0.75, 2),
    },
  },
  editableCell: {
    '& .MuiInputBase-input': { fontFamily: 'Lato, sans-serif', fontSize: 13, padding: theme.spacing(0.75, 1) },
    '& .MuiOutlinedInput-root fieldset': { borderColor: tokens.border },
  },
  primaryBtn: {
    backgroundColor: theme.palette.primary.main, color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif', fontSize: 13, textTransform: 'none', padding: theme.spacing(0.75, 2),
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  ghostBtn: {
    fontFamily: 'Lato, sans-serif', fontSize: 13, textTransform: 'none',
    color: theme.palette.text.secondary, padding: theme.spacing(0.75, 1.5),
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  strategyList: { display: 'flex', flexDirection: 'column', gap: theme.spacing(1.5) },
  strategyCard: {
    border: `1px solid ${tokens.border}`, borderRadius: 6, padding: theme.spacing(2),
    display: 'flex', alignItems: 'flex-start', gap: theme.spacing(1.5),
  },
  strategyActiveIndicator: { width: 4, borderRadius: 2, alignSelf: 'stretch', flexShrink: 0 },
  strategyName: { fontFamily: 'Lato, sans-serif', fontSize: 14, fontWeight: 700, color: theme.palette.text.primary },
  strategyMeta: { fontFamily: 'Lato, sans-serif', fontSize: 12, color: theme.palette.text.secondary, marginTop: 2 },
  strategyActions: { marginLeft: 'auto', display: 'flex', gap: theme.spacing(0.5) },
  strategyIconBtn: { color: '#9ca3af', padding: 4, '& .MuiIcon-root': { fontSize: '18px !important' } },
  formSection: { marginBottom: theme.spacing(2.5) },
  formRow: { display: 'flex', gap: theme.spacing(1.5), alignItems: 'flex-start', flexWrap: 'wrap' },
  formField: {
    '& .MuiInputBase-input': { fontFamily: 'Lato, sans-serif', fontSize: 13, padding: theme.spacing(0.875, 1.25) },
    '& .MuiOutlinedInput-root fieldset': { borderColor: tokens.border },
    '& .MuiInputLabel-root': { fontFamily: 'Lato, sans-serif', fontSize: 13 },
  },
  opSelect: {
    height: 38, fontFamily: 'Lato, sans-serif', fontSize: 13,
    '& .MuiSelect-select': { padding: theme.spacing(0.625, 3.5, 0.625, 1.25), fontFamily: 'Lato, sans-serif' },
    '& .MuiOutlinedInput-notchedOutline': { borderColor: tokens.border },
  },
  chipRow: { display: 'flex', flexWrap: 'wrap', gap: theme.spacing(0.5), marginTop: theme.spacing(0.75) },
  chip: {
    border: `1px solid ${tokens.border}`, borderRadius: 14, padding: theme.spacing(0.375, 1),
    cursor: 'pointer', fontFamily: 'Lato, sans-serif', fontSize: 12, color: theme.palette.text.secondary,
    '&:hover': { borderColor: theme.palette.primary.main, color: theme.palette.primary.main },
  },
  chipActive: { borderColor: theme.palette.primary.main, backgroundColor: 'rgba(0,100,97,0.08)', color: theme.palette.primary.main, fontWeight: 600 },
  stickyFooter: {
    position: 'sticky', bottom: 0, zIndex: 10,
    backgroundColor: theme.palette.common.white,
    borderTop: `2px solid ${theme.palette.primary.main}`,
    padding: theme.spacing(1.5, 3),
    display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: theme.spacing(1.5),
  },
  stopSalesGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: theme.spacing(2),
    marginBottom: theme.spacing(2.5),
  },
  stopSalesCard: { border: `1px solid ${tokens.border}`, borderRadius: 6, padding: theme.spacing(1.5) },
  stopSalesCardTitle: { fontFamily: 'Lato, sans-serif', fontSize: 13, fontWeight: 700, color: theme.palette.text.primary, marginBottom: theme.spacing(1) },
  configSelect: {
    width: '100%', height: 36, fontFamily: 'Lato, sans-serif', fontSize: 13,
    '& .MuiSelect-select': { padding: theme.spacing(0.625, 3.5, 0.625, 1.25), fontFamily: 'Lato, sans-serif' },
    '& .MuiOutlinedInput-notchedOutline': { borderColor: tokens.border },
  },
}));

const ROOM_TYPE_OPTIONS = ['All', 'Standard', 'Superior', 'Deluxe', 'Suite', 'Family'];
const MEAL_PLAN_OPTIONS = ['All', 'All Inclusive', 'Full Board', 'Half Board', 'B&B', 'Room Only'];

type StrategyFormState = Omit<Strategy, 'id'>;

const BLANK_FORM: StrategyFormState = {
  name: '',
  stayFrom: '',
  stayTo: '',
  dba: '14',
  dbaOp: '<=',
  demOp: '>=',
  demVal: 80,
  roomTypes: ['All'],
  mealPlans: ['All'],
  active: true,
};

export function ConfigurationPage() {
  const classes = useStyles();
  const [tab, setTab] = useState(0);
  const [strategies, setStrategies] = useState<Strategy[]>(INITIAL_STRATEGIES);
  const [formOpen, setFormOpen] = useState(false);
  const [editId, setEditId] = useState<number | null>(null);
  const [form, setForm] = useState<StrategyFormState>(BLANK_FORM);
  const [dirty, setDirty] = useState(false);
  const [snackbar, setSnackbar] = useState('');
  const [opNameField, setOpNameField] = useState('agent');
  const [sourceGeoField, setSourceGeoField] = useState('source_geo');

  const toggleChip = (field: 'roomTypes' | 'mealPlans', value: string) => {
    setForm((f) => {
      if (value === 'All') return { ...f, [field]: ['All'] };
      const without = f[field].filter((v) => v !== 'All');
      return {
        ...f,
        [field]: without.includes(value) ? without.filter((v) => v !== value) : [...without, value],
      };
    });
  };

  const handleNewStrategy = () => {
    setEditId(null);
    setForm(BLANK_FORM);
    setFormOpen(true);
  };

  const handleEditStrategy = (s: Strategy) => {
    setEditId(s.id);
    const { id, ...rest } = s;
    setForm(rest);
    setFormOpen(true);
  };

  const handleDeleteStrategy = (id: number) => {
    setStrategies((prev) => prev.filter((s) => s.id !== id));
  };

  const handleSaveStrategy = () => {
    if (!form.name.trim()) return;
    if (editId != null) {
      setStrategies((prev) => prev.map((s) => (s.id === editId ? { id: editId, ...form } : s)));
    } else {
      setStrategies((prev) => [...prev, { id: strategyIdCounter++, ...form }]);
    }
    setFormOpen(false);
    setSnackbar('Strategy saved');
  };

  const handleSaveConfig = () => {
    setDirty(false);
    setSnackbar('Configuration saved');
  };

  return (
    <div className={classes.root}>
      <Typography className={classes.pageTitle}>Configuration</Typography>

      <Paper className={classes.card} elevation={0}>
        <Tabs value={tab} onChange={(_, v) => setTab(v)} className={classes.tabBar}>
          <Tab label="Travel Distribution Hubs" />
          <Tab label="Stop Sales &amp; Alerts" />
          <Tab label="Autopilot" />
        </Tabs>

        {/* ── Tab 0: Travel Distribution Hubs ── */}
        {tab === 0 && (
          <div className={classes.tabContent}>
            <div className={classes.fieldGroup}>
              <span className={classes.fieldLabel}>Operator Name Field</span>
              <RadioGroup row value={opNameField} onChange={(_, v) => { setOpNameField(v); setDirty(true); }}>
                <FormControlLabel value="agent" control={<Radio size="small" />} label="Travel Agent" className={classes.radioLabel} />
                <FormControlLabel value="company" control={<Radio size="small" />} label="Company" className={classes.radioLabel} />
              </RadioGroup>
            </div>

            <div className={classes.fieldGroup}>
              <span className={classes.fieldLabel}>Source Geography</span>
              <RadioGroup row value={sourceGeoField} onChange={(_, v) => { setSourceGeoField(v); setDirty(true); }}>
                <FormControlLabel value="source_geo" control={<Radio size="small" />} label="Source Geo" className={classes.radioLabel} />
                <FormControlLabel value="origin" control={<Radio size="small" />} label="Origin" className={classes.radioLabel} />
              </RadioGroup>
            </div>

            <Divider style={{ margin: '16px 0' }} />

            <Typography className={classes.sectionTitle}>Meal Plans &amp; Codes</Typography>
            <Table className={classes.table} size="small" style={{ marginBottom: 16 }}>
              <TableHead>
                <TableRow>
                  <TableCell>Meal Plan Name</TableCell>
                  <TableCell>Code</TableCell>
                  <TableCell />
                </TableRow>
              </TableHead>
              <TableBody>
                {MEAL_PLAN_ROWS.map((row) => (
                  <TableRow key={row.id}>
                    <TableCell>
                      <TextField
                        defaultValue={row.name}
                        variant="outlined"
                        size="small"
                        className={classes.editableCell}
                        onChange={() => setDirty(true)}
                        style={{ width: 200 }}
                      />
                    </TableCell>
                    <TableCell>
                      <TextField
                        defaultValue={row.code}
                        variant="outlined"
                        size="small"
                        className={classes.editableCell}
                        onChange={() => setDirty(true)}
                        style={{ width: 80 }}
                      />
                    </TableCell>
                    <TableCell>
                      <IconButton size="small" style={{ color: '#9ca3af' }} onClick={() => setDirty(true)}>
                        <Icon style={{ fontSize: 16 }}>delete_outline</Icon>
                      </IconButton>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        )}

        {/* ── Tab 1: Stop Sales & Alerts ── */}
        {tab === 1 && (
          <div className={classes.tabContent}>
            <Typography className={classes.sectionTitle}>Room Type Visibility</Typography>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 24 }}>
              {['Standard', 'Superior', 'Deluxe', 'Suite', 'Family'].map((rt) => (
                <FormControlLabel
                  key={rt}
                  control={<Checkbox defaultChecked size="small" onChange={() => setDirty(true)} />}
                  label={rt}
                  className={classes.checkLabel}
                />
              ))}
            </div>

            <Divider style={{ marginBottom: 20 }} />

            <Typography className={classes.sectionTitle}>Alert Settings</Typography>
            <div className={classes.stopSalesGrid}>
              {[
                { label: 'Default Alert Frequency', options: ['Daily digest', 'Weekly digest', 'Monthly summary', 'Disabled'] },
                { label: 'Alert Channel', options: ['Email + In-app', 'Email only', 'In-app only'] },
                { label: 'Hotel-Wide Stop Threshold', options: ['85%', '90%', '95%', 'Disabled'] },
              ].map(({ label, options }) => (
                <div key={label} className={classes.stopSalesCard}>
                  <Typography className={classes.stopSalesCardTitle}>{label}</Typography>
                  <Select
                    defaultValue={options[0]}
                    variant="outlined"
                    className={classes.configSelect}
                    onChange={() => setDirty(true)}
                  >
                    {options.map((o) => (
                      <MenuItem key={o} value={o} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>{o}</MenuItem>
                    ))}
                  </Select>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── Tab 2: Autopilot ── */}
        {tab === 2 && (
          <div className={classes.tabContent}>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 20 }}>
              <Typography className={classes.sectionTitle} style={{ margin: 0, flex: 1 }}>
                Close-Out Strategies
              </Typography>
              <Button className={classes.primaryBtn} variant="contained" disableElevation onClick={handleNewStrategy}>
                <Icon>add</Icon>
                New Strategy
              </Button>
            </div>

            {!formOpen && (
              <div className={classes.strategyList}>
                {strategies.map((s) => (
                  <div key={s.id} className={classes.strategyCard}>
                    <div
                      className={classes.strategyActiveIndicator}
                      style={{ backgroundColor: s.active ? '#006461' : '#d1d5db' }}
                    />
                    <div style={{ flex: 1 }}>
                      <Typography className={classes.strategyName}>{s.name}</Typography>
                      <Typography className={classes.strategyMeta}>
                        Stay: {s.stayFrom} – {s.stayTo} · DBA {s.dbaOp} {s.dba} · Occ {s.demOp} {s.demVal}%
                      </Typography>
                      <div style={{ display: 'flex', gap: 4, marginTop: 6 }}>
                        {s.roomTypes.map((r) => (
                          <Chip key={r} label={r} style={{ height: 20, fontSize: 11, fontFamily: 'Lato, sans-serif' }} />
                        ))}
                        {s.mealPlans.map((m) => (
                          <Chip key={m} label={m} variant="outlined" style={{ height: 20, fontSize: 11, fontFamily: 'Lato, sans-serif' }} />
                        ))}
                      </div>
                    </div>
                    <div className={classes.strategyActions}>
                      <IconButton className={classes.strategyIconBtn} onClick={() => handleEditStrategy(s)}>
                        <Icon>edit</Icon>
                      </IconButton>
                      <IconButton className={classes.strategyIconBtn} onClick={() => handleDeleteStrategy(s.id)}>
                        <Icon>delete_outline</Icon>
                      </IconButton>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* Strategy form */}
            {formOpen && (
              <Paper style={{ border: `1px solid ${tokens.border}`, borderRadius: 6, padding: 24, marginTop: 8 }} elevation={0}>
                <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 15, fontWeight: 700, marginBottom: 20 }}>
                  {editId != null ? 'Edit Strategy' : 'New Strategy'}
                </Typography>

                <div className={classes.formSection}>
                  <TextField
                    value={form.name}
                    onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))}
                    label="Strategy Name"
                    variant="outlined"
                    size="small"
                    fullWidth
                    className={classes.formField}
                  />
                </div>

                <div className={classes.formSection}>
                  <span className={classes.fieldLabel}>Stay Date Range</span>
                  <div className={classes.formRow}>
                    <TextField
                      type="date"
                      value={form.stayFrom}
                      onChange={(e) => setForm((f) => ({ ...f, stayFrom: e.target.value }))}
                      variant="outlined"
                      size="small"
                      label="From"
                      className={classes.formField}
                      InputLabelProps={{ shrink: true }}
                      style={{ flex: 1, minWidth: 140 }}
                    />
                    <TextField
                      type="date"
                      value={form.stayTo}
                      onChange={(e) => setForm((f) => ({ ...f, stayTo: e.target.value }))}
                      variant="outlined"
                      size="small"
                      label="To"
                      className={classes.formField}
                      InputLabelProps={{ shrink: true }}
                      style={{ flex: 1, minWidth: 140 }}
                    />
                  </div>
                </div>

                <div className={classes.formSection}>
                  <span className={classes.fieldLabel}>Days Before Arrival</span>
                  <div className={classes.formRow}>
                    <Select
                      value={form.dbaOp}
                      onChange={(e) => setForm((f) => ({ ...f, dbaOp: e.target.value as string }))}
                      variant="outlined"
                      className={classes.opSelect}
                      style={{ width: 90 }}
                    >
                      {['<', '<=', '>', '>='].map((op) => (
                        <MenuItem key={op} value={op} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>{op}</MenuItem>
                      ))}
                    </Select>
                    <TextField
                      value={form.dba}
                      onChange={(e) => setForm((f) => ({ ...f, dba: e.target.value }))}
                      type="number"
                      variant="outlined"
                      size="small"
                      className={classes.formField}
                      inputProps={{ min: 1 }}
                      style={{ width: 80 }}
                    />
                    <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, alignSelf: 'center', color: '#6b7280' }}>days</Typography>
                  </div>
                </div>

                <div className={classes.formSection}>
                  <span className={classes.fieldLabel}>Demand Occupancy</span>
                  <div className={classes.formRow}>
                    <Select
                      value={form.demOp}
                      onChange={(e) => setForm((f) => ({ ...f, demOp: e.target.value as string }))}
                      variant="outlined"
                      className={classes.opSelect}
                      style={{ width: 90 }}
                    >
                      {['<', '<=', '>', '>='].map((op) => (
                        <MenuItem key={op} value={op} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>{op}</MenuItem>
                      ))}
                    </Select>
                    <TextField
                      value={form.demVal}
                      onChange={(e) => setForm((f) => ({ ...f, demVal: Number(e.target.value) }))}
                      type="number"
                      variant="outlined"
                      size="small"
                      className={classes.formField}
                      inputProps={{ min: 0, max: 100 }}
                      style={{ width: 80 }}
                    />
                    <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, alignSelf: 'center', color: '#6b7280' }}>%</Typography>
                  </div>
                </div>

                <div className={classes.formSection}>
                  <span className={classes.fieldLabel}>Room Types</span>
                  <div className={classes.chipRow}>
                    {ROOM_TYPE_OPTIONS.map((rt) => (
                      <span
                        key={rt}
                        className={clsx(classes.chip, form.roomTypes.includes(rt) && classes.chipActive)}
                        onClick={() => toggleChip('roomTypes', rt)}
                      >
                        {rt}
                      </span>
                    ))}
                  </div>
                </div>

                <div className={classes.formSection}>
                  <span className={classes.fieldLabel}>Meal Plans</span>
                  <div className={classes.chipRow}>
                    {MEAL_PLAN_OPTIONS.map((mp) => (
                      <span
                        key={mp}
                        className={clsx(classes.chip, form.mealPlans.includes(mp) && classes.chipActive)}
                        onClick={() => toggleChip('mealPlans', mp)}
                      >
                        {mp}
                      </span>
                    ))}
                  </div>
                </div>

                <FormControlLabel
                  control={
                    <Checkbox
                      checked={form.active}
                      onChange={(_, v) => setForm((f) => ({ ...f, active: v }))}
                      size="small"
                    />
                  }
                  label="Strategy active"
                  className={classes.checkLabel}
                  style={{ marginBottom: 20 }}
                />

                <div style={{ display: 'flex', gap: 12 }}>
                  <Button className={classes.ghostBtn} onClick={() => setFormOpen(false)}>
                    Cancel
                  </Button>
                  <Button className={classes.primaryBtn} variant="contained" disableElevation onClick={handleSaveStrategy}>
                    Save Strategy
                  </Button>
                </div>
              </Paper>
            )}
          </div>
        )}
      </Paper>

      {/* Sticky save footer */}
      {dirty && (
        <div className={classes.stickyFooter}>
          <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, color: '#6b7280' }}>
            You have unsaved changes
          </Typography>
          <Button className={classes.ghostBtn} onClick={() => setDirty(false)}>
            Discard
          </Button>
          <Button className={classes.primaryBtn} variant="contained" disableElevation onClick={handleSaveConfig}>
            Save Changes
          </Button>
        </div>
      )}

      <Snackbar
        open={Boolean(snackbar)}
        autoHideDuration={3000}
        onClose={() => setSnackbar('')}
        message={snackbar}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      />
    </div>
  );
}
