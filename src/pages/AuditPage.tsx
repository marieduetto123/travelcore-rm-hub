import React, { useState, useMemo } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import TextField from '@material-ui/core/TextField';
import InputAdornment from '@material-ui/core/InputAdornment';
import Paper from '@material-ui/core/Paper';
import Table from '@material-ui/core/Table';
import TableHead from '@material-ui/core/TableHead';
import TableBody from '@material-ui/core/TableBody';
import TableRow from '@material-ui/core/TableRow';
import TableCell from '@material-ui/core/TableCell';
import TableSortLabel from '@material-ui/core/TableSortLabel';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import FormControlLabel from '@material-ui/core/FormControlLabel';
import Switch from '@material-ui/core/Switch';
import Popover from '@material-ui/core/Popover';
import Divider from '@material-ui/core/Divider';
import RadioGroup from '@material-ui/core/RadioGroup';
import Radio from '@material-ui/core/Radio';

const tokens = { border: '#dde1e2', rowHover: 'rgba(0,100,97,0.04)', sectionBg: '#f8f9fa' };

type AuditRow = {
  id: number;
  period: string;
  operator: string;
  roomType: string;
  board: string;
  segment: string;
  revenue: number;
  lyRevenue?: number;
  rn: number;
  adr: number;
  occ: number;
};

const ROWS: AuditRow[] = [
  { id: 1, period: 'Jan 2026', operator: 'TUI Group', roomType: 'Standard', board: 'All Inclusive', segment: 'Travel Distribution Hub', revenue: 142000, lyRevenue: 128000, rn: 498, adr: 285, occ: 91 },
  { id: 2, period: 'Jan 2026', operator: 'Thomas Cook', roomType: 'Deluxe', board: 'Half Board', segment: 'Travel Distribution Hub', revenue: 89000, lyRevenue: 81000, rn: 312, adr: 285, occ: 86 },
  { id: 3, period: 'Jan 2026', operator: 'Jet2holidays', roomType: 'Standard', board: 'B&B', segment: 'OTA', revenue: 54000, lyRevenue: 62000, rn: 189, adr: 286, occ: 74 },
  { id: 4, period: 'Jan 2026', operator: 'Club Med', roomType: 'Suite', board: 'All Inclusive', segment: 'Travel Distribution Hub', revenue: 196000, lyRevenue: 180000, rn: 620, adr: 316, occ: 95 },
  { id: 5, period: 'Feb 2026', operator: 'TUI Group', roomType: 'Standard', board: 'All Inclusive', segment: 'Travel Distribution Hub', revenue: 138000, lyRevenue: 125000, rn: 484, adr: 285, occ: 89 },
  { id: 6, period: 'Feb 2026', operator: 'FTI Group', roomType: 'Superior', board: 'Room Only', segment: 'Corporate', revenue: 71000, lyRevenue: 69000, rn: 248, adr: 286, occ: 79 },
  { id: 7, period: 'Feb 2026', operator: 'Sunwing', roomType: 'Deluxe', board: 'Half Board', segment: 'OTA', revenue: 48000, lyRevenue: 55000, rn: 168, adr: 286, occ: 68 },
  { id: 8, period: 'Mar 2026', operator: 'TUI Group', roomType: 'Standard', board: 'All Inclusive', segment: 'Travel Distribution Hub', revenue: 159000, lyRevenue: 140000, rn: 556, adr: 286, occ: 93 },
];

type SortDir = 'asc' | 'desc';
type SortKey = keyof AuditRow;

const fmt = (n: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(n);

const useStyles = makeStyles((theme) => ({
  root: { padding: theme.spacing(3), display: 'flex', flexDirection: 'column', gap: theme.spacing(2.5) },
  pageTitle: { fontFamily: 'Lato, sans-serif', fontSize: 20, fontWeight: 700, color: theme.palette.text.primary },
  card: { border: `1px solid ${tokens.border}`, borderRadius: 4, backgroundColor: theme.palette.common.white, overflow: 'hidden' },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
    flexWrap: 'wrap',
  },
  searchInput: {
    width: 220,
    '& .MuiInputBase-input': { fontFamily: 'Lato, sans-serif', fontSize: 13, padding: theme.spacing(0.875, 1) },
    '& .MuiOutlinedInput-root fieldset': { borderColor: tokens.border },
  },
  periodSelect: {
    height: 36,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    '& .MuiSelect-select': { padding: theme.spacing(0.5, 3.5, 0.5, 1.25), fontFamily: 'Lato, sans-serif' },
    '& .MuiOutlinedInput-notchedOutline': { borderColor: tokens.border },
  },
  actionBtn: {
    height: 36,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    color: theme.palette.text.secondary,
    padding: theme.spacing(0, 1.5),
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  spacer: { flex: 1 },
  lySwitch: {
    '& .MuiFormControlLabel-label': { fontFamily: 'Lato, sans-serif', fontSize: 13, color: theme.palette.text.secondary },
    '& .MuiSwitch-colorPrimary.Mui-checked': { color: theme.palette.primary.main },
    '& .MuiSwitch-colorPrimary.Mui-checked + .MuiSwitch-track': { backgroundColor: theme.palette.primary.main },
  },
  filtersBadge: {
    width: 16,
    height: 16,
    borderRadius: '50%',
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontSize: 10,
    fontWeight: 700,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginLeft: 4,
  },
  filterPopover: {
    width: 260,
    padding: 0,
  },
  filterHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
  },
  filterHeaderTitle: { fontFamily: 'Lato, sans-serif', fontSize: 14, fontWeight: 700, color: theme.palette.text.primary },
  filterBody: { padding: theme.spacing(0, 2), maxHeight: 320, overflowY: 'auto' },
  filterSection: { paddingTop: theme.spacing(1.5), paddingBottom: theme.spacing(1) },
  filterSectionLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    fontWeight: 700,
    color: theme.palette.text.disabled,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: theme.spacing(0.5),
  },
  radioLabel: {
    margin: 0,
    '& .MuiFormControlLabel-label': { fontFamily: 'Lato, sans-serif', fontSize: 13 },
    '& .MuiRadio-root': { padding: theme.spacing(0.5, 1, 0.5, 0), color: '#6b7280' },
    '& .Mui-checked': { color: `${theme.palette.primary.main} !important` },
  },
  filterFooter: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 2),
    borderTop: `1px solid ${tokens.border}`,
  },
  applyBtn: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none' as const,
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
  table: {
    '& .MuiTableCell-head': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 12,
      fontWeight: 700,
      color: theme.palette.text.secondary,
      borderBottom: `1px solid ${tokens.border}`,
      padding: theme.spacing(1.25, 2),
      backgroundColor: tokens.sectionBg,
    },
    '& .MuiTableCell-body': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      color: theme.palette.text.primary,
      borderBottom: `1px solid ${tokens.border}`,
      padding: theme.spacing(1.25, 2),
    },
    '& .MuiTableRow-root:hover .MuiTableCell-body': { backgroundColor: tokens.rowHover },
  },
  deltaPositive: { color: '#15803d', fontWeight: 700 },
  deltaNegative: { color: '#dc2626', fontWeight: 700 },
}));

type Filters = { roomType: string; board: string; segment: string; operator: string };
const DEFAULT_FILTERS: Filters = { roomType: 'all', board: 'all', segment: 'all', operator: 'all' };

export function AuditPage() {
  const classes = useStyles();
  const [search, setSearch] = useState('');
  const [compareLY, setCompareLY] = useState(false);
  const [period, setPeriod] = useState('custom');
  const [filtersAnchor, setFiltersAnchor] = useState<HTMLElement | null>(null);
  const [filters, setFilters] = useState<Filters>(DEFAULT_FILTERS);
  const [draftFilters, setDraftFilters] = useState<Filters>(DEFAULT_FILTERS);
  const [sortKey, setSortKey] = useState<SortKey>('period');
  const [sortDir, setSortDir] = useState<SortDir>('asc');

  const activeFilterCount = Object.values(filters).filter((v) => v !== 'all').length;

  const openFilters = (e: React.MouseEvent<HTMLElement>) => {
    setDraftFilters(filters);
    setFiltersAnchor(e.currentTarget);
  };

  const applyFilters = () => {
    setFilters(draftFilters);
    setFiltersAnchor(null);
  };

  const handleSort = (key: SortKey) =>
    setSortKey((k) => {
      if (k === key) setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
      return key;
    });

  const rows = useMemo(() => {
    const q = search.toLowerCase();
    return [...ROWS]
      .filter((r) => {
        const matchSearch = !q || r.operator.toLowerCase().includes(q) || r.period.toLowerCase().includes(q);
        const matchRoomType = filters.roomType === 'all' || r.roomType === filters.roomType;
        const matchBoard = filters.board === 'all' || r.board === filters.board;
        const matchSeg = filters.segment === 'all' || r.segment === filters.segment;
        const matchOp = filters.operator === 'all' || r.operator === filters.operator;
        return matchSearch && matchRoomType && matchBoard && matchSeg && matchOp;
      })
      .sort((a, b) => {
        const av = a[sortKey] as string | number | undefined;
        const bv = b[sortKey] as string | number | undefined;
        const cmp = String(av ?? '').localeCompare(String(bv ?? ''), undefined, { numeric: true });
        return sortDir === 'asc' ? cmp : -cmp;
      });
  }, [search, filters, sortKey, sortDir]);

  const sortProps = (key: SortKey) => ({
    active: sortKey === key,
    direction: sortKey === key ? sortDir : ('asc' as SortDir),
    onClick: () => handleSort(key),
  });

  return (
    <div className={classes.root}>
      <Typography className={classes.pageTitle}>Audit</Typography>

      <Paper className={classes.card} elevation={0}>
        <div className={classes.toolbar}>
          <TextField
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search…"
            variant="outlined"
            size="small"
            className={classes.searchInput}
            InputProps={{
              startAdornment: (
                <InputAdornment position="start">
                  <Icon style={{ fontSize: 16, color: '#9ca3af' }}>search</Icon>
                </InputAdornment>
              ),
            }}
          />

          <Select
            value={period}
            onChange={(e) => setPeriod(e.target.value as string)}
            variant="outlined"
            className={classes.periodSelect}
          >
            {[
              { value: 'custom', label: 'Custom Range' },
              { value: 'month', label: 'Month' },
              { value: 'quarter', label: 'Quarter' },
              { value: 'year', label: 'Year' },
            ].map((o) => (
              <MenuItem key={o.value} value={o.value} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>
                {o.label}
              </MenuItem>
            ))}
          </Select>

          <Button className={classes.actionBtn} onClick={openFilters}>
            <Icon>filter_list</Icon>
            Filters
            {activeFilterCount > 0 && <span className={classes.filtersBadge}>{activeFilterCount}</span>}
          </Button>

          <FormControlLabel
            control={
              <Switch
                checked={compareLY}
                onChange={(_, v) => setCompareLY(v)}
                size="small"
              />
            }
            label="Compare to LY"
            className={classes.lySwitch}
          />

          <span className={classes.spacer} />

          <Button className={classes.actionBtn}>
            <Icon>download</Icon>
            Export
          </Button>
        </div>

        <Table className={classes.table} size="small">
          <TableHead>
            <TableRow>
              <TableCell><TableSortLabel {...sortProps('period')}>Period</TableSortLabel></TableCell>
              <TableCell><TableSortLabel {...sortProps('operator')}>Operator</TableSortLabel></TableCell>
              <TableCell><TableSortLabel {...sortProps('roomType')}>Room Type</TableSortLabel></TableCell>
              <TableCell><TableSortLabel {...sortProps('board')}>Board</TableSortLabel></TableCell>
              <TableCell><TableSortLabel {...sortProps('segment')}>Segment</TableSortLabel></TableCell>
              <TableCell align="right"><TableSortLabel {...sortProps('revenue')}>Revenue</TableSortLabel></TableCell>
              {compareLY && <TableCell align="right">LY Revenue</TableCell>}
              {compareLY && <TableCell align="right">Δ vs LY</TableCell>}
              <TableCell align="right"><TableSortLabel {...sortProps('rn')}>RN</TableSortLabel></TableCell>
              <TableCell align="right"><TableSortLabel {...sortProps('adr')}>ADR</TableSortLabel></TableCell>
              <TableCell align="right"><TableSortLabel {...sortProps('occ')}>Occ%</TableSortLabel></TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {rows.map((row) => {
              const delta = compareLY && row.lyRevenue != null ? row.revenue - row.lyRevenue : null;
              const deltaPct = delta != null && row.lyRevenue ? ((delta / row.lyRevenue) * 100).toFixed(1) : null;
              return (
                <TableRow key={row.id}>
                  <TableCell>{row.period}</TableCell>
                  <TableCell>{row.operator}</TableCell>
                  <TableCell>{row.roomType}</TableCell>
                  <TableCell>{row.board}</TableCell>
                  <TableCell>{row.segment}</TableCell>
                  <TableCell align="right">{fmt(row.revenue)}</TableCell>
                  {compareLY && <TableCell align="right">{row.lyRevenue ? fmt(row.lyRevenue) : '–'}</TableCell>}
                  {compareLY && (
                    <TableCell align="right">
                      {deltaPct != null ? (
                        <span className={Number(deltaPct) >= 0 ? classes.deltaPositive : classes.deltaNegative}>
                          {Number(deltaPct) >= 0 ? '+' : ''}{deltaPct}%
                        </span>
                      ) : '–'}
                    </TableCell>
                  )}
                  <TableCell align="right">{row.rn.toLocaleString()}</TableCell>
                  <TableCell align="right">${row.adr}</TableCell>
                  <TableCell align="right">{row.occ}%</TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      </Paper>

      {/* Filters popover */}
      <Popover
        open={Boolean(filtersAnchor)}
        anchorEl={filtersAnchor}
        onClose={() => setFiltersAnchor(null)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
        transformOrigin={{ vertical: 'top', horizontal: 'left' }}
        PaperProps={{ className: classes.filterPopover }}
      >
        <div className={classes.filterHeader}>
          <Typography className={classes.filterHeaderTitle}>Filters</Typography>
        </div>
        <div className={classes.filterBody}>
          {[
            { label: 'Room Type', key: 'roomType' as keyof Filters, opts: ['all', 'Standard', 'Superior', 'Deluxe', 'Suite'] },
            { label: 'Board', key: 'board' as keyof Filters, opts: ['all', 'All Inclusive', 'Half Board', 'B&B', 'Room Only'] },
            { label: 'Segment', key: 'segment' as keyof Filters, opts: ['all', 'Travel Distribution Hub', 'OTA', 'Corporate'] },
            { label: 'Tour Operator', key: 'operator' as keyof Filters, opts: ['all', 'TUI Group', 'Thomas Cook', 'Jet2holidays', 'FTI Group', 'Sunwing', 'Club Med'] },
          ].map(({ label, key, opts }) => (
            <React.Fragment key={key}>
              <div className={classes.filterSection}>
                <Typography className={classes.filterSectionLabel}>{label}</Typography>
                <RadioGroup
                  value={draftFilters[key]}
                  onChange={(_, v) => setDraftFilters((f) => ({ ...f, [key]: v }))}
                >
                  {opts.map((opt) => (
                    <FormControlLabel
                      key={opt}
                      value={opt}
                      control={<Radio size="small" />}
                      label={opt === 'all' ? 'All' : opt}
                      className={classes.radioLabel}
                    />
                  ))}
                </RadioGroup>
              </div>
              <Divider />
            </React.Fragment>
          ))}
        </div>
        <div className={classes.filterFooter}>
          <Button
            style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, textTransform: 'none' }}
            onClick={() => { setDraftFilters(DEFAULT_FILTERS); setFilters(DEFAULT_FILTERS); setFiltersAnchor(null); }}
          >
            Clear All
          </Button>
          <Button
            variant="contained"
            disableElevation
            onClick={applyFilters}
            className={classes.applyBtn}
          >
            Apply
          </Button>
        </div>
      </Popover>
    </div>
  );
}
