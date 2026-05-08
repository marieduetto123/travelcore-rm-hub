import React, { useState, useMemo } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Tabs from '@material-ui/core/Tabs';
import Tab from '@material-ui/core/Tab';
import Table from '@material-ui/core/Table';
import TableHead from '@material-ui/core/TableHead';
import TableBody from '@material-ui/core/TableBody';
import TableRow from '@material-ui/core/TableRow';
import TableCell from '@material-ui/core/TableCell';
import TableSortLabel from '@material-ui/core/TableSortLabel';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import TextField from '@material-ui/core/TextField';
import InputAdornment from '@material-ui/core/InputAdornment';
import Paper from '@material-ui/core/Paper';
import Chip from '@material-ui/core/Chip';
import Divider from '@material-ui/core/Divider';
import Dialog from '@material-ui/core/Dialog';
import DialogTitle from '@material-ui/core/DialogTitle';
import DialogContent from '@material-ui/core/DialogContent';
import DialogActions from '@material-ui/core/DialogActions';
import clsx from 'clsx';

const tokens = {
  border: '#dde1e2',
  rowHover: 'rgba(0,100,97,0.04)',
  sectionBg: '#f8f9fa',
};

type Operator = {
  id: number;
  name: string;
  contact: string;
  country: string;
  status: 'Active' | 'Inactive' | 'Pending';
  revenue: string;
  bookings: number;
  closureRate: string;
};

type Contract = {
  id: number;
  operator: string;
  type: string;
  season: string;
  rooms: number;
  rate: string;
  status: 'Active' | 'Pending' | 'Expired';
  signed: string;
};

const OPERATORS: Operator[] = [
  { id: 1, name: 'TUI Group', contact: 'Anna Schmidt', country: 'Germany', status: 'Active', revenue: '$420k', bookings: 1240, closureRate: '8.2%' },
  { id: 2, name: 'Thomas Cook', contact: 'James Wilson', country: 'UK', status: 'Active', revenue: '$310k', bookings: 890, closureRate: '5.1%' },
  { id: 3, name: 'Jet2holidays', contact: 'Sarah Davies', country: 'UK', status: 'Active', revenue: '$185k', bookings: 620, closureRate: '3.4%' },
  { id: 4, name: 'FTI Group', contact: 'Klaus Müller', country: 'Germany', status: 'Active', revenue: '$140k', bookings: 510, closureRate: '2.9%' },
  { id: 5, name: 'Sunwing', contact: 'Marc Leblanc', country: 'Canada', status: 'Inactive', revenue: '$95k', bookings: 310, closureRate: '6.7%' },
  { id: 6, name: 'Club Med', contact: 'Sophie Durand', country: 'France', status: 'Active', revenue: '$260k', bookings: 740, closureRate: '4.2%' },
];

const CONTRACTS: Contract[] = [
  { id: 1, operator: 'TUI Group', type: 'Allotment', season: 'Summer 2026', rooms: 120, rate: '$285/night', status: 'Active', signed: '2025-11-15' },
  { id: 2, operator: 'Thomas Cook', type: 'Guarantee', season: 'Winter 2025/26', rooms: 80, rate: '$220/night', status: 'Active', signed: '2025-09-01' },
  { id: 3, operator: 'Jet2holidays', type: 'Allotment', season: 'Summer 2026', rooms: 55, rate: '$275/night', status: 'Pending', signed: '–' },
  { id: 4, operator: 'FTI Group', type: 'Free Sale', season: 'Annual 2026', rooms: 40, rate: '$260/night', status: 'Active', signed: '2025-12-01' },
  { id: 5, operator: 'Club Med', type: 'Guarantee', season: 'Summer 2026', rooms: 95, rate: '$310/night', status: 'Active', signed: '2025-10-20' },
];

type SortDir = 'asc' | 'desc';

const useStyles = makeStyles((theme) => ({
  root: {
    padding: theme.spacing(3),
    display: 'flex',
    flexDirection: 'column',
    gap: theme.spacing(2.5),
  },
  pageTitle: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 20,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  card: {
    border: `1px solid ${tokens.border}`,
    borderRadius: 4,
    backgroundColor: theme.palette.common.white,
    overflow: 'hidden',
  },
  tabBar: {
    borderBottom: `1px solid ${tokens.border}`,
    minHeight: 44,
    '& .MuiTab-root': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      textTransform: 'none',
      minHeight: 44,
      color: theme.palette.text.secondary,
    },
    '& .Mui-selected': { color: theme.palette.primary.main, fontWeight: 700 },
    '& .MuiTabs-indicator': { backgroundColor: theme.palette.primary.main },
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
    flexWrap: 'wrap',
  },
  searchInput: {
    width: 240,
    '& .MuiInputBase-input': {
      fontFamily: 'Lato, sans-serif',
      fontSize: 13,
      padding: theme.spacing(0.875, 1),
    },
    '& .MuiOutlinedInput-root fieldset': { borderColor: tokens.border },
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
  primaryBtn: {
    height: 36,
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    padding: theme.spacing(0, 1.5),
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
  },
  spacer: { flex: 1 },
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
    '& .MuiTableRow-root:hover .MuiTableCell-body': {
      backgroundColor: tokens.rowHover,
    },
    '& .MuiTableRow-root': { cursor: 'pointer' },
  },
  selectedRow: {
    '& .MuiTableCell-body': {
      backgroundColor: `${tokens.rowHover} !important`,
    },
  },
  statusChip: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    fontWeight: 600,
    height: 24,
    '& .MuiChip-label': { padding: theme.spacing(0, 1) },
  },
  statusActive: { backgroundColor: '#dcfce7', color: '#15803d' },
  statusPending: { backgroundColor: '#fef9c3', color: '#92400e' },
  statusInactive: { backgroundColor: '#f3f4f6', color: '#6b7280' },
  statusExpired: { backgroundColor: '#fee2e2', color: '#991f1f' },

  // Detail panel
  detailPanel: {
    borderTop: `1px solid ${tokens.border}`,
    padding: theme.spacing(2.5),
    backgroundColor: tokens.sectionBg,
  },
  detailHeader: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: theme.spacing(2),
  },
  detailName: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 18,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  detailContact: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
    marginTop: theme.spacing(0.25),
  },
  statsRow: {
    display: 'flex',
    gap: theme.spacing(2),
    marginBottom: theme.spacing(2),
    flexWrap: 'wrap',
  },
  statCard: {
    border: `1px solid ${tokens.border}`,
    borderRadius: 6,
    padding: theme.spacing(1.5, 2),
    backgroundColor: theme.palette.common.white,
    minWidth: 130,
  },
  statValue: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 20,
    fontWeight: 700,
    color: theme.palette.text.primary,
  },
  statLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.secondary,
    marginTop: theme.spacing(0.25),
  },

  // Insights
  insightsToolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
    flexWrap: 'wrap',
  },
  viewToggle: {
    display: 'flex',
    border: `1px solid ${tokens.border}`,
    borderRadius: 4,
    overflow: 'hidden',
  },
  viewToggleBtn: {
    height: 32,
    padding: theme.spacing(0, 1.5),
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    borderRadius: 0,
    color: theme.palette.text.secondary,
    border: 'none',
    '&:not(:last-child)': { borderRight: `1px solid ${tokens.border}` },
  },
  viewToggleBtnActive: {
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    '&:hover': { backgroundColor: theme.palette.primary.dark },
  },
  insightBlock: {
    borderBottom: `1px solid ${tokens.border}`,
  },
  insightHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(1.5, 2),
    cursor: 'pointer',
    '&:hover': { backgroundColor: tokens.rowHover },
  },
  insightHeaderLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    fontWeight: 700,
    color: theme.palette.text.primary,
    flex: 1,
  },
  insightBody: {
    padding: theme.spacing(1, 2, 2),
  },
  insightMetricRow: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(0.75, 0),
    borderBottom: `1px solid ${tokens.border}`,
  },
  insightMetricLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
    width: 180,
    flexShrink: 0,
  },
  progressBar: {
    flex: 1,
    height: 6,
    borderRadius: 3,
    backgroundColor: tokens.border,
    overflow: 'hidden',
    margin: theme.spacing(0, 2),
  },
  progressFill: {
    height: '100%',
    borderRadius: 3,
    backgroundColor: theme.palette.primary.main,
  },
  insightMetricValue: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 700,
    color: theme.palette.text.primary,
    minWidth: 80,
    textAlign: 'right',
  },
}));

type SortConfig<T> = { key: keyof T; dir: SortDir };

export function ContactsPage() {
  const classes = useStyles();
  const [tab, setTab] = useState(0);
  const [search, setSearch] = useState('');
  const [contractSearch, setContractSearch] = useState('');
  const [selectedOp, setSelectedOp] = useState<Operator | null>(null);
  const [insightView, setInsightView] = useState<'to' | 'contract' | 'overall'>('to');
  const [expandedOps, setExpandedOps] = useState<Set<number>>(new Set([1]));
  const [newOpOpen, setNewOpOpen] = useState(false);
  const [opSort, setOpSort] = useState<SortConfig<Operator>>({ key: 'revenue', dir: 'desc' });
  const [ctSort, setCtSort] = useState<SortConfig<Contract>>({ key: 'signed', dir: 'desc' });

  const filteredOps = useMemo(() => {
    const q = search.toLowerCase();
    const sorted = [...OPERATORS].sort((a, b) => {
      const av = a[opSort.key] as string | number;
      const bv = b[opSort.key] as string | number;
      const cmp = String(av).localeCompare(String(bv));
      return opSort.dir === 'asc' ? cmp : -cmp;
    });
    return q ? sorted.filter((o) => o.name.toLowerCase().includes(q) || o.contact.toLowerCase().includes(q)) : sorted;
  }, [search, opSort]);

  const filteredContracts = useMemo(() => {
    const q = contractSearch.toLowerCase();
    const sorted = [...CONTRACTS].sort((a, b) => {
      const av = String(a[ctSort.key]);
      const bv = String(b[ctSort.key]);
      const cmp = av.localeCompare(bv);
      return ctSort.dir === 'asc' ? cmp : -cmp;
    });
    return q ? sorted.filter((c) => c.operator.toLowerCase().includes(q) || c.type.toLowerCase().includes(q)) : sorted;
  }, [contractSearch, ctSort]);

  const handleOpSort = (key: keyof Operator) =>
    setOpSort((s) => ({ key, dir: s.key === key && s.dir === 'asc' ? 'desc' : 'asc' }));

  const handleCtSort = (key: keyof Contract) =>
    setCtSort((s) => ({ key, dir: s.key === key && s.dir === 'asc' ? 'desc' : 'asc' }));

  const statusClass = (status: string) => {
    if (status === 'Active') return classes.statusActive;
    if (status === 'Pending') return classes.statusPending;
    if (status === 'Expired') return classes.statusExpired;
    return classes.statusInactive;
  };

  const toggleInsightOp = (id: number) =>
    setExpandedOps((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });

  return (
    <div className={classes.root}>
      <Typography className={classes.pageTitle}>Contacts &amp; Contracts</Typography>

      <Paper className={classes.card} elevation={0}>
        <Tabs value={tab} onChange={(_, v) => setTab(v)} className={classes.tabBar}>
          <Tab label="Contacts" />
          <Tab label="Contracts &amp; Promotions" />
          <Tab label="Insights" />
        </Tabs>

        {/* ── Contacts tab ── */}
        {tab === 0 && (
          <>
            <div className={classes.toolbar}>
              <TextField
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search operators…"
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
              <span className={classes.spacer} />
              <Button className={classes.primaryBtn} variant="contained" disableElevation onClick={() => setNewOpOpen(true)}>
                <Icon>add</Icon>
                New Operator
              </Button>
            </div>

            <Table className={classes.table} size="small">
              <TableHead>
                <TableRow>
                  {(['name', 'contact', 'country', 'status', 'revenue', 'bookings', 'closureRate'] as (keyof Operator)[]).map((col) => (
                    <TableCell key={col}>
                      <TableSortLabel
                        active={opSort.key === col}
                        direction={opSort.key === col ? opSort.dir : 'asc'}
                        onClick={() => handleOpSort(col)}
                      >
                        {col === 'closureRate' ? 'Closure Rate' : col.charAt(0).toUpperCase() + col.slice(1)}
                      </TableSortLabel>
                    </TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {filteredOps.map((op) => (
                  <TableRow
                    key={op.id}
                    onClick={() => setSelectedOp(op.id === selectedOp?.id ? null : op)}
                    className={clsx(selectedOp?.id === op.id && classes.selectedRow)}
                  >
                    <TableCell>{op.name}</TableCell>
                    <TableCell>{op.contact}</TableCell>
                    <TableCell>{op.country}</TableCell>
                    <TableCell>
                      <Chip label={op.status} className={clsx(classes.statusChip, statusClass(op.status))} />
                    </TableCell>
                    <TableCell>{op.revenue}</TableCell>
                    <TableCell>{op.bookings.toLocaleString()}</TableCell>
                    <TableCell>{op.closureRate}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>

            {/* Detail panel */}
            {selectedOp && (
              <div className={classes.detailPanel}>
                <div className={classes.detailHeader}>
                  <div>
                    <Typography className={classes.detailName}>{selectedOp.name}</Typography>
                    <Typography className={classes.detailContact}>
                      {selectedOp.contact} · {selectedOp.country}
                    </Typography>
                  </div>
                  <Chip label={selectedOp.status} className={clsx(classes.statusChip, statusClass(selectedOp.status))} />
                </div>
                <div className={classes.statsRow}>
                  {[
                    { label: 'Total Revenue', value: selectedOp.revenue },
                    { label: 'Total Bookings', value: selectedOp.bookings.toLocaleString() },
                    { label: 'Closure Rate', value: selectedOp.closureRate },
                  ].map(({ label, value }) => (
                    <div key={label} className={classes.statCard}>
                      <Typography className={classes.statValue}>{value}</Typography>
                      <Typography className={classes.statLabel}>{label}</Typography>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </>
        )}

        {/* ── Contracts & Promotions tab ── */}
        {tab === 1 && (
          <>
            <div className={classes.toolbar}>
              <TextField
                value={contractSearch}
                onChange={(e) => setContractSearch(e.target.value)}
                placeholder="Search contracts…"
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
              <span className={classes.spacer} />
              <Button className={classes.actionBtn}>
                <Icon>filter_list</Icon>
                Filters
              </Button>
              <Button className={classes.primaryBtn} variant="contained" disableElevation>
                <Icon>add</Icon>
                New Contract
              </Button>
            </div>

            <Table className={classes.table} size="small">
              <TableHead>
                <TableRow>
                  {(['operator', 'type', 'season', 'rooms', 'rate', 'status', 'signed'] as (keyof Contract)[]).map((col) => (
                    <TableCell key={col}>
                      <TableSortLabel
                        active={ctSort.key === col}
                        direction={ctSort.key === col ? ctSort.dir : 'asc'}
                        onClick={() => handleCtSort(col)}
                      >
                        {col.charAt(0).toUpperCase() + col.slice(1)}
                      </TableSortLabel>
                    </TableCell>
                  ))}
                  <TableCell>Actions</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {filteredContracts.map((c) => (
                  <TableRow key={c.id}>
                    <TableCell>{c.operator}</TableCell>
                    <TableCell>{c.type}</TableCell>
                    <TableCell>{c.season}</TableCell>
                    <TableCell>{c.rooms}</TableCell>
                    <TableCell>{c.rate}</TableCell>
                    <TableCell>
                      <Chip label={c.status} className={clsx(classes.statusChip, statusClass(c.status))} />
                    </TableCell>
                    <TableCell>{c.signed}</TableCell>
                    <TableCell>
                      <IconButton size="small" style={{ color: '#6b7280' }}>
                        <Icon style={{ fontSize: 16 }}>visibility</Icon>
                      </IconButton>
                      <IconButton size="small" style={{ color: '#6b7280' }}>
                        <Icon style={{ fontSize: 16 }}>content_copy</Icon>
                      </IconButton>
                      {c.status === 'Pending' && (
                        <IconButton size="small" style={{ color: '#15803d' }}>
                          <Icon style={{ fontSize: 16 }}>send</Icon>
                        </IconButton>
                      )}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </>
        )}

        {/* ── Insights tab ── */}
        {tab === 2 && (
          <>
            <div className={classes.insightsToolbar}>
              <div className={classes.viewToggle}>
                {(['to', 'contract', 'overall'] as const).map((v) => (
                  <Button
                    key={v}
                    className={clsx(classes.viewToggleBtn, insightView === v && classes.viewToggleBtnActive)}
                    onClick={() => setInsightView(v)}
                    disableElevation
                    variant={insightView === v ? 'contained' : 'text'}
                  >
                    {v === 'to' ? 'Operator' : v === 'contract' ? 'Contract' : 'Overall'}
                  </Button>
                ))}
              </div>
              <span className={classes.spacer} />
              <Button className={classes.actionBtn}>
                <Icon>expand</Icon>
                Expand All
              </Button>
              <Button className={classes.actionBtn}>
                <Icon>compress</Icon>
                Collapse All
              </Button>
            </div>

            {OPERATORS.filter((o) => o.status === 'Active').map((op) => (
              <div key={op.id} className={classes.insightBlock}>
                <div className={classes.insightHeader} onClick={() => toggleInsightOp(op.id)}>
                  <Typography className={classes.insightHeaderLabel}>{op.name}</Typography>
                  <Chip label={op.status} className={clsx(classes.statusChip, classes.statusActive)} style={{ marginRight: 8 }} />
                  <Icon style={{ fontSize: 18, color: '#9ca3af' }}>
                    {expandedOps.has(op.id) ? 'expand_less' : 'expand_more'}
                  </Icon>
                </div>
                {expandedOps.has(op.id) && (
                  <div className={classes.insightBody}>
                    {[
                      { label: 'Revenue', value: op.revenue, pct: 0.72 },
                      { label: 'Bookings', value: op.bookings.toLocaleString(), pct: 0.58 },
                      { label: 'Closure Rate', value: op.closureRate, pct: 0.35 },
                      { label: 'Avg Lead Time', value: '28 days', pct: 0.45 },
                    ].map(({ label, value, pct }) => (
                      <div key={label} className={classes.insightMetricRow}>
                        <Typography className={classes.insightMetricLabel}>{label}</Typography>
                        <div className={classes.progressBar}>
                          <div className={classes.progressFill} style={{ width: `${pct * 100}%` }} />
                        </div>
                        <Typography className={classes.insightMetricValue}>{value}</Typography>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </>
        )}
      </Paper>

      {/* New Operator dialog */}
      <Dialog open={newOpOpen} onClose={() => setNewOpOpen(false)} maxWidth="sm" fullWidth>
        <DialogTitle style={{ fontFamily: 'Lato, sans-serif', fontWeight: 700 }}>New Operator</DialogTitle>
        <DialogContent>
          {['Name', 'Contact', 'Phone', 'Country'].map((field) => (
            <TextField
              key={field}
              label={field}
              variant="outlined"
              size="small"
              fullWidth
              style={{ marginBottom: 12, marginTop: 4 }}
              InputProps={{
                style: { fontFamily: 'Lato, sans-serif', fontSize: 13 },
              }}
              InputLabelProps={{ style: { fontFamily: 'Lato, sans-serif', fontSize: 13 } }}
            />
          ))}
        </DialogContent>
        <DialogActions style={{ padding: '12px 24px' }}>
          <Button onClick={() => setNewOpOpen(false)} style={{ fontFamily: 'Lato, sans-serif', textTransform: 'none', fontSize: 13 }}>
            Cancel
          </Button>
          <Button
            onClick={() => setNewOpOpen(false)}
            variant="contained"
            disableElevation
            style={{ backgroundColor: '#006461', color: '#fff', fontFamily: 'Lato, sans-serif', textTransform: 'none', fontSize: 13 }}
          >
            Save Operator
          </Button>
        </DialogActions>
      </Dialog>
    </div>
  );
}
