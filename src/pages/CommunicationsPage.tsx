import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Button from '@material-ui/core/Button';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Paper from '@material-ui/core/Paper';
import Table from '@material-ui/core/Table';
import TableHead from '@material-ui/core/TableHead';
import TableBody from '@material-ui/core/TableBody';
import TableRow from '@material-ui/core/TableRow';
import TableCell from '@material-ui/core/TableCell';
import Chip from '@material-ui/core/Chip';
import Dialog from '@material-ui/core/Dialog';
import DialogTitle from '@material-ui/core/DialogTitle';
import DialogContent from '@material-ui/core/DialogContent';
import DialogActions from '@material-ui/core/DialogActions';
import TextField from '@material-ui/core/TextField';
import Select from '@material-ui/core/Select';
import MenuItem from '@material-ui/core/MenuItem';
import Snackbar from '@material-ui/core/Snackbar';
import clsx from 'clsx';

const tokens = { border: '#dde1e2', rowHover: 'rgba(0,100,97,0.04)', sectionBg: '#f8f9fa' };

type ActionItem = { id: number; title: string; operator: string; priority: 'High' | 'Medium' | 'Low'; due: string };
type CommRecord = {
  id: number;
  date: string;
  operator: string;
  subject: string;
  type: string;
  channel: 'External' | 'Internal';
  user: string;
};

const ACTION_ITEMS: ActionItem[] = [
  { id: 1, title: 'Review stop-sale override request', operator: 'TUI Group', priority: 'High', due: '2026-05-10' },
  { id: 2, title: 'Confirm Summer 2026 allotment release', operator: 'Thomas Cook', priority: 'Medium', due: '2026-05-15' },
];

const COMMS: CommRecord[] = [
  { id: 1, date: '2026-05-07', operator: 'TUI Group', subject: 'Rate amendment request Q3 2026', type: 'Email', channel: 'External', user: 'M. Dare' },
  { id: 2, date: '2026-05-06', operator: 'Club Med', subject: 'Allotment release confirmation', type: 'Note', channel: 'Internal', user: 'M. Dare' },
  { id: 3, date: '2026-05-05', operator: 'Jet2holidays', subject: 'Stop-sale request received', type: 'Email', channel: 'External', user: 'System' },
  { id: 4, date: '2026-05-04', operator: 'FTI Group', subject: 'Contract addendum for winter season', type: 'Email', channel: 'External', user: 'M. Dare' },
  { id: 5, date: '2026-05-03', operator: 'Thomas Cook', subject: 'Internal review: occupancy concern', type: 'Note', channel: 'Internal', user: 'M. Dare' },
];

const useStyles = makeStyles((theme) => ({
  root: { padding: theme.spacing(3), display: 'flex', flexDirection: 'column', gap: theme.spacing(2.5) },
  pageTitle: { fontFamily: 'Lato, sans-serif', fontSize: 20, fontWeight: 700, color: theme.palette.text.primary },
  card: { border: `1px solid ${tokens.border}`, borderRadius: 4, backgroundColor: theme.palette.common.white, overflow: 'hidden' },
  cardHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
  },
  cardTitle: { fontFamily: 'Lato, sans-serif', fontSize: 14, fontWeight: 700, color: theme.palette.text.primary },
  actionItemRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: theme.spacing(2),
    padding: theme.spacing(1.5, 2),
    borderBottom: `1px solid ${tokens.border}`,
    '&:last-child': { borderBottom: 'none' },
  },
  priorityDot: { width: 8, height: 8, borderRadius: '50%', marginTop: 6, flexShrink: 0 },
  actionTitle: { fontFamily: 'Lato, sans-serif', fontSize: 13, fontWeight: 600, color: theme.palette.text.primary },
  actionMeta: { fontFamily: 'Lato, sans-serif', fontSize: 12, color: theme.palette.text.secondary, marginTop: 2 },
  actionDue: { fontFamily: 'Lato, sans-serif', fontSize: 12, color: theme.palette.text.disabled },
  takeActionBtn: {
    marginLeft: 'auto',
    height: 30,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    textTransform: 'none',
    color: theme.palette.primary.main,
    border: `1px solid ${theme.palette.primary.main}`,
    padding: theme.spacing(0, 1.5),
    flexShrink: 0,
    '&:hover': { backgroundColor: 'rgba(0,100,97,0.06)' },
  },
  channelTabs: {
    display: 'flex',
    borderBottom: `1px solid ${tokens.border}`,
    paddingLeft: theme.spacing(2),
  },
  channelTab: {
    padding: theme.spacing(1, 2),
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
    cursor: 'pointer',
    borderBottom: '2px solid transparent',
    '&:hover': { color: theme.palette.primary.main },
  },
  channelTabActive: { color: theme.palette.primary.main, fontWeight: 700, borderBottom: `2px solid ${theme.palette.primary.main}` },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    padding: theme.spacing(1, 2),
    borderBottom: `1px solid ${tokens.border}`,
  },
  spacer: { flex: 1 },
  addNoteBtn: {
    height: 32,
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    textTransform: 'none',
    padding: theme.spacing(0, 1.5),
    '&:hover': { backgroundColor: theme.palette.primary.dark },
    '& .MuiIcon-root': { fontSize: '16px !important', marginRight: theme.spacing(0.5) },
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
  typeChip: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    height: 20,
    '& .MuiChip-label': { padding: '0 8px' },
  },
  typeEmail: { backgroundColor: '#dbeafe', color: '#1d4ed8' },
  typeNote: { backgroundColor: '#f3f4f6', color: '#4b5563' },
}));

export function CommunicationsPage() {
  const classes = useStyles();
  const [channel, setChannel] = useState<'External' | 'Internal'>('External');
  const [noteOpen, setNoteOpen] = useState(false);
  const [noteSubject, setNoteSubject] = useState('');
  const [noteBody, setNoteBody] = useState('');
  const [noteOperator, setNoteOperator] = useState('');
  const [snackbar, setSnackbar] = useState('');
  const [dismissedActions, setDismissedActions] = useState<Set<number>>(new Set());

  const visibleComms = COMMS.filter((c) => c.channel === channel);

  const handleTakeAction = (item: ActionItem) => {
    setSnackbar(`Action taken on: ${item.title}`);
    setDismissedActions((d) => new Set([...d, item.id]));
  };

  const handleAddNote = () => {
    if (!noteSubject.trim()) return;
    setSnackbar('Note added successfully');
    setNoteOpen(false);
    setNoteSubject('');
    setNoteBody('');
    setNoteOperator('');
  };

  const activeActions = ACTION_ITEMS.filter((a) => !dismissedActions.has(a.id));

  return (
    <div className={classes.root}>
      <Typography className={classes.pageTitle}>Communications &amp; Notes</Typography>

      {/* Action items */}
      {activeActions.length > 0 && (
        <Paper className={classes.card} elevation={0}>
          <div className={classes.cardHeader}>
            <Typography className={classes.cardTitle}>
              Action Required
              <span style={{ marginLeft: 8, backgroundColor: '#dc2626', color: '#fff', borderRadius: '50%', width: 18, height: 18, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700 }}>
                {activeActions.length}
              </span>
            </Typography>
          </div>
          {activeActions.map((item) => (
            <div key={item.id} className={classes.actionItemRow}>
              <div
                className={classes.priorityDot}
                style={{
                  backgroundColor: item.priority === 'High' ? '#dc2626' : item.priority === 'Medium' ? '#f59e0b' : '#6b7280',
                }}
              />
              <div style={{ flex: 1 }}>
                <Typography className={classes.actionTitle}>{item.title}</Typography>
                <Typography className={classes.actionMeta}>{item.operator}</Typography>
                <Typography className={classes.actionDue}>Due: {item.due}</Typography>
              </div>
              <Button
                variant="outlined"
                className={classes.takeActionBtn}
                onClick={() => handleTakeAction(item)}
              >
                Take Action
              </Button>
            </div>
          ))}
        </Paper>
      )}

      {/* Communications log */}
      <Paper className={classes.card} elevation={0}>
        <div className={classes.channelTabs}>
          {(['External', 'Internal'] as const).map((ch) => (
            <div
              key={ch}
              className={clsx(classes.channelTab, channel === ch && classes.channelTabActive)}
              onClick={() => setChannel(ch)}
            >
              {ch}
            </div>
          ))}
        </div>

        <div className={classes.toolbar}>
          <Typography style={{ fontFamily: 'Lato, sans-serif', fontSize: 13, color: '#6b7280' }}>
            {visibleComms.length} record{visibleComms.length !== 1 ? 's' : ''}
          </Typography>
          <span className={classes.spacer} />
          <Button
            className={classes.addNoteBtn}
            variant="contained"
            disableElevation
            onClick={() => setNoteOpen(true)}
          >
            <Icon>add</Icon>
            Add Note
          </Button>
        </div>

        <Table className={classes.table} size="small">
          <TableHead>
            <TableRow>
              <TableCell>Date</TableCell>
              <TableCell>Operator</TableCell>
              <TableCell>Subject</TableCell>
              <TableCell>Type</TableCell>
              <TableCell>User</TableCell>
              <TableCell />
            </TableRow>
          </TableHead>
          <TableBody>
            {visibleComms.map((rec) => (
              <TableRow key={rec.id}>
                <TableCell>{rec.date}</TableCell>
                <TableCell>{rec.operator}</TableCell>
                <TableCell>{rec.subject}</TableCell>
                <TableCell>
                  <Chip
                    label={rec.type}
                    className={clsx(classes.typeChip, rec.type === 'Email' ? classes.typeEmail : classes.typeNote)}
                  />
                </TableCell>
                <TableCell>{rec.user}</TableCell>
                <TableCell>
                  <IconButton size="small" style={{ color: '#9ca3af' }}>
                    <Icon style={{ fontSize: 16 }}>more_vert</Icon>
                  </IconButton>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </Paper>

      {/* Add note dialog */}
      <Dialog open={noteOpen} onClose={() => setNoteOpen(false)} maxWidth="sm" fullWidth>
        <DialogTitle style={{ fontFamily: 'Lato, sans-serif', fontWeight: 700 }}>Add Note</DialogTitle>
        <DialogContent style={{ paddingTop: 8 }}>
          <Select
            value={noteOperator}
            onChange={(e) => setNoteOperator(e.target.value as string)}
            displayEmpty
            variant="outlined"
            fullWidth
            style={{ marginBottom: 12, fontFamily: 'Lato, sans-serif', fontSize: 13 }}
          >
            <MenuItem value="" style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>
              <em>Select Operator</em>
            </MenuItem>
            {['TUI Group', 'Thomas Cook', 'Jet2holidays', 'FTI Group', 'Sunwing', 'Club Med'].map((op) => (
              <MenuItem key={op} value={op} style={{ fontFamily: 'Lato, sans-serif', fontSize: 13 }}>{op}</MenuItem>
            ))}
          </Select>
          <TextField
            value={noteSubject}
            onChange={(e) => setNoteSubject(e.target.value)}
            label="Subject"
            variant="outlined"
            size="small"
            fullWidth
            style={{ marginBottom: 12 }}
            InputProps={{ style: { fontFamily: 'Lato, sans-serif', fontSize: 13 } }}
            InputLabelProps={{ style: { fontFamily: 'Lato, sans-serif', fontSize: 13 } }}
          />
          <TextField
            value={noteBody}
            onChange={(e) => setNoteBody(e.target.value)}
            label="Note"
            variant="outlined"
            multiline
            rows={4}
            fullWidth
            InputProps={{ style: { fontFamily: 'Lato, sans-serif', fontSize: 13 } }}
            InputLabelProps={{ style: { fontFamily: 'Lato, sans-serif', fontSize: 13 } }}
          />
        </DialogContent>
        <DialogActions style={{ padding: '12px 24px' }}>
          <Button
            onClick={() => setNoteOpen(false)}
            style={{ fontFamily: 'Lato, sans-serif', textTransform: 'none', fontSize: 13 }}
          >
            Cancel
          </Button>
          <Button
            onClick={handleAddNote}
            variant="contained"
            disableElevation
            style={{ backgroundColor: '#006461', color: '#fff', fontFamily: 'Lato, sans-serif', textTransform: 'none', fontSize: 13 }}
          >
            Save Note
          </Button>
        </DialogActions>
      </Dialog>

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
