import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Paper from '@material-ui/core/Paper';
import Typography from '@material-ui/core/Typography';
import Icon from '@material-ui/core/Icon';
import clsx from 'clsx';
import { DayDetailGroup } from './types';
import { calendarTokens } from './tokens';

const useStyles = makeStyles((theme) => ({
  root: {
    position: 'absolute',
    width: 350,
    height: 638,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    borderRadius: 4,
    zIndex: theme.zIndex.modal,
  },
  header: {
    height: 48,
    flexShrink: 0,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: theme.spacing(0, 2),
    backgroundColor: theme.palette.primary.main,
  },
  headerTitle: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 14,
    color: theme.palette.common.white,
  },
  closeBtn: {
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    '& .MuiIcon-root': {
      fontSize: '20px !important',
      color: theme.palette.common.white,
    },
  },
  body: {
    flex: 1,
    overflowY: 'auto',
    backgroundColor: theme.palette.common.white,
  },
  sectionRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: 44,
    padding: theme.spacing(0, 2),
    borderBottom: `1px solid ${calendarTokens.border}`,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: calendarTokens.cellBackground,
    },
  },
  sectionTitle: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 13,
    color: theme.palette.text.primary,
  },
  expandIcon: {
    fontSize: '20px !important',
    color: theme.palette.text.secondary,
  },
  itemRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: 32,
    padding: theme.spacing(0, 2),
    borderBottom: `1px solid ${calendarTokens.border}`,
  },
  itemLabel: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.text.secondary,
  },
  itemRight: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
  },
  itemPercentage: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    color: theme.palette.text.disabled,
  },
  itemValue: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 600,
    color: theme.palette.text.primary,
  },
  itemNegative: {
    color: theme.palette.error.main,
  },
}));

type Props = {
  date: string;
  groups: DayDetailGroup[];
  onClose?: () => void;
  style?: React.CSSProperties;
  onToggleGroup?: (index: number) => void;
};

export function DayDetailPopup({ date, groups, onClose, style, onToggleGroup }: Props) {
  const classes = useStyles();

  return (
    <Paper className={classes.root} style={style} elevation={4}>
      <div className={classes.header}>
        <Typography className={classes.headerTitle}>{date}</Typography>
        <div className={classes.closeBtn} onClick={onClose}>
          <Icon>close</Icon>
        </div>
      </div>

      <div className={classes.body}>
        {groups.map((group, gi) => (
          <React.Fragment key={gi}>
            <div className={classes.sectionRow} onClick={() => onToggleGroup?.(gi)}>
              <Typography className={classes.sectionTitle}>{group.title}</Typography>
              <Icon className={classes.expandIcon}>
                {group.isExpanded ? 'expand_less' : 'expand_more'}
              </Icon>
            </div>

            {group.isExpanded && group.items.map((item, ii) => (
              <div key={ii} className={classes.itemRow}>
                <Typography className={classes.itemLabel}>{item.label}</Typography>
                <div className={classes.itemRight}>
                  {item.percentage != null && (
                    <span className={classes.itemPercentage}>{item.percentage}%</span>
                  )}
                  {item.seats != null && (
                    <span className={classes.itemPercentage}>· {item.seats} seats</span>
                  )}
                  {item.value != null && (
                    <Typography
                      className={clsx(classes.itemValue, item.isNegative && classes.itemNegative)}
                    >
                      {item.value}
                    </Typography>
                  )}
                </div>
              </div>
            ))}
          </React.Fragment>
        ))}
      </div>
    </Paper>
  );
}
