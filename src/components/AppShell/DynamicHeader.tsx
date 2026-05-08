import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Badge from '@material-ui/core/Badge';
import clsx from 'clsx';

export type PageId = 0 | 1 | 2 | 3 | 4;
export type BreadcrumbOption = { title: string; url?: string | false };

const TOP_NAV_TABS = [
  'Home',
  'Advance',
  'Pricing & Strategy',
  'Forecasts',
  'Reports',
  'Travel Distribution Hub',
];

type Props = {
  activePage: PageId;
  onNavigate: (page: PageId) => void;
  propertyName?: string;
};

const useStyles = makeStyles((theme) => ({
  root: {
    width: '100%',
    height: 80,
    flexShrink: 0,
    backgroundColor: theme.palette.common.white,
    borderBottom: '1px solid #e0e4e6',
    boxShadow: '0px 2px 6px 0px rgba(0,0,0,0.06)',
    display: 'flex',
    alignItems: 'stretch',
    position: 'sticky',
    top: 0,
    zIndex: theme.zIndex.appBar,
    boxSizing: 'border-box',
  },

  // Logo box
  logo: {
    width: 110,
    height: 32,
    backgroundColor: theme.palette.primary.main,
    borderRadius: 4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
    alignSelf: 'center',
    marginLeft: 23,
    marginRight: 20,
    cursor: 'default',
  },
  logoText: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 13,
    color: theme.palette.common.white,
    letterSpacing: 0.5,
    userSelect: 'none',
  },

  // Nav tabs
  navTabs: {
    display: 'flex',
    alignItems: 'stretch',
    flex: 1,
  },
  navTab: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(0, 1.5),
    position: 'relative',
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 400,
    color: '#6b7280',
    whiteSpace: 'nowrap',
    borderBottom: '2px solid transparent',
    userSelect: 'none',
    '&:hover': {
      color: theme.palette.text.primary,
    },
  },
  navTabActive: {
    color: theme.palette.primary.main,
    fontWeight: 600,
    borderBottom: `2px solid ${theme.palette.primary.main}`,
  },

  // Right icons
  icons: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.5),
    paddingRight: theme.spacing(2),
    flexShrink: 0,
  },
  iconBtn: {
    color: '#6b7280',
    padding: 6,
    '&:hover': { color: theme.palette.text.primary },
    '& .MuiIcon-root': { fontSize: '20px !important' },
  },
  avatar: {
    width: 32,
    height: 32,
    borderRadius: '50%',
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 700,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    flexShrink: 0,
    marginLeft: theme.spacing(0.5),
  },
}));

export function DynamicHeader({ activePage, onNavigate, propertyName = 'Hotel Las Américas' }: Props) {
  const classes = useStyles();

  return (
    <div className={classes.root}>
      {/* Logo */}
      <div className={classes.logo}>
        <span className={classes.logoText}>duetto</span>
      </div>

      {/* Top nav tabs */}
      <div className={classes.navTabs}>
        {TOP_NAV_TABS.map((tab) => {
          const isTDH = tab === 'Travel Distribution Hub';
          return (
            <div
              key={tab}
              className={clsx(classes.navTab, isTDH && classes.navTabActive)}
            >
              {tab}
            </div>
          );
        })}
      </div>

      {/* Right icons */}
      <div className={classes.icons}>
        <IconButton className={classes.iconBtn} size="small">
          <Badge badgeContent="99+" color="error" style={{ fontSize: 9 }}>
            <Icon>notifications</Icon>
          </Badge>
        </IconButton>
        <IconButton className={classes.iconBtn} size="small">
          <Icon>help_outline</Icon>
        </IconButton>
        <IconButton className={classes.iconBtn} size="small">
          <Icon>settings</Icon>
        </IconButton>
        <div className={classes.avatar}>M</div>
      </div>
    </div>
  );
}
