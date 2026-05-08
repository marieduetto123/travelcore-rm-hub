import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import clsx from 'clsx';

export type PageId = 0 | 1 | 2 | 3 | 4;
export type BreadcrumbOption = { title: string; url?: string | false };

const TOP_NAV = [
  { label: 'Home', hasChevron: false },
  { label: 'Advance', hasChevron: false },
  { label: 'Pricing & Strategy', hasChevron: true },
  { label: 'Forecasts & Budgets', hasChevron: true },
  { label: 'Reports', hasChevron: true },
  { label: 'Groups', hasChevron: true },
  { label: 'Travel Distribution Hub', hasChevron: true, isActive: true },
];

type Props = {
  activePage: PageId;
  onNavigate: (page: PageId) => void;
  propertyName?: string;
};

const useStyles = makeStyles((theme) => ({
  root: {
    width: '100%',
    flexShrink: 0,
    position: 'sticky',
    top: 0,
    zIndex: theme.zIndex.appBar,
  },

  // ── Top bar ────────────────────────────────────────────────────────
  topBar: {
    height: 40,
    backgroundColor: theme.palette.secondary.main, // #0E2124
    display: 'flex',
    alignItems: 'stretch',
    paddingLeft: 24,
    paddingRight: 14,
    boxSizing: 'border-box',
  },
  logo: {
    display: 'flex',
    alignItems: 'center',
    paddingRight: 32,
    flexShrink: 0,
  },
  logoText: {
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    fontSize: 16,
    color: theme.palette.common.white,
    letterSpacing: 1,
    userSelect: 'none',
  },

  // Nav tabs
  nav: {
    display: 'flex',
    alignItems: 'stretch',
    flex: 1,
  },
  tab: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 2,
    padding: theme.spacing(0, 2),
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 400,
    color: theme.palette.common.white,
    whiteSpace: 'nowrap',
    userSelect: 'none',
    '&:hover': {
      backgroundColor: 'rgba(255,255,255,0.08)',
    },
  },
  tabActive: {
    backgroundColor: '#c4ff45',
    color: theme.palette.secondary.main, // #0E2124
    '&:hover': {
      backgroundColor: '#bdf03e',
    },
  },
  tabChevron: {
    fontSize: '16px !important',
    lineHeight: '16px',
  },

  // Right icons
  icons: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flexShrink: 0,
  },
  iconBtn: {
    width: 32,
    height: 32,
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    color: 'rgba(255,255,255,0.8)',
    '&:hover': {
      backgroundColor: 'rgba(255,255,255,0.1)',
    },
    '& .MuiIcon-root': { fontSize: '20px !important' },
  },
  notifWrap: {
    position: 'relative',
  },
  notifBadge: {
    position: 'absolute',
    top: 1,
    right: -1,
    backgroundColor: theme.palette.error.main,
    color: theme.palette.common.white,
    fontSize: 7,
    fontFamily: 'Lato, sans-serif',
    fontWeight: 700,
    height: 14,
    minWidth: 14,
    padding: '0 2px',
    borderRadius: 4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    lineHeight: 1,
  },
  avatar: {
    width: 32,
    height: 32,
    borderRadius: 16,
    backgroundColor: '#f59e0b', // orange
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    flexShrink: 0,
  },

  // ── Breadcrumb bar ─────────────────────────────────────────────────
  breadcrumbBar: {
    height: 32,
    backgroundColor: '#f8fafc',
    borderBottom: '1px solid #dde1e2',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingLeft: 24,
    boxSizing: 'border-box',
  },
  breadcrumbs: {
    display: 'flex',
    alignItems: 'center',
  },
  crumbLink: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.primary.main, // #006461
    cursor: 'pointer',
    '&:hover': { textDecoration: 'underline' },
  },
  crumbSep: {
    fontSize: '12px !important',
    color: '#585858',
    margin: '0 4px',
    lineHeight: '12px',
  },
  crumbCurrent: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: '#4f5b60',
    cursor: 'default',
  },

  // Property picker (right side of breadcrumb bar)
  propertyPicker: {
    height: '100%',
    minWidth: 200,
    backgroundColor: theme.palette.common.white,
    borderLeft: '1px solid #dde1e2',
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    paddingLeft: 13,
    paddingRight: 12,
    paddingBottom: 1,
    boxSizing: 'border-box',
    cursor: 'pointer',
  },
  propertyIcon: {
    fontSize: '20px !important',
    color: '#19393e',
  },
  propertyName: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    color: theme.palette.primary.main, // #006461
    flex: 1,
    whiteSpace: 'nowrap',
  },
  propertyChevron: {
    fontSize: '16px !important',
    color: '#19393e',
  },
}));

export function DynamicHeader({ activePage, onNavigate, propertyName = 'Hotel Sevilla' }: Props) {
  const classes = useStyles();

  return (
    <div className={classes.root}>
      {/* Top bar */}
      <div className={classes.topBar}>
        {/* Logo */}
        <div className={classes.logo}>
          <span className={classes.logoText}>duetto</span>
        </div>

        {/* Nav tabs */}
        <div className={classes.nav}>
          {TOP_NAV.map((tab) => (
            <div
              key={tab.label}
              className={clsx(classes.tab, tab.isActive && classes.tabActive)}
            >
              {tab.label}
              {tab.hasChevron && (
                <Icon className={classes.tabChevron}>expand_more</Icon>
              )}
            </div>
          ))}
        </div>

        {/* Right icons */}
        <div className={classes.icons}>
          <div className={classes.iconBtn}>
            <Icon>dark_mode</Icon>
          </div>
          <div className={clsx(classes.iconBtn, classes.notifWrap)}>
            <Icon>notifications</Icon>
            <span className={classes.notifBadge}>99+</span>
          </div>
          <div className={classes.iconBtn}>
            <Icon>help</Icon>
          </div>
          <div className={classes.iconBtn}>
            <Icon>settings</Icon>
          </div>
          <div className={classes.avatar}>M</div>
        </div>
      </div>

      {/* Breadcrumb bar */}
      <div className={classes.breadcrumbBar}>
        <div className={classes.breadcrumbs}>
          <span className={classes.crumbLink}>Home</span>
          <Icon className={classes.crumbSep}>chevron_right</Icon>
          <span className={classes.crumbCurrent}>Travel Distribution Hubs</span>
        </div>

        <div className={classes.propertyPicker}>
          <Icon className={classes.propertyIcon}>home</Icon>
          <span className={classes.propertyName}>{propertyName}</span>
          <Icon className={classes.propertyChevron}>expand_more</Icon>
        </div>
      </div>
    </div>
  );
}
