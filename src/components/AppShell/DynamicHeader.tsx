import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Typography from '@material-ui/core/Typography';
import Icon from '@material-ui/core/Icon';
import IconButton from '@material-ui/core/IconButton';
import Badge from '@material-ui/core/Badge';
import Paper from '@material-ui/core/Paper';
import clsx from 'clsx';

export type PageId = 0 | 1 | 2 | 3 | 4;
export type BreadcrumbOption = { title: string; url?: string | false };

const PAGE_LABELS: { id: PageId; label: string }[] = [
  { id: 0, label: 'Dashboard' },
  { id: 1, label: 'Contacts & Contracts' },
  { id: 2, label: 'Audit' },
  { id: 3, label: 'Communications & Notes' },
  { id: 4, label: 'Configuration' },
];

const TOP_NAV_TABS = ['Sales', 'Rates', 'Revenue', 'Travel Distribution Hub', 'Reports', 'Settings'];

type Props = {
  activePage: PageId;
  onNavigate: (page: PageId) => void;
  breadcrumbs?: BreadcrumbOption[];
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

  // ── Top bar ────────────────────────────────────────────────────────────────
  topBar: {
    height: 40,
    backgroundColor: theme.palette.secondary.main,
    display: 'flex',
    alignItems: 'stretch',
    paddingLeft: theme.spacing(2),
    paddingRight: theme.spacing(1),
  },
  logo: {
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 15,
    fontWeight: 700,
    letterSpacing: 1,
    display: 'flex',
    alignItems: 'center',
    marginRight: theme.spacing(3),
    flexShrink: 0,
    userSelect: 'none',
  },
  topNavTabs: {
    display: 'flex',
    alignItems: 'stretch',
    flex: 1,
  },
  topNavTab: {
    display: 'flex',
    alignItems: 'center',
    padding: theme.spacing(0, 1.75),
    cursor: 'pointer',
    color: 'rgba(255,255,255,0.55)',
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 400,
    whiteSpace: 'nowrap',
    borderBottom: '2px solid transparent',
    userSelect: 'none',
    '&:hover': {
      color: theme.palette.common.white,
      backgroundColor: 'rgba(255,255,255,0.05)',
    },
  },
  topNavTabActive: {
    color: theme.palette.common.white,
    fontWeight: 700,
    borderBottom: `2px solid ${theme.palette.primary.main}`,
  },
  utilities: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.25),
    marginLeft: 'auto',
    flexShrink: 0,
  },
  utilBtn: {
    color: 'rgba(255,255,255,0.65)',
    padding: 6,
    '&:hover': { color: theme.palette.common.white },
    '& .MuiIcon-root': { fontSize: '18px !important' },
  },
  avatar: {
    width: 26,
    height: 26,
    borderRadius: '50%',
    backgroundColor: theme.palette.primary.main,
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    fontWeight: 700,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    flexShrink: 0,
    marginLeft: theme.spacing(0.5),
  },

  // ── Mega menu ──────────────────────────────────────────────────────────────
  megaMenuWrap: {
    position: 'fixed',
    top: 40,
    left: 0,
    right: 0,
    zIndex: theme.zIndex.appBar - 1,
  },
  megaMenu: {
    backgroundColor: theme.palette.secondary.main,
    borderTop: '1px solid rgba(255,255,255,0.08)',
    paddingTop: theme.spacing(1.5),
    paddingBottom: theme.spacing(1.5),
    paddingLeft: theme.spacing(20),
    display: 'flex',
    gap: theme.spacing(0.5),
  },
  megaItem: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1),
    padding: theme.spacing(1, 2),
    borderRadius: 4,
    cursor: 'pointer',
    color: 'rgba(255,255,255,0.75)',
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    '& .MuiIcon-root': { fontSize: '16px !important', opacity: 0.7 },
    '&:hover': {
      backgroundColor: 'rgba(255,255,255,0.08)',
      color: theme.palette.common.white,
    },
  },
  megaItemActive: {
    backgroundColor: 'rgba(255,255,255,0.1)',
    color: theme.palette.common.white,
    fontWeight: 700,
  },

  // ── Breadcrumb bar ─────────────────────────────────────────────────────────
  breadcrumbBar: {
    height: 32,
    backgroundColor: theme.palette.background.default,
    borderBottom: `1px solid ${theme.palette.divider}`,
    display: 'flex',
    alignItems: 'center',
    paddingLeft: theme.spacing(2.5),
    paddingRight: theme.spacing(2.5),
    gap: theme.spacing(0.25),
  },
  breadcrumbItem: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: theme.palette.text.secondary,
    cursor: 'pointer',
    padding: theme.spacing(0, 0.5),
    '&:hover': {
      color: theme.palette.primary.main,
      textDecoration: 'underline',
    },
  },
  breadcrumbCurrent: {
    color: theme.palette.text.primary,
    fontWeight: 600,
    cursor: 'default',
    '&:hover': {
      color: theme.palette.text.primary,
      textDecoration: 'none',
    },
  },
  breadcrumbSep: {
    color: theme.palette.text.disabled,
    fontFamily: 'Lato, sans-serif',
    fontSize: 11,
    userSelect: 'none',
  },
  propertyPicker: {
    marginLeft: 'auto',
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(0.5),
    cursor: 'pointer',
    color: theme.palette.text.secondary,
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    padding: theme.spacing(0.25, 1),
    borderRadius: 3,
    border: `1px solid ${theme.palette.divider}`,
    backgroundColor: theme.palette.common.white,
    '&:hover': { backgroundColor: '#f8f9fa' },
    '& .MuiIcon-root': { fontSize: '14px !important' },
  },
  propertyText: {
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: 'inherit',
  },
}));

const PAGE_ICONS = ['dashboard', 'handshake', 'analytics', 'forum', 'settings'];

export function DynamicHeader({
  activePage,
  onNavigate,
  breadcrumbs,
  propertyName = 'Hotel Las Américas',
}: Props) {
  const classes = useStyles();
  const [megaOpen, setMegaOpen] = useState(false);

  const crumbs: BreadcrumbOption[] = breadcrumbs ?? [
    { title: 'Home' },
    { title: 'Travel Distribution Hubs' },
    { title: PAGE_LABELS[activePage].label },
  ];

  const handleMegaSelect = (page: PageId) => {
    onNavigate(page);
    setMegaOpen(false);
  };

  return (
    <div className={classes.root}>
      {/* Top nav bar — 40px, secondary.main */}
      <div className={classes.topBar}>
        <span className={classes.logo}>duetto</span>

        <div className={classes.topNavTabs}>
          {TOP_NAV_TABS.map((tab) => {
            const isTDH = tab === 'Travel Distribution Hub';
            return (
              <div
                key={tab}
                className={clsx(classes.topNavTab, isTDH && classes.topNavTabActive)}
                onClick={isTDH ? () => setMegaOpen((o) => !o) : undefined}
              >
                {tab}
                {isTDH && (
                  <Icon style={{ fontSize: 14, marginLeft: 2, verticalAlign: 'middle' }}>
                    {megaOpen ? 'expand_less' : 'expand_more'}
                  </Icon>
                )}
              </div>
            );
          })}
        </div>

        <div className={classes.utilities}>
          <IconButton className={classes.utilBtn} size="small">
            <Badge badgeContent="99+" color="error" style={{ fontSize: 9 }}>
              <Icon>notifications</Icon>
            </Badge>
          </IconButton>
          <IconButton className={classes.utilBtn} size="small">
            <Icon>help_outline</Icon>
          </IconButton>
          <IconButton className={classes.utilBtn} size="small">
            <Icon>settings</Icon>
          </IconButton>
          <div className={classes.avatar}>M</div>
        </div>
      </div>

      {/* Mega menu (Travel Distribution Hub) */}
      {megaOpen && (
        <>
          <div
            style={{ position: 'fixed', inset: 0, zIndex: 998 }}
            onClick={() => setMegaOpen(false)}
          />
          <div className={classes.megaMenuWrap} style={{ zIndex: 999 }}>
            <Paper elevation={3} className={classes.megaMenu}>
              {PAGE_LABELS.map((item) => (
                <div
                  key={item.id}
                  className={clsx(classes.megaItem, activePage === item.id && classes.megaItemActive)}
                  onClick={() => handleMegaSelect(item.id)}
                >
                  <Icon>{PAGE_ICONS[item.id]}</Icon>
                  {item.label}
                </div>
              ))}
            </Paper>
          </div>
        </>
      )}

      {/* Breadcrumb bar — 32px, background.default */}
      <div className={classes.breadcrumbBar}>
        {crumbs.map((crumb, i) => (
          <React.Fragment key={i}>
            {i > 0 && <span className={classes.breadcrumbSep}>/</span>}
            <Typography
              className={clsx(
                classes.breadcrumbItem,
                i === crumbs.length - 1 && classes.breadcrumbCurrent,
              )}
            >
              {crumb.title}
            </Typography>
          </React.Fragment>
        ))}

        <div className={classes.propertyPicker}>
          <Icon style={{ fontSize: 14 }}>business</Icon>
          <Typography className={classes.propertyText}>{propertyName}</Typography>
          <Icon style={{ fontSize: 14 }}>expand_more</Icon>
        </div>
      </div>
    </div>
  );
}
