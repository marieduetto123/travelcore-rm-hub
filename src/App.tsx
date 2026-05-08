import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import clsx from 'clsx';
import { DynamicHeader, type PageId } from './components/AppShell/DynamicHeader';
import { CalendarDashboardPage } from './components/CalendarDashboard';
import { ContactsPage } from './pages/ContactsPage';
import { AuditPage } from './pages/AuditPage';
import { CommunicationsPage } from './pages/CommunicationsPage';
import { ConfigurationPage } from './pages/ConfigurationPage';

const NAV_ITEMS = [
  { icon: 'dashboard', label: 'Dashboard' },
  { icon: 'handshake', label: 'Contacts & Contracts' },
  { icon: 'analytics', label: 'Audit' },
  { icon: 'forum', label: 'Communications & Notes' },
  { icon: 'settings', label: 'Configuration' },
];

const SIDEBAR_BG = '#19393e';
const SIDEBAR_ACTIVE_BG = '#2a5258';
const SIDEBAR_ACCENT = '#c4ff45';
const SIDEBAR_TEXT_INACTIVE = '#94a3b8';
const SIDEBAR_TEXT_COLLAPSE = '#64748b';

const useStyles = makeStyles((theme) => ({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    overflow: 'hidden',
    backgroundColor: theme.palette.background.default,
  },
  body: {
    display: 'flex',
    flex: 1,
    overflow: 'hidden',
  },
  sidebar: {
    width: 240,
    backgroundColor: SIDEBAR_BG,
    flexShrink: 0,
    display: 'flex',
    flexDirection: 'column',
    paddingTop: theme.spacing(2),
    transition: 'width 0.2s ease',
    overflow: 'hidden',
  },
  sidebarCollapsed: {
    width: 48,
  },
  navItem: {
    position: 'relative',
    height: 48,
    display: 'flex',
    alignItems: 'center',
    paddingLeft: 20,
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 13,
    fontWeight: 400,
    color: SIDEBAR_TEXT_INACTIVE,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    transition: 'background-color 0.15s, color 0.15s',
    '&:hover': {
      backgroundColor: SIDEBAR_ACTIVE_BG,
      color: theme.palette.common.white,
    },
  },
  navItemActive: {
    backgroundColor: SIDEBAR_ACTIVE_BG,
    color: theme.palette.common.white,
    fontWeight: 600,
    '&::before': {
      content: '""',
      position: 'absolute',
      left: 0,
      top: '50%',
      transform: 'translateY(-50%)',
      width: 3,
      height: 30,
      backgroundColor: SIDEBAR_ACCENT,
      borderRadius: '0 2px 2px 0',
    },
  },
  collapseBtn: {
    marginTop: 'auto',
    height: 40,
    display: 'flex',
    alignItems: 'center',
    paddingLeft: 20,
    cursor: 'pointer',
    fontFamily: 'Lato, sans-serif',
    fontSize: 12,
    color: SIDEBAR_TEXT_COLLAPSE,
    whiteSpace: 'nowrap',
    '&:hover': { color: theme.palette.common.white },
  },
  main: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'hidden',
    backgroundColor: theme.palette.background.default,
  },
}));

export function App() {
  const classes = useStyles();
  const [activePage, setActivePage] = useState<PageId>(0);
  const [collapsed, setCollapsed] = useState(false);

  const renderPage = () => {
    switch (activePage) {
      case 1: return <ContactsPage />;
      case 2: return <AuditPage />;
      case 3: return <CommunicationsPage />;
      case 4: return <ConfigurationPage />;
      default: return <CalendarDashboardPage />;
    }
  };

  return (
    <div className={classes.root}>
      <DynamicHeader activePage={activePage} onNavigate={setActivePage} />

      <div className={classes.body}>
        {/* Left sidebar */}
        <div className={clsx(classes.sidebar, collapsed && classes.sidebarCollapsed)}>
          {NAV_ITEMS.map((item, i) => (
            <div
              key={item.label}
              className={clsx(classes.navItem, activePage === i && classes.navItemActive)}
              onClick={() => setActivePage(i as PageId)}
            >
              {!collapsed && item.label}
            </div>
          ))}

          <div
            className={classes.collapseBtn}
            onClick={() => setCollapsed((c) => !c)}
          >
            {collapsed ? '›' : '‹ Collapse'}
          </div>
        </div>

        {/* Page content */}
        <div className={classes.main}>{renderPage()}</div>
      </div>
    </div>
  );
}
