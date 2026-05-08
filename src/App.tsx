import React, { useState } from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Icon from '@material-ui/core/Icon';
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
    width: 220,
    backgroundColor: theme.palette.secondary.main,
    flexShrink: 0,
    display: 'flex',
    flexDirection: 'column',
    paddingTop: theme.spacing(1),
    transition: 'width 0.2s ease',
    overflow: 'hidden',
  },
  sidebarCollapsed: {
    width: 52,
  },
  navItem: {
    display: 'flex',
    alignItems: 'center',
    gap: theme.spacing(1.5),
    padding: theme.spacing(1.5, 2),
    cursor: 'pointer',
    color: theme.palette.common.white,
    fontFamily: 'Lato, sans-serif',
    fontSize: 14,
    opacity: 0.65,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    transition: 'opacity 0.15s, background-color 0.15s',
    '&:hover': {
      opacity: 1,
      backgroundColor: 'rgba(255,255,255,0.08)',
    },
    '& .MuiIcon-root': {
      fontSize: '20px !important',
      flexShrink: 0,
    },
  },
  navItemActive: {
    opacity: 1,
    backgroundColor: 'rgba(255,255,255,0.12)',
    fontWeight: 700,
  },
  collapseBtn: {
    marginTop: 'auto',
    display: 'flex',
    alignItems: 'center',
    justifyContent: collapsed => (collapsed ? 'center' : 'flex-start'),
    padding: theme.spacing(1.5, 2),
    cursor: 'pointer',
    color: 'rgba(255,255,255,0.4)',
    gap: theme.spacing(1),
    '&:hover': { color: theme.palette.common.white },
    '& .MuiIcon-root': { fontSize: '18px !important' },
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
              <Icon>{item.icon}</Icon>
              {!collapsed && item.label}
            </div>
          ))}

          <div
            className={classes.collapseBtn}
            onClick={() => setCollapsed((c) => !c)}
          >
            <Icon>{collapsed ? 'chevron_right' : 'chevron_left'}</Icon>
            {!collapsed && (
              <span style={{ fontFamily: 'Lato, sans-serif', fontSize: 12 }}>Collapse</span>
            )}
          </div>
        </div>

        {/* Page content */}
        <div className={classes.main}>{renderPage()}</div>
      </div>
    </div>
  );
}
