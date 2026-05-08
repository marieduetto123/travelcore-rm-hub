import React from 'react';
import ReactDOM from 'react-dom/client';
import { createMuiTheme, ThemeProvider } from '@material-ui/core/styles';
import CssBaseline from '@material-ui/core/CssBaseline';
import { CalendarDashboardPage } from './components/CalendarDashboard';

const duettoTheme = createMuiTheme({
  palette: {
    primary: {
      main: '#006461',
      dark: '#004d4a',
    },
    secondary: {
      main: '#0E2124',
    },
    text: {
      primary: '#1c1c1c',
      secondary: '#4f5b60',
      disabled: '#8a9096',
    },
    background: {
      default: '#f1f5f9',
      paper: '#ffffff',
    },
    error: {
      main: '#d32f2f',
    },
    divider: '#dde1e2',
  },
  typography: {
    fontFamily: 'Lato, sans-serif',
  },
});

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ThemeProvider theme={duettoTheme}>
      <CssBaseline />
      <CalendarDashboardPage />
    </ThemeProvider>
  </React.StrictMode>
);
