import React from 'react';
import ReactDOM from 'react-dom/client';
import { createMuiTheme, ThemeProvider } from '@material-ui/core/styles';
import CssBaseline from '@material-ui/core/CssBaseline';
import { App } from './App';

const duettoTheme = createMuiTheme({
  palette: {
    primary: {
      main: '#006461',
      dark: '#053c3c',
    },
    secondary: {
      main: '#0E2124',
    },
    text: {
      primary: '#1c1c1c',
      secondary: '#4f5b60',
      disabled: '#aeb4ba',
    },
    background: {
      default: '#fafafa',
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
      <App />
    </ThemeProvider>
  </React.StrictMode>
);
