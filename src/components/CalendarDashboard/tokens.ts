// Duetto design tokens not yet in the standard MUI v4 theme palette.
// These should be migrated to the shared Duetto theme config once confirmed.
export const calendarTokens = {
  border: '#dde1e2',
  cellBackground: '#f5f5f5',
  cellBackgroundClosed: '#f5f6f7',
  compareRowBackground: '#fee6c9', // "personal" token — compare/benchmark metric row bg
  benchmarkColor: '#8c7843',       // Sycamore — compare delta text colour (week view etc.)
  checkboxBorder: '#4f5b60',       // text/secondary — matches Figma checkbox idle state
  headerShadow: '0px 2px 3px rgba(0,0,0,0.07)',
  dropdownBackground: '#fafafa',
  primaryHover: 'rgba(0, 100, 97, 0.08)',
} as const;
