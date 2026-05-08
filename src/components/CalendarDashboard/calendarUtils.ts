import { CalendarDay, MonthData } from './types';

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

export function buildMonthData(year: number, month: number, options?: {
  isLocked?: boolean;
}): MonthData {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const today = new Date();

  const days: CalendarDay[] = [];

  // Leading empty cells (days before month starts, Sunday = 0)
  for (let i = 0; i < firstDay.getDay(); i++) {
    days.push({
      date: new Date(year, month, 1 - (firstDay.getDay() - i)),
      dayNumber: 0,
      isInMonth: false,
      isClosed: false,
      metrics: [],
    });
  }

  for (let d = 1; d <= lastDay.getDate(); d++) {
    const date = new Date(year, month, d);
    const isToday =
      date.getFullYear() === today.getFullYear() &&
      date.getMonth() === today.getMonth() &&
      date.getDate() === today.getDate();

    days.push({
      date,
      dayNumber: d,
      isInMonth: true,
      isClosed: false,
      isToday,
      metrics: [
        { label: 'Occ', value: String(Math.floor(80 + Math.random() * 15)) },
        { label: 'Occ', value: String(Math.floor(75 + Math.random() * 15)), isCompare: true },
        { label: 'AvR', value: `${Math.floor(15 + Math.random() * 20)} rm` },
        { label: 'Occ', value: `${(70 + Math.random() * 20).toFixed(1)}%` },
      ],
    });
  }

  return {
    year,
    month,
    label: `${MONTH_NAMES[month]} ${year}`,
    isLocked: options?.isLocked,
    days,
  };
}
