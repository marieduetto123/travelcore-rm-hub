export type DayCellMetricRow = {
  label: string;
  value: string;
  isCompare?: boolean;
};

export type CalendarDay = {
  date: Date;
  dayNumber: number;
  isInMonth: boolean;
  isClosed: boolean;
  isToday?: boolean;
  isSelected?: boolean;
  metrics: DayCellMetricRow[];
};

export type MonthData = {
  year: number;
  month: number;
  label: string;
  isLocked?: boolean;
  days: CalendarDay[];
};

export type DayDetailItem = {
  label: string;
  percentage?: number;
  seats?: number;
  value?: string;
  isNegative?: boolean;
};

export type DayDetailGroup = {
  title: string;
  isExpanded?: boolean;
  items: DayDetailItem[];
};

export type RoomTypeRow = {
  roomType: string;
  sold: number;
  available: number;
  occupancy: number;
  adr: number;
  revPar: number;
  revenue: number;
};
