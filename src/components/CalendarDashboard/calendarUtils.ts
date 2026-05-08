import { CalendarDay, DayCellMetricRow, MonthData } from './types';

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

type MetricDef = {
  label: string;
  isCompare: boolean;
  gen: (day: number, month: number) => string;
};

export const METRIC_LABEL_MAP: Record<string, MetricDef> = {
  occ_act_1_hotel: { label: 'Occ',      isCompare: false, gen: (d, m) => `${65 + ((d * 7 + m) % 30)}%` },
  occ_act_2_hotel: { label: 'Occ LY',   isCompare: false, gen: (d, m) => `${60 + ((d * 5 + m) % 25)}%` },
  occ_act_3_hotel: { label: 'Occ STLY', isCompare: false, gen: (d, m) => `${58 + ((d * 9 + m) % 28)}%` },
  occ_act_1_op:    { label: 'Occ',      isCompare: true,  gen: (d, m) => `${63 + ((d * 11 + m) % 22)}%` },
  occ_act_2_op:    { label: 'Occ LY',   isCompare: true,  gen: (d, m) => `${60 + ((d * 3 + m) % 20)}%` },
  occ_act_3_op:    { label: 'Occ STLY', isCompare: true,  gen: (d, m) => `${55 + ((d * 13 + m) % 25)}%` },
  adr_act_1_hotel: { label: 'ADR',      isCompare: false, gen: (d, m) => `$${240 + ((d * 7 + m * 3) % 80)}` },
  adr_act_2_hotel: { label: 'ADR LY',   isCompare: false, gen: (d, m) => `$${225 + ((d * 5 + m * 2) % 70)}` },
  adr_act_3_hotel: { label: 'ADR STLY', isCompare: false, gen: (d, m) => `$${220 + ((d * 11 + m) % 65)}` },
  adr_act_4_hotel: { label: 'ADR Fcst', isCompare: false, gen: (d, m) => `$${235 + ((d * 9 + m * 4) % 75)}` },
  adr_act_5_hotel: { label: 'ADR Bgt',  isCompare: false, gen: (d, m) => `$${230 + ((d * 13 + m) % 72)}` },
  adr_act_1_op:    { label: 'ADR',      isCompare: true,  gen: (d, m) => `$${228 + ((d * 7 + m * 2) % 58)}` },
  adr_act_2_op:    { label: 'ADR LY',   isCompare: true,  gen: (d, m) => `$${222 + ((d * 5 + m) % 52)}` },
  adr_act_3_op:    { label: 'ADR',      isCompare: true,  gen: (d, m) => `$${230 + ((d * 9 + m * 3) % 60)}` },
  adr_act_4_op:    { label: 'ADR LY',   isCompare: true,  gen: (d, m) => `$${220 + ((d * 11 + m * 2) % 55)}` },
  adr_act_5_op:    { label: 'ADR STLY', isCompare: true,  gen: (d, m) => `$${215 + ((d * 13 + m) % 50)}` },
  av_rooms_hotel:  { label: 'AvR',      isCompare: false, gen: (d, m) => `${15 + ((d * 7 + m) % 25)} rm` },
  av_guar_op:      { label: 'Guar',     isCompare: true,  gen: (d, m) => `${8 + ((d * 5 + m) % 15)}` },
};

export const DEFAULT_METRIC_IDS = [
  'occ_act_1_hotel',
  'occ_act_1_op',
  'adr_act_1_hotel',
  'av_rooms_hotel',
];

export function buildMonthData(
  year: number,
  month: number,
  options?: { isLocked?: boolean; cellMetricIds?: string[] },
): MonthData {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const today = new Date();

  const metricIds = options?.cellMetricIds ?? DEFAULT_METRIC_IDS;
  const days: CalendarDay[] = [];

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

    const metrics: DayCellMetricRow[] = metricIds
      .filter((id) => id in METRIC_LABEL_MAP)
      .map((id) => {
        const def = METRIC_LABEL_MAP[id];
        return { label: def.label, value: def.gen(d, month), isCompare: def.isCompare };
      });

    days.push({
      date,
      dayNumber: d,
      isInMonth: true,
      isClosed: false,
      isToday,
      metrics,
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
