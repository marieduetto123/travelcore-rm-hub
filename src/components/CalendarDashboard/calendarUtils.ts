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
  // Occupancy
  occ_act_1_hotel: { label: 'Occ',       isCompare: false, gen: (d, m) => `${65 + ((d * 7 + m) % 30)}%` },
  occ_act_2_hotel: { label: 'Occ LY',    isCompare: false, gen: (d, m) => `${60 + ((d * 5 + m) % 25)}%` },
  occ_act_3_hotel: { label: 'Occ STLY',  isCompare: false, gen: (d, m) => `${58 + ((d * 9 + m) % 28)}%` },
  occ_fcst_hotel:  { label: 'Occ Fcst',  isCompare: false, gen: (d, m) => `${62 + ((d * 11 + m) % 27)}%` },
  occ_act_1_op:    { label: 'Occ',       isCompare: true,  gen: (d, m) => `${63 + ((d * 11 + m) % 22)}%` },
  occ_act_2_op:    { label: 'Occ LY',    isCompare: true,  gen: (d, m) => `${60 + ((d * 3 + m) % 20)}%` },
  occ_act_3_op:    { label: 'Occ STLY',  isCompare: true,  gen: (d, m) => `${55 + ((d * 13 + m) % 25)}%` },
  occ_fcst_op:     { label: 'Occ Fcst',  isCompare: true,  gen: (d, m) => `${57 + ((d * 7 + m) % 23)}%` },
  // ADR
  adr_act_1_hotel: { label: 'ADR',       isCompare: false, gen: (d, m) => `$${240 + ((d * 7 + m * 3) % 80)}` },
  adr_act_2_hotel: { label: 'ADR LY',    isCompare: false, gen: (d, m) => `$${225 + ((d * 5 + m * 2) % 70)}` },
  adr_act_3_hotel: { label: 'ADR STLY',  isCompare: false, gen: (d, m) => `$${220 + ((d * 11 + m) % 65)}` },
  adr_act_4_hotel: { label: 'ADR Fcst',  isCompare: false, gen: (d, m) => `$${235 + ((d * 9 + m * 4) % 75)}` },
  adr_act_5_hotel: { label: 'ADR Bgt',   isCompare: false, gen: (d, m) => `$${230 + ((d * 13 + m) % 72)}` },
  adr_fcst_hotel:  { label: 'ADR Fcst',  isCompare: false, gen: (d, m) => `$${232 + ((d * 6 + m * 3) % 68)}` },
  adr_act_1_op:    { label: 'ADR',       isCompare: true,  gen: (d, m) => `$${228 + ((d * 7 + m * 2) % 58)}` },
  adr_act_2_op:    { label: 'ADR LY',    isCompare: true,  gen: (d, m) => `$${222 + ((d * 5 + m) % 52)}` },
  adr_act_3_op:    { label: 'ADR',       isCompare: true,  gen: (d, m) => `$${230 + ((d * 9 + m * 3) % 60)}` },
  adr_act_4_op:    { label: 'ADR LY',    isCompare: true,  gen: (d, m) => `$${220 + ((d * 11 + m * 2) % 55)}` },
  adr_act_5_op:    { label: 'ADR STLY',  isCompare: true,  gen: (d, m) => `$${215 + ((d * 13 + m) % 50)}` },
  adr_fcst_op:     { label: 'ADR Fcst',  isCompare: true,  gen: (d, m) => `$${218 + ((d * 8 + m * 2) % 54)}` },
  // Revenue
  rev_act_hotel:   { label: 'Rev',       isCompare: false, gen: (d, m) => `$${18 + ((d * 7 + m) % 12)}k` },
  rev_ly_hotel:    { label: 'Rev LY',    isCompare: false, gen: (d, m) => `$${17 + ((d * 5 + m) % 10)}k` },
  rev_stly_hotel:  { label: 'Rev STLY',  isCompare: false, gen: (d, m) => `$${16 + ((d * 9 + m) % 11)}k` },
  rev_fcst_hotel:  { label: 'Rev Fcst',  isCompare: false, gen: (d, m) => `$${19 + ((d * 11 + m) % 9)}k` },
  rev_act_op:      { label: 'Rev',       isCompare: true,  gen: (d, m) => `$${14 + ((d * 7 + m) % 8)}k` },
  rev_ly_op:       { label: 'Rev LY',    isCompare: true,  gen: (d, m) => `$${13 + ((d * 5 + m) % 7)}k` },
  rev_stly_op:     { label: 'Rev STLY',  isCompare: true,  gen: (d, m) => `$${12 + ((d * 9 + m) % 8)}k` },
  rev_fcst_op:     { label: 'Rev Fcst',  isCompare: true,  gen: (d, m) => `$${15 + ((d * 11 + m) % 7)}k` },
  // RN Sold
  rn_act_hotel:    { label: 'RN',        isCompare: false, gen: (d, m) => `${160 + ((d * 7 + m) % 60)}` },
  rn_ly_hotel:     { label: 'RN LY',     isCompare: false, gen: (d, m) => `${155 + ((d * 5 + m) % 55)}` },
  rn_stly_hotel:   { label: 'RN STLY',   isCompare: false, gen: (d, m) => `${150 + ((d * 9 + m) % 50)}` },
  rn_fcst_hotel:   { label: 'RN Fcst',   isCompare: false, gen: (d, m) => `${158 + ((d * 11 + m) % 52)}` },
  rn_act_op:       { label: 'RN',        isCompare: true,  gen: (d, m) => `${80 + ((d * 7 + m) % 40)}` },
  rn_ly_op:        { label: 'RN LY',     isCompare: true,  gen: (d, m) => `${75 + ((d * 5 + m) % 35)}` },
  rn_stly_op:      { label: 'RN STLY',   isCompare: true,  gen: (d, m) => `${72 + ((d * 9 + m) % 38)}` },
  rn_fcst_op:      { label: 'RN Fcst',   isCompare: true,  gen: (d, m) => `${78 + ((d * 11 + m) % 36)}` },
  // RevPAR
  rp_act_hotel:    { label: 'RevPAR',    isCompare: false, gen: (d, m) => `$${155 + ((d * 7 + m * 3) % 60)}` },
  rp_ly_hotel:     { label: 'RvPAR LY',  isCompare: false, gen: (d, m) => `$${148 + ((d * 5 + m * 2) % 55)}` },
  rp_stly_hotel:   { label: 'RvPAR ST',  isCompare: false, gen: (d, m) => `$${142 + ((d * 9 + m) % 50)}` },
  rp_fcst_hotel:   { label: 'RvPAR Fc',  isCompare: false, gen: (d, m) => `$${152 + ((d * 11 + m) % 58)}` },
  rp_act_op:       { label: 'RevPAR',    isCompare: true,  gen: (d, m) => `$${120 + ((d * 7 + m * 2) % 45)}` },
  rp_ly_op:        { label: 'RvPAR LY',  isCompare: true,  gen: (d, m) => `$${115 + ((d * 5 + m) % 42)}` },
  rp_stly_op:      { label: 'RvPAR ST',  isCompare: true,  gen: (d, m) => `$${110 + ((d * 9 + m) % 40)}` },
  rp_fcst_op:      { label: 'RvPAR Fc',  isCompare: true,  gen: (d, m) => `$${118 + ((d * 11 + m) % 44)}` },
  // Other Metrics
  oth_pickup_hotel:   { label: 'Pickup',   isCompare: false, gen: (d, m) => `${2 + ((d * 3 + m) % 8)}` },
  oth_los_hotel:      { label: 'Avg LOS',  isCompare: false, gen: (d, m) => `${2 + ((d * 5 + m) % 5)}.${(d * 3) % 9}` },
  oth_lead_hotel:     { label: 'Lead',     isCompare: false, gen: (d, m) => `${28 + ((d * 7 + m) % 30)}d` },
  oth_adults_hotel:   { label: 'Adults',   isCompare: false, gen: (d, m) => `${1 + ((d * 3 + m) % 2)}.${(d * 7) % 9}` },
  oth_children_hotel: { label: 'Children', isCompare: false, gen: (d, m) => `0.${((d * 5 + m) % 9)}` },
  oth_guests_hotel:   { label: 'Guests',   isCompare: false, gen: (d, m) => `${380 + ((d * 11 + m) % 80)}` },
  oth_pickup_op:      { label: 'Pickup',   isCompare: true,  gen: (d, m) => `${1 + ((d * 3 + m) % 6)}` },
  oth_los_op:         { label: 'Avg LOS',  isCompare: true,  gen: (d, m) => `${3 + ((d * 5 + m) % 4)}.${(d * 3) % 9}` },
  oth_lead_op:        { label: 'Lead',     isCompare: true,  gen: (d, m) => `${35 + ((d * 7 + m) % 25)}d` },
  oth_adults_op:      { label: 'Adults',   isCompare: true,  gen: (d, m) => `${1 + ((d * 3 + m) % 3)}.${(d * 5) % 9}` },
  oth_children_op:    { label: 'Children', isCompare: true,  gen: (d, m) => `0.${((d * 7 + m) % 9)}` },
  oth_guests_op:      { label: 'Guests',   isCompare: true,  gen: (d, m) => `${120 + ((d * 9 + m) % 50)}` },
  // Availability
  av_rooms_hotel:     { label: 'AvR',      isCompare: false, gen: (d, m) => `${15 + ((d * 7 + m) % 25)} rm` },
  av_guar_op:         { label: 'Guar',     isCompare: true,  gen: (d, m) => `${8 + ((d * 5 + m) % 15)}` },
  // Business Mix
  bm_to_hotel:        { label: 'TO Mix',   isCompare: false, gen: (d, m) => `${35 + ((d * 7 + m) % 20)}%` },
  bm_direct_hotel:    { label: 'Dir Mix',  isCompare: false, gen: (d, m) => `${25 + ((d * 5 + m) % 15)}%` },
  bm_ota_hotel:       { label: 'OTA Mix',  isCompare: false, gen: (d, m) => `${20 + ((d * 9 + m) % 12)}%` },
  bm_to_op:           { label: 'TO Mix',   isCompare: true,  gen: (d, m) => `${40 + ((d * 7 + m) % 18)}%` },
  bm_direct_op:       { label: 'Dir Mix',  isCompare: true,  gen: (d, m) => `${22 + ((d * 5 + m) % 14)}%` },
  bm_ota_op:          { label: 'OTA Mix',  isCompare: true,  gen: (d, m) => `${18 + ((d * 9 + m) % 10)}%` },
  // Selling Rates
  sr_contract_hotel:  { label: 'Contract', isCompare: false, gen: (d, m) => `$${185 + ((d * 7 + m * 2) % 60)}` },
  sr_promo_hotel:     { label: 'Promo',    isCompare: false, gen: (d, m) => `${5 + ((d * 5 + m) % 10)}%` },
  sr_base_hotel:      { label: 'Base',     isCompare: false, gen: (d, m) => `$${200 + ((d * 9 + m * 3) % 70)}` },
  sr_contract_op:     { label: 'Contract', isCompare: true,  gen: (d, m) => `$${175 + ((d * 7 + m * 2) % 55)}` },
  sr_promo_op:        { label: 'Promo',    isCompare: true,  gen: (d, m) => `${4 + ((d * 5 + m) % 8)}%` },
  sr_base_op:         { label: 'Base',     isCompare: true,  gen: (d, m) => `$${190 + ((d * 9 + m * 3) % 65)}` },
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
