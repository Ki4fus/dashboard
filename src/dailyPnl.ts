import type {Point} from './data';

function summarizeDay(rows: Point[]) {
  const sum = (field: 'total' | 'fxPnl', issue: 'issue' | 'fxIssue') => {
    if (!rows.length || rows.some(point => point[issue] || typeof point[field] !== 'number' || !Number.isFinite(point[field]))) return null;
    return rows.reduce((total, point) => total + point[field]!, 0);
  };
  const our = sum('total', 'issue');
  const fx = sum('fxPnl', 'fxIssue');
  return {our, fx, combined: (our ?? 0) + (fx ?? 0)};
}

export function dailyPnl(points: Point[], today: string) {
  return summarizeDay(points.filter(point => point.date === today));
}

export function periodPnl(points: Point[], from: string, to: string) {
  const days = new Map<string, Point[]>();
  for (const point of points) {
    if (!point.date || (from && point.date < from) || (to && point.date > to)) continue;
    const rows = days.get(point.date) ?? [];
    rows.push(point);
    days.set(point.date, rows);
  }
  // Sum the same combined daily result displayed by the Today card for each day.
  const combined = [...days.values()].reduce((sum, rows) => sum + summarizeDay(rows).combined, 0);
  const fxValues = [...days.values()].map(rows => summarizeDay(rows).fx).filter((value): value is number => value !== null);
  const fx = fxValues.length ? fxValues.reduce((sum, value) => sum + value, 0) : null;
  return {fx, combined, days: days.size};
}
