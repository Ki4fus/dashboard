import type {Point} from './data';

export function dailyPnl(points: Point[], today: string) {
  const rows = points.filter(point => point.date === today);
  const sum = (field: 'total' | 'fxPnl', issue: 'issue' | 'fxIssue') => {
    if (!rows.length || rows.some(point => point[issue] || typeof point[field] !== 'number' || !Number.isFinite(point[field]))) return null;
    return rows.reduce((total, point) => total + point[field]!, 0);
  };
  const our = sum('total', 'issue');
  const fx = sum('fxPnl', 'fxIssue');
  return {our, fx, combined: (our ?? 0) + (fx ?? 0)};
}
