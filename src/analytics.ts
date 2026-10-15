export type Operation = {date: string; row: number; currency: string; rate: number | null; fixedUsdAmount: number | null; hasAmount: boolean};
export function analytics(operations: Operation[], today: string, from: string, to: string) {
  const inPeriod = operations.filter(p => (!from || p.date >= from) && (!to || p.date <= to));
  const average = (rows: Operation[]) => {
    const rates = rows.map(p => p.rate).filter((v): v is number => v !== null && Number.isFinite(v));
    return {value: rates.length ? rates.reduce((a,b)=>a+b,0)/rates.length : null, count: rates.length};
  };
  const current = average(operations.filter(p=>p.date===today));
  const previousDate = operations.filter(p=>p.date<today && p.rate!==null).map(p=>p.date).sort().at(-1);
  const previous = average(operations.filter(p=>p.date===previousDate));
  const change = current.value !== null && previous.value !== null && previous.value !== 0 ? (current.value-previous.value)/Math.abs(previous.value)*100 : null;
  const currencies = new Map<string, number>();
  for (const p of inPeriod.filter(p=>p.fixedUsdAmount !== null)) {
    const name = p.currency || 'Не указана';
    currencies.set(name, (currencies.get(name) ?? 0)+p.fixedUsdAmount!);
  }
  return {current, period: average(inPeriod), previousDate, change, currencyIssue: [...currencies.values()].some(v=>v<0), currencies: [...currencies].map(([name,value])=>({name,value})).filter(p=>p.value!==0).sort((a,b)=>b.value-a.value)};
}
