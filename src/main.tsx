import React from 'react';
import {createRoot} from 'react-dom/client';
import {Button, ThemeProvider, configure, DropdownMenu} from '@gravity-ui/uikit';
import {Chart} from '@gravity-ui/charts';
import '@fontsource-variable/inter';
import '@gravity-ui/uikit/styles/styles.css';
import './styles.css';
import {useCardTitles} from './useCardTitles';
import {AppIcon} from './Icons';
import {DateFilter} from './DateFilter';
import {DashboardGrid, type DashboardCard} from './DashboardGrid';
import {dailyPnl,periodPnl} from './dailyPnl';
import {analytics} from './analytics';
import type {Data} from './data';
configure({lang:'en'});
const OneDriveSource = React.lazy(() => import('./OneDriveSource').then(module => ({default: module.OneDriveSource})));
const desktop = Boolean(window.dashboard);
const todayDate = () => new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
declare global {interface Window {dashboard?: {restore(): Promise<Data | null>; open(): Promise<Data | null>; refresh(): Promise<Data>}}}
const number = (n: number | null | undefined) => n == null ? '—' : new Intl.NumberFormat('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(n);
const counted = (value:number,unit:string) => `${value} ${unit}${value===1?'':'s'}`;
const changeColor = (value: number | null) => value === null || value === 0 ? 'change-neutral' : value < 0 ? 'change-negative' : 'change-positive';
type Section = 'overview' | 'total' | 'fixed' | 'rates';
const sections: {id:Section;label:string}[] = [{id:'overview',label:'Overview'},{id:'total',label:'Our P&L'},{id:'fixed',label:'Fixed Amount'},{id:'rates',label:'Rates'}];
const currencyColors = ['#49a2ed','#ff3765','#78d85e','#ffc43b','#b17bf2','#37c6b7','#f48b56'];
const day = (d: string) => d.split('-').reverse().join('/');
function App() {
  const {titles,rename}=useCardTitles();
  const [signedIn,setSignedIn] = React.useState(desktop);
  const [toolsElement,setToolsElement] = React.useState<HTMLDivElement|null>(null);
  const [section,setSection] = React.useState<Section>('overview');
  const [theme, setTheme] = React.useState<'light' | 'dark'>(()=>{try{return localStorage.getItem('netting-theme')==='dark'?'dark':'light';}catch{return 'light';}});
  const [data, setData] = React.useState<Data | null>(null);
  const [error, setError] = React.useState('');
  const [busy, setBusy] = React.useState(true);
  const [filterReset,setFilterReset] = React.useState(0);
  const [from, setFrom] = React.useState(() => todayDate().slice(0, 7) + '-01');
  const [to, setTo] = React.useState(() => new Intl.DateTimeFormat('en-CA', {timeZone: 'Asia/Dubai', year: 'numeric', month: '2-digit', day: '2-digit'}).format(new Date()));
  React.useEffect(()=>{try{localStorage.setItem('netting-theme',theme);}catch{/* Theme remains available without storage. */}},[theme]);
  const busyRef = React.useRef(false);
  const load = React.useCallback(async (refresh: boolean | 'restore' = false) => {
    if (busyRef.current) return;
    if (!window.dashboard) {setError('Open the desktop app to load a local Excel file.'); return;}
    busyRef.current = true; setBusy(true);
    try {const result = await window.dashboard[refresh === 'restore' ? 'restore' : refresh ? 'refresh' : 'open'](); if (result) {setData(result); setError('');}}
    catch (e) {setError(e instanceof Error ? e.message : 'Could not read the Excel workbook.');}
    finally {busyRef.current = false; setBusy(false);}
  }, []);
  React.useEffect(() => {if (desktop) void load('restore'); else setBusy(false);}, [load]);
  React.useEffect(() => {if (!desktop || !data) return; const timer = setInterval(() => void load(true), 60000); return () => clearInterval(timer);}, [Boolean(data), load]);
  const visible = (data?.points ?? []).filter(p => p.date && (!from || p.date >= from) && (!to || p.date <= to)).sort((a,b) => a.date!.localeCompare(b.date!));
  const valid = visible.filter(p => p.total !== null && !p.issue);
  const fixedValid = visible.filter(p => p.fixedUsdAmount !== null && !p.fixedIssue);
  const fixedSum = fixedValid.reduce((s,p) => s + p.fixedUsdAmount!, 0);
  const sum = valid.reduce((s,p) => s + p.total!, 0);
  const latest = visible.at(-1), previous = visible.at(-2);
  const change = latest?.total != null && previous?.total != null && previous.total !== 0 ? (latest.total - previous.total) / Math.abs(previous.total) * 100 : null;
  const issues = (data?.points ?? []).filter(p => p.issue);
  const today = todayDate();
  const rates = analytics(data?.operations ?? [], today, from, to);
  const currencyTotal = rates.currencies.reduce((sum,p)=>sum+p.value,0);
  const fixedToday = (data?.points ?? []).filter(p => p.date === today && p.fixedUsdAmount !== null && !p.fixedIssue);
  const fixedTodaySum = fixedToday.length ? fixedToday.reduce((sum,p)=>sum+p.fixedUsdAmount!,0) : null;
  const pnlToday = dailyPnl(data?.points ?? [], today);
  const pnlPeriod = periodPnl(data?.points ?? [],from,to);
  const stats = [number(valid.length ? sum : null), number(pnlToday.our), number(valid.length ? sum / valid.length : null), change === null ? '—' : number(change) + '%', number(fixedValid.length ? fixedSum : null), number(fixedTodaySum)];
  const cards: DashboardCard[] = [
    ...["Our P&L · Period (USD)","Our P&L · Today (USD)","Our P&L · Daily Average (USD)","Our P&L · Daily Change (%)","Fixed Amt · Period (USD)","Fixed Amt · Today (USD)"].map((label,i) => ({id:['total-period','total-latest','total-average','total-change','fixed-period','fixed-today'][i],title:label,kind:'metric' as const,content:<><div>{label}</div><strong className={i===3 ? changeColor(change) : undefined}>{stats[i]}</strong><small>{i===1 ? day(today) : i===5 ? day(today) : i===4 ? `Period · ${counted(fixedValid.length,'day')} · column U` : `${counted(valid.length,'day')} with numeric totals`}</small></>})),
    {id:'total-fx-period',title:'FX P&L · Period (USD)',kind:'metric',content:<><div>FX P&L · Period (USD)</div><strong>{number(pnlPeriod.fx)}</strong><small>Selected period · column T</small></>},
    {id:'total-combined-period',title:'P&L+FX· Period (USD)',kind:'metric',content:<><div>P&L+FX· Period (USD)</div><strong>{number(pnlPeriod.combined)}</strong><small>Selected period · Y + T</small></>},
    {id:'total-fx-today',title:'FX P&L · Today (USD)',kind:'metric',content:<><div>FX P&L · Today (USD)</div><strong>{number(pnlToday.fx)}</strong><small>{day(today)} · column T</small></>},
    {id:'total-combined-today',title:'P&L+FX· Today (USD)',kind:'metric',content:<><div>P&L+FX· Today (USD)</div><strong>{number(pnlToday.combined)}</strong><small>{day(today)} · Y + T</small></>},
    {id:'rate-today',title:"Avg. Rate · Today",kind:'metric',content:<><div>Avg. Rate · Today</div><strong>{number(rates.current.value)}</strong><small>{day(today)} · {rates.current.count} values in E</small></>},
    {id:'rate-period',title:"Avg. Rate · Period",kind:'metric',content:<><div>Avg. Rate · Period</div><strong>{number(rates.period.value)}</strong><small>Simple average of E · {rates.period.count} values</small></>},
    {id:'rate-change',title:"Rate · Daily Change (%)",kind:'metric',content:<><div>Rate · Daily Change (%)</div><strong className={changeColor(rates.change)}>{rates.change === null ? '—' : number(rates.change) + '%'}</strong><small>{rates.previousDate ? 'Compared with ' + day(rates.previousDate) : 'No earlier day with a rate'}</small></>},
    {id:'currencies',title:"Currency breakdown",kind:'panel',content:<><h2>Currency breakdown</h2><p>Column L · share of Fixed Amount (U) for the selected period</p>{rates.currencyIssue ? <div className="empty"><h3>Negative currency totals</h3><p>A pie chart requires non-negative currency totals.</p></div> : rates.currencies.length ? <div className="currency-layout"><div className="currency-chart"><Chart data={{legend:{enabled:false},series:{data:[{type:'pie',tooltip:{valueFormat:{type:'custom',formatter:({value})=>number(typeof value === 'number' ? value : null) + ' USD'}},data:rates.currencies.map((p,i)=>({...p,color:currencyColors[i%currencyColors.length]})),dataLabels:{enabled:false}}]}}}/></div><div className="currency-legend">{rates.currencies.map((p,i)=><div className="currency-entry" key={p.name}><div className="currency-entry-heading"><span className="currency-name"><span className="currency-dot" style={{background:currencyColors[i%currencyColors.length]}}/>{p.name}</span><span>{number(p.value/currencyTotal*100)}%</span></div><div className="numeric-value currency-amount">{number(p.value)} USD</div></div>)}</div></div> : <div className="empty"><h3>No transactions in the selected period</h3></div>}</>},
    {id:'total-chart',title:"Trend of our P&L · Today (USD)",kind:'panel',content:<><h2>Trend of our P&L · Today (USD)</h2><p>Block date: A · daily Our P&L: Y · USD</p>{valid.length ? <div style={{height:320}}><Chart data={{tooltip:{valueFormat:{type:'custom',formatter:({value})=>number(typeof value === 'number' ? value : null)}},yAxis:[{labels:{numberFormat:{precision:2,showRankDelimiter:true,lang:'en',unit:null},style:{fontColor:'#3585f5'}}}],xAxis:{type:'category',categories:visible.map(p=>day(p.date!))},series:{data:[{type:'line',color:'#3585f5',name:'Our P&L (USD)',data:visible.map((p,i)=>({x:i,y:p.issue ? null : p.total}))}]}}}/></div> : <div className="empty"><span className="chart-icon">▥</span><h3>No numeric totals in the selected period</h3><p>Load an Excel workbook or change the date range.</p></div>}</>},
    {id:'daily-table',title:"Daily data",kind:'panel',content:<><h2>Daily data</h2><div className="table-scroll"><table><thead><tr><th>Date</th><th>Our P&L (USD)</th><th>Fixed Amount (USD)</th><th>Excel row</th><th>Status</th></tr></thead><tbody>{visible.length ? [...visible].reverse().map(p=><tr key={p.row}><td>{day(p.date!)}</td><td className="numeric-value">{number(p.total)}</td><td className="numeric-value" title={p.fixedIssue ?? undefined}>{number(p.fixedUsdAmount)}</td><td>{p.row}</td><td>{p.issue ?? 'Saved Excel total'}</td></tr>) : <tr><td colSpan={5} className="empty-row">No data loaded or no data in this period</td></tr>}</tbody></table></div></>},
    {id:'date-issues',title:"Check Excel dates",kind:'panel',content:<><h2>Check Excel dates</h2><p>These blocks have missing or conflicting dates. Their totals are excluded from the trend until corrected.</p>{issues.filter(p=>!p.date).map(p=><p key={p.row}>Row {p.row}: {p.issue}</p>)}</>},
  ];
  const sectionCards = cards.filter(card => section === 'overview' || ['daily-table','date-issues'].includes(card.id) || (section === 'total' ? card.id.startsWith('total-') : section === 'rates' ? card.id.startsWith('rate-') : card.id.startsWith('fixed-') || card.id === 'currencies'));
  return <ThemeProvider theme={theme}><div className={signedIn?'shell':'auth-shell'}>

    <main className={signedIn?'dashboard-main':'auth-main'}>{!signedIn && <div className="auth-brand"><img src="./netting-logo.svg" width="36" height="36" alt=""/><span>Netting Dashboard</span></div>}{signedIn && <header><div><div className="eyebrow">EXECUTIVE DASHBOARD</div><h1 className="dashboard-title"><img src="./netting-logo.svg" width="32" height="32" alt=""/><span>Netting Dashboard</span></h1><p>{sections.find(item=>item.id===section)!.label} · RUB Deals Rates</p></div><div className="header-right"><div className="powered-by">Powered by LSD</div><div className="actions header-actions">{desktop && <><Button size="l" view="action" loading={busy} onClick={() => void load()}>Open Excel</Button><Button size="l" disabled={!data || busy} onClick={() => void load(true)}>Refresh</Button></>}<Button view="flat" size="l" className="icon-button" title={theme==='light'?'Dark theme':'Light theme'} aria-label={theme==='light'?'Dark theme':'Light theme'} onClick={()=>setTheme(theme==='light'?'dark':'light')}><AppIcon name={theme==='light'?'moon':'sun'}/></Button><div ref={setToolsElement} className="header-menu-slot">{desktop && <DropdownMenu size="l" icon={<AppIcon name="menu"/>} defaultSwitcherClassName="icon-button" defaultSwitcherProps={{view:'flat','aria-label':'Open menu',title:'Menu'}} popupProps={{placement:'bottom-end'}} items={sections.map(item=>({text:item.label,selected:section===item.id,action:()=>setSection(item.id)}))}/>}</div></div></div></header>}
    {!desktop && <React.Suspense fallback={<p role="status">Preparing sign-in…</p>}><OneDriveSource onData={setData} onBusy={setBusy} onSessionChange={setSignedIn} toolsElement={toolsElement} navigation={sections} section={section} onSection={id=>setSection(id as Section)}/></React.Suspense>}
    {error && <div role="alert" className="notice">{error} {data && 'Showing data from the last successful load.'}</div>}
    {(desktop || data) && <>

    <div className="filters"><DateFilter key={`from-${filterReset}`} label="From" value={from} onChange={setFrom}/><DateFilter key={`to-${filterReset}`} label="To" value={to} onChange={setTo}/><Button onClick={()=>{setFrom('');setTo('');setFilterReset(value=>value+1);}}>All dates</Button></div>
    {from && to && from > to && <p role="alert">The start date is after the end date.</p>}
    <DashboardGrid key={section} cards={sectionCards} loading={busy} layoutId={section} titles={titles} onRename={rename}/>
    </>}
    </main></div></ThemeProvider>;
}
createRoot(document.getElementById('root')!).render(<App/>);
