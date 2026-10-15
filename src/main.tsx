import React from 'react';
import {createRoot} from 'react-dom/client';
import {Button, ThemeProvider} from '@gravity-ui/uikit';
import {Chart} from '@gravity-ui/charts';
import '@fontsource-variable/inter';
import '@gravity-ui/uikit/styles/styles.css';
import './styles.css';
import {useCardTitles} from './useCardTitles';
import {AppIcon} from './Icons';
import {DateFilter} from './DateFilter';
import {DashboardGrid, type DashboardCard} from './DashboardGrid';
import {analytics} from './analytics';
import type {Data} from './data';
const OneDriveSource = React.lazy(() => import('./OneDriveSource').then(module => ({default: module.OneDriveSource})));
const desktop = Boolean(window.dashboard);
const todayDate = () => new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
declare global {interface Window {dashboard?: {restore(): Promise<Data | null>; open(): Promise<Data | null>; refresh(): Promise<Data>}}}
const number = (n: number | null | undefined) => n == null ? '—' : new Intl.NumberFormat('ru-RU', {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(n);
const changeColor = (value: number | null) => value === null || value === 0 ? 'change-neutral' : value < 0 ? 'change-negative' : 'change-positive';
type Section = 'overview' | 'total' | 'fixed' | 'rates';
const sections: {id:Section;label:string}[] = [{id:'overview',label:'Обзор показателей'},{id:'total',label:'Total'},{id:'fixed',label:'Fixed USD Amount'},{id:'rates',label:'Курс'}];
const currencyColors = ['#49a2ed','#ff3765','#78d85e','#ffc43b','#b17bf2','#37c6b7','#f48b56'];
const day = (d: string) => d.split('-').reverse().join('.');
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
    if (!window.dashboard) {setError('Откройте приложение через Electron для загрузки Excel.'); return;}
    busyRef.current = true; setBusy(true);
    try {const result = await window.dashboard[refresh === 'restore' ? 'restore' : refresh ? 'refresh' : 'open'](); if (result) {setData(result); setError('');}}
    catch (e) {setError(e instanceof Error ? e.message : 'Не удалось прочитать Excel.');}
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
  const stats = [number(valid.length ? sum : null), number(latest?.total), number(valid.length ? sum / valid.length : null), change === null ? '—' : number(change) + '%', number(fixedValid.length ? fixedSum : null), number(fixedTodaySum)];
  const cards: DashboardCard[] = [
    ...["Our P&L · Period (USD)","Our P&L · Latest Day (USD)","Our P&L · Daily Average (USD)","Our P&L · Daily Change (%)","Fixed USD · Period","Fixed USD · Today"].map((label,i) => ({id:['total-period','total-latest','total-average','total-change','fixed-period','fixed-today'][i],title:label,kind:'metric' as const,content:<><div>{label}</div><strong className={i===3 ? changeColor(change) : undefined}>{stats[i]}</strong><small>{i===1 && latest?.date ? day(latest.date) : i===5 ? day(today) : i===4 ? `За период · ${fixedValid.length} дней · столбец U` : `${valid.length} дней с числовым итогом`}</small></>})),
    {id:'rate-today',title:"Avg. Rate · Today",kind:'metric',content:<><div>Avg. Rate · Today</div><strong>{number(rates.current.value)}</strong><small>{day(today)} · {rates.current.count} значений E</small></>},
    {id:'rate-period',title:"Avg. Rate · Period",kind:'metric',content:<><div>Avg. Rate · Period</div><strong>{number(rates.period.value)}</strong><small>Простое среднее E · {rates.period.count} значений</small></>},
    {id:'rate-change',title:"Rate · Daily Change (%)",kind:'metric',content:<><div>Rate · Daily Change (%)</div><strong className={changeColor(rates.change)}>{rates.change === null ? '—' : number(rates.change) + '%'}</strong><small>{rates.previousDate ? 'Сравнение с ' + day(rates.previousDate) : 'Нет предыдущего дня с курсом'}</small></>},
    {id:'currencies',title:"Разбивка по валютам",kind:'panel',content:<><h2>Разбивка по валютам</h2><p>Столбец L · доля по сумме FIXED USD AMOUNT (U) за выбранный период</p>{rates.currencyIssue ? <div className="empty"><h3>Есть отрицательные суммы по валютам</h3><p>Круговая диаграмма требует неотрицательных сумм.</p></div> : rates.currencies.length ? <div className="currency-layout"><div className="currency-chart"><Chart data={{legend:{enabled:false},series:{data:[{type:'pie',tooltip:{valueFormat:{type:'custom',formatter:({value})=>number(typeof value === 'number' ? value : null) + ' USD'}},data:rates.currencies.map((p,i)=>({...p,color:currencyColors[i%currencyColors.length]})),dataLabels:{enabled:false}}]}}}/></div><div className="currency-legend">{rates.currencies.map((p,i)=><div className="currency-entry" key={p.name}><div className="currency-entry-heading"><span className="currency-name"><span className="currency-dot" style={{background:currencyColors[i%currencyColors.length]}}/>{p.name}</span><span>{number(p.value/currencyTotal*100)}%</span></div><div className="numeric-value currency-amount">{number(p.value)} USD</div></div>)}</div></div> : <div className="empty"><h3>Нет операций за выбранный период</h3></div>}</>},
    {id:'total-chart',title:"Динамика TOTAL",kind:'panel',content:<><h2>Динамика TOTAL</h2><p>Дата блока — A · дневной TOTAL — Y · USD</p>{valid.length ? <div style={{height:320}}><Chart data={{tooltip:{valueFormat:{type:'custom',formatter:({value})=>number(typeof value === 'number' ? value : null)}},yAxis:[{labels:{numberFormat:{precision:2,showRankDelimiter:true,lang:'ru',unit:null},style:{fontColor:'#3585f5'}}}],xAxis:{type:'category',categories:visible.map(p=>day(p.date!))},series:{data:[{type:'line',color:'#3585f5',name:'TOTAL, USD',data:visible.map((p,i)=>({x:i,y:p.issue ? null : p.total}))}]}}}/></div> : <div className="empty"><span className="chart-icon">▥</span><h3>Нет числовых итогов за выбранный период</h3><p>Загрузите Excel или измените диапазон дат.</p></div>}</>},
    {id:'daily-table',title:"Данные по дням",kind:'panel',content:<><h2>Данные по дням</h2><div className="table-scroll"><table><thead><tr><th>Дата</th><th>TOTAL, USD</th><th>FIXED USD AMOUNT</th><th>Строка Excel</th><th>Состояние</th></tr></thead><tbody>{visible.length ? [...visible].reverse().map(p=><tr key={p.row}><td>{day(p.date!)}</td><td className="numeric-value">{number(p.total)}</td><td className="numeric-value" title={p.fixedIssue ?? undefined}>{number(p.fixedUsdAmount)}</td><td>{p.row}</td><td>{p.issue ?? 'Сохранённый итог Excel'}</td></tr>) : <tr><td colSpan={5} className="empty-row">Данные пока не загружены или отсутствуют за период</td></tr>}</tbody></table></div></>},
    {id:'date-issues',title:"Проверьте даты в Excel",kind:'panel',content:<><h2>Проверьте даты в Excel</h2><p>В этих блоках встречаются разные даты. Итоги исключены из графика до уточнения.</p>{issues.filter(p=>!p.date).map(p=><p key={p.row}>Строка {p.row}: {p.issue}</p>)}</>},
  ];
  const sectionCards = cards.filter(card => section === 'overview' || ['daily-table','date-issues'].includes(card.id) || (section === 'total' ? card.id.startsWith('total-') : section === 'rates' ? card.id.startsWith('rate-') : card.id.startsWith('fixed-') || card.id === 'currencies'));
  return <ThemeProvider theme={theme}><div className={signedIn?'shell':'auth-shell'}>
    {signedIn && <aside id="dashboard-sidebar"><div className="brand"><span className="logo"><img src="./netting-logo.svg" width="36" height="36" alt=""/></span> Netting<span className="subtitle">Dashboard</span></div><nav className="sidebar-nav" aria-label="Разделы дашборда">{sections.map(item=><button key={item.id} className={`nav ${section===item.id ? 'active' : ''}`} aria-current={section===item.id ? 'page' : undefined} onClick={()=>setSection(item.id)}>{item.label}</button>)}</nav><div className="source"><span className="dot"/> {data ? desktop ? 'Локальный Excel' : 'Excel из OneDrive' : 'Источник не выбран'}<br/><small>Обновление раз в минуту</small><br/><small>{desktop ? 'Локальное приложение' : 'Microsoft 365'}</small></div></aside>}
    <main className={signedIn?undefined:'auth-main'}>{!signedIn && <div className="auth-brand"><img src="./netting-logo.svg" width="36" height="36" alt=""/><span>Netting Dashboard</span></div>}{signedIn && <header><div><div className="eyebrow">ПАНЕЛЬ РУКОВОДИТЕЛЯ</div><h1>Netting Dashboard</h1><p>{sections.find(item=>item.id===section)!.label} · RUB Deals Rates</p></div><div className="header-right"><div className="powered-by">Powered by LSD</div><div className="actions header-actions">{desktop && <><Button size="l" view="action" loading={busy} onClick={() => void load()}>Открыть Excel</Button><Button size="l" disabled={!data || busy} onClick={() => void load(true)}>Обновить</Button></>}<Button view="flat" size="l" className="icon-button" title={theme==='light'?'Тёмная тема':'Светлая тема'} aria-label={theme==='light'?'Тёмная тема':'Светлая тема'} onClick={()=>setTheme(theme==='light'?'dark':'light')}><AppIcon name={theme==='light'?'moon':'sun'}/></Button><div ref={setToolsElement} className="header-menu-slot"/></div></div></header>}
    {!desktop && <React.Suspense fallback={<p role="status">Подготовка входа…</p>}><OneDriveSource onData={setData} onBusy={setBusy} onSessionChange={setSignedIn} toolsElement={toolsElement} navigation={sections} section={section} onSection={id=>setSection(id as Section)}/></React.Suspense>}
    {error && <div role="alert" className="notice">{error} {data && 'Показаны данные последнего успешного чтения.'}</div>}
    {(desktop || data) && <>
    {(desktop || issues.length>0) && <div className="notice">{data ? `${data.name} · Прочитан ${new Date(data.updated).toLocaleTimeString('ru-RU')} · ${issues.length} строк TOTAL с пропуском или ошибкой` : 'Выберите Excel. Дата берётся из блока операций, итог — из Y строки TOTAL.'}</div>}
    <div className="filters"><DateFilter key={`from-${filterReset}`} label="С даты" value={from} onChange={setFrom}/><DateFilter key={`to-${filterReset}`} label="По дату" value={to} onChange={setTo}/><Button onClick={()=>{setFrom('');setTo('');setFilterReset(value=>value+1);}}>Все даты</Button></div>
    {from && to && from > to && <p role="alert">Начальная дата позже конечной.</p>}
    <DashboardGrid key={section} cards={sectionCards} loading={busy} layoutId={section} titles={titles} onRename={rename}/>
    </>}
    </main></div></ThemeProvider>;
}
createRoot(document.getElementById('root')!).render(<App/>);
