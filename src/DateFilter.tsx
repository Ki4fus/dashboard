import React from 'react';
import {Popup,TextInput} from '@gravity-ui/uikit';
import {AppIcon} from './Icons';
export const formatDate = (value: string) => value ? value.split('-').reverse().join('/') : '';
export function parseDate(text: string): string | null {
 if (!text.trim()) return '';
 const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text.trim());
 if (!match) return null;
 const [,dd,mm,yyyy] = match;
 const iso = `${yyyy}-${mm}-${dd}`;
 const date = new Date(iso+'T00:00:00Z');
 return Number(yyyy)>=1900 && Number(yyyy)<=9999 && Number.isFinite(date.getTime()) && date.toISOString().slice(0,10)===iso ? iso : null;
}
const isoDay=(date:Date)=>date.toISOString().slice(0,10);
const today=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
export function DateFilter({label,value,onChange}: {label:string;value:string;onChange:(value:string)=>void}) {
 const [draft,setDraft] = React.useState(formatDate(value));
 const [touched,setTouched] = React.useState(false);
 const [open,setOpen] = React.useState(false);
 const [cursor,setCursor] = React.useState(value || today());
 const [anchor,setAnchor] = React.useState<HTMLButtonElement|null>(null);
 const dayRefs=React.useRef(new Map<string,HTMLButtonElement>());
 const id=React.useId();
 React.useEffect(()=>{setDraft(formatDate(value));setTouched(false);},[value]);
 const focusDate=new Date(cursor+'T00:00:00Z');
 const year=focusDate.getUTCFullYear(),month=focusDate.getUTCMonth();
 const first=new Date(Date.UTC(year,month,1));
 const offset=(first.getUTCDay()+6)%7;
 const cells=Array.from({length:42},(_,i)=>new Date(Date.UTC(year,month,1-offset+i)));
 const monthLabel=new Intl.DateTimeFormat('ru-RU',{year:'numeric',month:'long',timeZone:'UTC'}).format(first);
 const invalid = touched && parseDate(draft)===null;
 function shiftMonth(delta:number){const last=new Date(Date.UTC(year,month+delta+1,0)).getUTCDate();const next=new Date(Date.UTC(year,month+delta,Math.min(focusDate.getUTCDate(),last)));if(next.getUTCFullYear()>=1900 && next.getUTCFullYear()<=9999)setCursor(isoDay(next));}
 function keyboard(event:React.KeyboardEvent){
  const delta:Record<string,number>={ArrowLeft:-1,ArrowRight:1,ArrowUp:-7,ArrowDown:7,Home:-((focusDate.getUTCDay()+6)%7),End:6-((focusDate.getUTCDay()+6)%7)};
  if(event.key==='PageUp'||event.key==='PageDown'){event.preventDefault();shiftMonth(event.key==='PageUp'?-1:1);return;}
  if(!(event.key in delta))return;event.preventDefault();const next=new Date(focusDate);next.setUTCDate(next.getUTCDate()+delta[event.key]);if(next.getUTCFullYear()>=1900 && next.getUTCFullYear()<=9999)setCursor(isoDay(next));
 }
 React.useEffect(()=>{if(open)dayRefs.current.get(cursor)?.focus();},[open,cursor]);
 function choose(next:string){onChange(next);setDraft(formatDate(next));setTouched(false);setOpen(false);anchor?.focus();}
 return <div className="date-filter"><TextInput size="l" label={label} value={draft} placeholder="дд/мм/гггг" controlProps={{'aria-label':label,maxLength:10}} validationState={invalid ? 'invalid' : undefined} errorMessage={invalid ? 'Введите существующую дату: дд/мм/гггг' : undefined} onBlur={()=>setTouched(true)} onUpdate={text=>{setDraft(text);const parsed=parseDate(text);if(parsed!==null){onChange(parsed);setTouched(false);}}} endContent={<button ref={setAnchor} type="button" className="calendar-button" aria-label={`Выбрать дату: ${label}`} aria-expanded={open} aria-controls={id} aria-haspopup="dialog" onClick={()=>{if(!open)setCursor(value||today());setOpen(!open);}}><AppIcon name="calendar" size={18}/></button>}/>
  <Popup open={open} anchorElement={anchor} onOpenChange={setOpen} placement={['bottom-start','top-start']} id={id} disableTransition initialFocus={-1} className="calendar-popup">
   <div className="date-calendar" role="dialog" aria-label={`Календарь: ${label}`}><div className="calendar-month"><button className="icon-button" aria-label="Предыдущий месяц" disabled={year===1900&&month===0} onClick={()=>shiftMonth(-1)}><AppIcon name="chevronLeft" size={16}/></button><span aria-live="polite">{monthLabel}</span><button className="icon-button" aria-label="Следующий месяц" disabled={year===9999&&month===11} onClick={()=>shiftMonth(1)}><AppIcon name="chevronRight" size={16}/></button></div>
   <div role="grid" aria-label={monthLabel} onKeyDown={keyboard}><div role="row" className="calendar-week">{['Пн','Вт','Ср','Чт','Пт','Сб','Вс'].map(d=><span role="columnheader" key={d}>{d}</span>)}</div>{Array.from({length:6},(_,week)=><div role="row" className="calendar-week" key={week}>{cells.slice(week*7,week*7+7).map(date=>{const iso=isoDay(date);return <div role="gridcell" aria-selected={value===iso} key={iso}><button type="button" ref={node=>{if(node)dayRefs.current.set(iso,node);else dayRefs.current.delete(iso);}} tabIndex={cursor===iso?0:-1} disabled={date.getUTCFullYear()<1900 || date.getUTCFullYear()>9999} aria-label={formatDate(iso)} aria-current={iso===today()?'date':undefined} className={`calendar-day ${date.getUTCMonth()!==month?'outside-month':''} ${value===iso?'selected':''}`} onClick={()=>choose(iso)}>{date.getUTCDate()}</button></div>;})}</div>)}</div>
   <div className="calendar-footer"><button onClick={()=>choose(today())}>Сегодня</button><button onClick={()=>choose('')}>Очистить</button></div></div>
  </Popup></div>;
}
