import React from 'react';
import {TextInput} from '@gravity-ui/uikit';
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
export function DateFilter({label,value,onChange}: {label:string;value:string;onChange:(value:string)=>void}) {
 const [draft,setDraft] = React.useState(formatDate(value));
 const [touched,setTouched] = React.useState(false);
 const calendar = React.useRef<HTMLInputElement>(null);
 React.useEffect(()=>{setDraft(formatDate(value));setTouched(false);},[value]);
 const invalid = touched && parseDate(draft)===null;
 return <div className="date-filter"><TextInput size="l" label={label} value={draft} placeholder="дд/мм/гггг" controlProps={{'aria-label':label,maxLength:10}} validationState={invalid ? 'invalid' : undefined} errorMessage={invalid ? 'Введите существующую дату: дд/мм/гггг' : undefined} onBlur={()=>setTouched(true)} onUpdate={text=>{setDraft(text);const parsed=parseDate(text);if(parsed!==null){onChange(parsed);setTouched(false);}}} endContent={<button type="button" className="calendar-button" aria-label={`Выбрать дату: ${label}`} onClick={()=>{try{calendar.current?.showPicker();}catch{calendar.current?.focus();}}}>▦</button>}/><input ref={calendar} className="calendar-native" type="date" tabIndex={-1} aria-hidden="true" value={value} onChange={e=>{onChange(e.target.value);setDraft(formatDate(e.target.value));setTouched(false);}}/></div>;
}
