import React from 'react';
import {cardHelp} from './cardHelp';
const renamedTitles:Record<string,Record<string,string>>={
  "total-period": {
    "TOTAL \u00b7 \u0437\u0430 \u043f\u0435\u0440\u0438\u043e\u0434, USD": "Our P&L \u00b7 Period (USD)"
  },
  "total-latest": {
    "TOTAL \u00b7 \u043f\u043e\u0441\u043b\u0435\u0434\u043d\u0438\u0439 \u0434\u0435\u043d\u044c, USD": "Our P&L \u00b7 Today (USD)",
    "Our P&L \u00b7 Latest Day (USD)": "Our P&L \u00b7 Today (USD)"
  },
  "total-average": {
    "TOTAL \u00b7 \u0441\u0440\u0435\u0434\u043d\u0435\u0435 \u0437\u0430 \u0434\u0435\u043d\u044c, USD": "Our P&L \u00b7 Daily Average (USD)"
  },
  "total-change": {
    "TOTAL \u00b7 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u043a \u043f\u0440\u0435\u0434\u044b\u0434\u0443\u0449\u0435\u043c\u0443 \u0434\u043d\u044e": "Our P&L \u00b7 Daily Change (%)"
  },
  "fixed-period": {
    "FIXED USD AMOUNT \u00b7 \u0437\u0430 \u043f\u0435\u0440\u0438\u043e\u0434": "Fixed Amt \u00b7 Period (USD)",
    "Fixed USD \u00b7 Period": "Fixed Amt \u00b7 Period (USD)"
  },
  "fixed-today": {
    "FIXED USD AMOUNT \u00b7 \u0441\u0435\u0433\u043e\u0434\u043d\u044f": "Fixed Amt \u00b7 Today (USD)",
    "Fixed USD \u00b7 Today": "Fixed Amt \u00b7 Today (USD)"
  },
  "rate-today": {
    "\u0421\u0440\u0435\u0434\u043d\u0438\u0439 \u043a\u0443\u0440\u0441 \u0441\u0435\u0433\u043e\u0434\u043d\u044f": "Avg. Rate \u00b7 Today"
  },
  "rate-period": {
    "\u0421\u0440\u0435\u0434\u043d\u0438\u0439 \u043a\u0443\u0440\u0441 \u0437\u0430 \u043f\u0435\u0440\u0438\u043e\u0434": "Avg. Rate \u00b7 Period"
  },
  "rate-change": {
    "\u0418\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u043a\u0443\u0440\u0441\u0430 \u043a \u043f\u0440\u0435\u0434\u044b\u0434\u0443\u0449\u0435\u043c\u0443 \u0434\u043d\u044e": "Rate \u00b7 Daily Change (%)"
  },
  "total-chart": {
    "\u0414\u0438\u043d\u0430\u043c\u0438\u043a\u0430 TOTAL": "Trend of our P&L \u00b7 Today (USD)"
  },
  "currencies": {
    "\u0420\u0430\u0437\u0431\u0438\u0432\u043a\u0430 \u043f\u043e \u0432\u0430\u043b\u044e\u0442\u0430\u043c": "Currency breakdown"
  },
  "daily-table": {
    "\u0414\u0430\u043d\u043d\u044b\u0435 \u043f\u043e \u0434\u043d\u044f\u043c": "Daily data"
  },
  "date-issues": {
    "\u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0434\u0430\u0442\u044b \u0432 Excel": "Check Excel dates"
  }
};
const STORAGE_KEY='netting.dashboard.card-titles.v1';
export const MAX_CARD_TITLE_LENGTH=120;
export function useCardTitles() {
 const [titles,setTitles]=React.useState<Record<string,string>>(()=>{
  try {
   const saved:unknown=JSON.parse(localStorage.getItem(STORAGE_KEY)??'null');
   if(!saved || typeof saved!=='object' || Array.isArray(saved))return {};
   return Object.fromEntries(Object.entries(saved).filter(([id,title])=>Object.hasOwn(cardHelp,id) && typeof title==='string' && title.trim().length>0 && title.length<=MAX_CARD_TITLE_LENGTH).map(([id,title])=>[id,renamedTitles[id]?.[(title as string).trim()] ?? (title as string).trim()]));
  }catch{return {};}
 });
 const rename=React.useCallback((id:string,title:string|null)=>{
  if(!Object.hasOwn(cardHelp,id))return;
  setTitles(current=>{
   const next={...current};
   if(title===null)delete next[id];else {const value=title.trim();if(!value || value.length>MAX_CARD_TITLE_LENGTH)return current;next[id]=value;}
   try{localStorage.setItem(STORAGE_KEY,JSON.stringify(next));}catch{/* Keep names for the current session when persistent storage is unavailable. */}
   return next;
  });
 },[]);
 return {titles,rename};
}
