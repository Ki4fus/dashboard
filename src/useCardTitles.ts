import React from 'react';
import {cardHelp} from './cardHelp';
const STORAGE_KEY='netting.dashboard.card-titles.v1';
export const MAX_CARD_TITLE_LENGTH=120;
export function useCardTitles() {
 const [titles,setTitles]=React.useState<Record<string,string>>(()=>{
  try {
   const saved:unknown=JSON.parse(localStorage.getItem(STORAGE_KEY)??'null');
   if(!saved || typeof saved!=='object' || Array.isArray(saved))return {};
   return Object.fromEntries(Object.entries(saved).filter(([id,title])=>Object.hasOwn(cardHelp,id) && typeof title==='string' && title.trim().length>0 && title.length<=MAX_CARD_TITLE_LENGTH).map(([id,title])=>[id,(title as string).trim()]));
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
