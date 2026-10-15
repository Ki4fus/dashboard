import React from 'react';
type Name = 'menu'|'sun'|'moon'|'grip'|'calendar'|'chevronLeft'|'chevronRight'|'folder'|'file'|'refresh'|'cloud'|'logout'|'close';
export function AppIcon({name,size=20}: {name:Name;size?:number}) {
 const paths: Record<Exclude<Name,'grip'>,React.ReactNode> = {
  menu:<><path d="M4 6h16M4 12h16M4 18h16"/></>,
  sun:<><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.5 1.5m11 11L19 19M5 19l1.5-1.5m11-11L19 5"/></>,
  moon:<path d="M20.5 13.3A8.7 8.7 0 0 1 10.7 3.5 8.7 8.7 0 1 0 20.5 13.3Z"/>,
  calendar:<><rect x="3" y="5" width="18" height="16" rx="3"/><path d="M7 3v4m10-4v4M3 11h18m-13 5h2m4 0h2"/></>,
  chevronLeft:<path d="m14 5-7 7 7 7"/>,chevronRight:<path d="m10 5 7 7-7 7"/>,
  folder:<path d="M3 7a2 2 0 0 1 2-2h5l2 2h7a2 2 0 0 1 2 2v10H3Z"/>,
  file:<><path d="M14 3H5v18h14V8Z"/><path d="M14 3v5h5M8 12h8m-8 4h8"/></>,
  refresh:<><path d="M20 11a8 8 0 0 0-14-5L3 9m0-5v5h5M4 13a8 8 0 0 0 14 5l3-3m0 5v-5h-5"/></>,
  cloud:<path d="M6 18a4 4 0 0 1-.6-8 7 7 0 0 1 13-1A4.5 4.5 0 0 1 18 18Z"/>,
  logout:<><path d="M10 4H4v16h6m4-12 4 4-4 4m-6-4h14"/></>,
  close:<path d="m6 6 12 12M6 18 18 6"/>,
 };
 return <svg aria-hidden="true" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" focusable="false">{name==='grip' ? [7,12,17].flatMap(y=>[9,15].map(x=><circle key={`${x}-${y}`} cx={x} cy={y} r="1.25" fill="currentColor" stroke="none"/>)) : paths[name]}</svg>;
}
