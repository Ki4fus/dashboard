import React from 'react';
import {Button} from '@gravity-ui/uikit';
import {AppIcon} from './Icons';
import {itemKey,type Item} from './oneDrive';
// Build a folder view of readable search results without requesting access to whole drives.
export function catalogPath(item:Item):string[] {
 if(!item.webUrl)return [];
 try {
  const url=new URL(item.webUrl);
  const parts=url.pathname.split('/').filter(Boolean).map(p=>decodeURIComponent(p));
  const root=parts.findIndex(p=>['personal','sites','teams'].includes(p.toLowerCase()));
  if(root>=0 && parts[root+1])return [`${parts[root]==='personal'?'OneDrive':'SharePoint'} · ${parts[root+1]}`,...parts.slice(root+2,-1).map(p=>p==='Documents'?'Документы':p)];
  return [url.hostname,...parts.slice(0,-1)];
 }catch{return [];}
}
export function SharedCatalog({items,busy,onSelect,resetKey}:{items:Item[];busy:boolean;onSelect:(item:Item)=>void;resetKey:number}) {
 const [path,setPath]=React.useState<string[]>([]);
 React.useEffect(()=>setPath([]),[resetKey]);
 const rows=items.map(item=>({item,path:catalogPath(item)})).filter(entry=>path.every((part,index)=>entry.path[index]===part));
 const folders=[...new Set(rows.filter(entry=>entry.path.length>path.length).map(entry=>entry.path[path.length]))].sort((a,b)=>a.localeCompare(b));
 const files=rows.filter(entry=>entry.path.length===path.length).map(entry=>entry.item);
 return <div className="shared-catalog"><nav className="catalog-breadcrumbs" aria-label="Путь к общей книге"><button disabled={busy} onClick={()=>setPath([])}>Все папки</button>{path.map((part,i)=><React.Fragment key={i}><span aria-hidden="true">/</span><button disabled={busy} aria-current={i===path.length-1?'location':undefined} onClick={()=>setPath(path.slice(0,i+1))}>{part}</button></React.Fragment>)}</nav>
 {path.length>0 && <Button disabled={busy} onClick={()=>setPath(path.slice(0,-1))}>Назад</Button>}
 <ul className="catalog-list">{folders.map(folder=><li key={folder}><button disabled={busy} onClick={()=>setPath([...path,folder])}><span className="catalog-icon"><AppIcon name="folder"/></span><span className="catalog-name">{folder}</span><AppIcon name="chevronRight" size={16}/></button></li>)}{files.map(item=><li key={itemKey(item)}><button disabled={busy} onClick={()=>onSelect(item)}><span className="catalog-icon file"><AppIcon name="file"/></span><span className="catalog-name">{item.name}</span></button></li>)}</ul>
 {!rows.length && !busy && <p>В этой папке нет доступных книг.</p>}
 </div>;
}
