import React from 'react';
import {Button, TextInput} from '@gravity-ui/uikit';
import {auth, initialized, signIn, listItems, loadItem, restoreSource, saveSource, searchSharedFiles, itemKey, OneDriveError, type Item} from './oneDrive';
import type {Data} from './data';

export function OneDriveSource({onData,onBusy}: {onData:(data:Data|null)=>void;onBusy:(busy:boolean)=>void}) {
  const [ready,setReady] = React.useState(false);
  const [account,setAccount] = React.useState('');
  const [source,setSource] = React.useState<Item|null>(null);
  const [items,setItems] = React.useState<Item[]>([]);
  const [trail,setTrail] = React.useState<Item[]>([]);
  const [mode,setMode] = React.useState<'own'|'shared'>('own');
  const [query,setQuery] = React.useState('RUB RATES LOG');
  const [searchedQuery,setSearchedQuery] = React.useState('');
  const [more,setMore] = React.useState(false);
  const [offset,setOffset] = React.useState(0);
  const [picker,setPicker] = React.useState(false);
  const [busy,setBusy] = React.useState(false);
  const [error,setError] = React.useState('');
  const [needsSignIn,setNeedsSignIn] = React.useState(false);
  const [retryItem,setRetryItem] = React.useState<Item|null>(null);
  const [checked,setChecked] = React.useState('');
  const cache = React.useRef<{id:string;tag:string;data:Data}|undefined>(undefined);
  const working = React.useRef(false);
  const generation = React.useRef(0);
  const refresh = React.useCallback(async (item:Item) => {
    if (working.current) return;
    working.current=true; setBusy(true); onBusy(true);setError('');setNeedsSignIn(false);setRetryItem(null);
    const current = generation.current;
    try {
      const previous = cache.current?.id === itemKey(item) ? cache.current : undefined;
      const result = await loadItem(item,previous);
      if (current !== generation.current) return;
      cache.current={id:itemKey(item),tag:result.tag,data:result.data};
      onData(result.data); setSource(item); saveSource(item); setPicker(false); setError('');
      setChecked(new Date().toLocaleTimeString('ru-RU'));
    } catch(e) {if (current === generation.current) {setError(e instanceof Error ? e.message : 'Не удалось прочитать Excel.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);setRetryItem(item);}}
    finally {working.current=false; setBusy(false); onBusy(false);}
  },[onData,onBusy]);
  const browse = async (next:Item[]) => {
    if (working.current) return;
    working.current=true;setMode('own');setBusy(true);setPicker(true);setError('');setNeedsSignIn(false);setRetryItem(null);setItems([]);
    try {setItems(await listItems(next.at(-1)?.id));setTrail(next);}
    catch(e) {setError(e instanceof Error ? e.message : 'Не удалось открыть папку.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);}
    finally {working.current=false;setBusy(false);}
  };
  const search = async (nextOffset=0) => {
    if (working.current) return;
    working.current=true;setBusy(true);setMode('shared');setPicker(true);setError('');setNeedsSignIn(false);setRetryItem(null);
    if (!nextOffset) setItems([]);
    const term=nextOffset ? searchedQuery : query;
    try {
      const result=await searchSharedFiles(term,nextOffset);
      setItems(previous=>nextOffset ? [...new Map([...previous,...result.items].map(item=>[itemKey(item),item])).values()] : result.items);
      setMore(result.more);setOffset(nextOffset+100);setSearchedQuery(term);
    } catch(e) {setError(e instanceof Error ? e.message : 'Не удалось найти общие файлы.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);}
    finally {working.current=false;setBusy(false);}
  };
  React.useEffect(()=>{
    let disposed=false;
    void initialized.then(()=>{
      if(disposed)return;
      setReady(true);setAccount(auth.getActiveAccount()?.username ?? '');
      const item=restoreSource(); if(item && auth.getActiveAccount()) {setSource(item);void refresh(item);}
    }).catch(()=>{if(!disposed){setReady(true);setError('Не удалось подготовить вход Microsoft. Обновите страницу.');}});
    return()=>{disposed=true;};
  },[refresh]);
  React.useEffect(()=>{
    if(!source || !account)return;
    const check=()=>{if(document.visibilityState==='visible' && navigator.onLine)void refresh(source);};
    const timer=setInterval(check,60000);
    window.addEventListener('online',check);document.addEventListener('visibilitychange',check);
    return()=>{clearInterval(timer);window.removeEventListener('online',check);document.removeEventListener('visibilitychange',check);};
  },[source,account,refresh]);
  const login=async()=>{setError('');try{await signIn();}catch{setError('Не удалось начать вход Microsoft. Попробуйте ещё раз.');}};
  const logout=async()=>{
    generation.current++;cache.current=undefined;onData(null);setSource(null);setAccount('');setItems([]);setPicker(false);setError('');
    try{await auth.logoutRedirect({account:auth.getActiveAccount() ?? undefined});}catch{setError('Не удалось завершить выход Microsoft. Обновите страницу.');}
  };
  return <section className={`cloud-source ${!account ? 'welcome' : ''}`} aria-label="Подключение OneDrive">
    {!account ? <><h2>Ваш дашборд из OneDrive</h2><p>Войдите с рабочей учётной записью Microsoft 365 и выберите Excel. Данные обрабатываются на этом устройстве.</p><Button size="l" view="action" disabled={!ready} onClick={()=>void login()}>Войти через Microsoft</Button></> : <>
      <div className="cloud-toolbar"><div><b>OneDrive · {account}</b><p>{source ? source.name : 'Свой Excel — «Выбрать Excel», книга коллеги — «Общие файлы»'}</p>{checked && <small>Проверено в {checked} · обновление раз в минуту</small>}</div><div className="actions">
      <Button disabled={busy} onClick={()=>void browse([])}>{source ? 'Мой OneDrive' : 'Выбрать Excel'}</Button>
      <Button disabled={busy} onClick={()=>void search()}>Общие файлы</Button>
      {source && <Button disabled={busy} onClick={()=>void refresh(source)} loading={busy}>Обновить</Button>}
      <Button disabled={busy} onClick={()=>void logout()}>Выйти</Button></div></div>
      {picker && <div className="file-picker">
        <div className="actions"><Button disabled={busy} onClick={()=>void browse([])}>Мой OneDrive</Button><Button disabled={busy} onClick={()=>void search()}>Общие файлы</Button>{mode==='own' && trail.length>0 && <Button disabled={busy} onClick={()=>void browse(trail.slice(0,-1))}>Назад</Button>}<span>{mode==='own' ? trail.map(i=>i.name).join(' / ') : 'Excel, доступные вашей учётной записи'}</span><Button disabled={busy} onClick={()=>setPicker(false)}>Закрыть</Button></div>
        {mode==='shared' && <form className="shared-search" onSubmit={e=>{e.preventDefault();void search();}}><TextInput size="l" placeholder="Название или часть названия Excel" controlProps={{'aria-label':'Название общего Excel'}} value={query} onUpdate={setQuery}/><Button type="submit" disabled={busy}>Найти</Button></form>}
        {busy && <p role="status">Загрузка…</p>}
        {items.length>0 ? <ul>{items.map(item=><li key={itemKey(item)}><button disabled={busy} onClick={()=>void (item.folder ? browse([...trail,item]) : refresh(item))}><span aria-hidden="true">{item.folder ? '📁' : '▦'}</span> {item.name}{item.folder ? ' /' : ''}{mode==='shared' && item.webUrl && <small className="file-location">{item.webUrl}</small>}</button></li>)}</ul> : !busy && <p>{mode==='own' ? 'В этой папке нет папок или файлов .xlsx. Если книгу предоставил коллега, откройте «Общие файлы».' : 'Книги не найдены. Уточните название. Недавно предоставленный файл может появиться в поиске с задержкой.'}</p>}
        {mode==='shared' && more && <Button disabled={busy} onClick={()=>void search(offset)}>Показать ещё</Button>}
      </div>}
    </>}
    {error && <div className="cloud-error" role="alert"><p>{error} {cache.current && 'Показаны данные последнего успешного чтения.'}</p>{account && needsSignIn ? <Button disabled={busy} onClick={()=>void login()}>Войти снова</Button> : retryItem && <Button disabled={busy} onClick={()=>void refresh(retryItem)}>Повторить загрузку</Button>}</div>}
    {source && !cache.current && !busy && !picker && <Button onClick={()=>void browse([])}>Выбрать другую книгу</Button>}
  </section>;
}
