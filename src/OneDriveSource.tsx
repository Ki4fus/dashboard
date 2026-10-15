import React from 'react';
import {Button} from '@gravity-ui/uikit';
import {auth, initialized, signIn, listItems, loadItem, restoreSource, saveSource, type Item} from './oneDrive';
import type {Data} from './data';

export function OneDriveSource({onData,onBusy}: {onData:(data:Data|null)=>void;onBusy:(busy:boolean)=>void}) {
  const [ready,setReady] = React.useState(false);
  const [account,setAccount] = React.useState('');
  const [source,setSource] = React.useState<Item|null>(null);
  const [items,setItems] = React.useState<Item[]>([]);
  const [trail,setTrail] = React.useState<Item[]>([]);
  const [picker,setPicker] = React.useState(false);
  const [busy,setBusy] = React.useState(false);
  const [error,setError] = React.useState('');
  const [checked,setChecked] = React.useState('');
  const cache = React.useRef<{id:string;tag:string;data:Data}|undefined>(undefined);
  const working = React.useRef(false);
  const generation = React.useRef(0);
  const refresh = React.useCallback(async (item:Item) => {
    if (working.current) return;
    working.current=true; setBusy(true); onBusy(true);
    const current = generation.current;
    try {
      const previous = cache.current?.id === item.id ? cache.current : undefined;
      const result = await loadItem(item,previous);
      if (current !== generation.current) return;
      cache.current={id:item.id,tag:result.tag,data:result.data};
      onData(result.data); setSource(item); saveSource(item); setPicker(false); setError('');
      setChecked(new Date().toLocaleTimeString('ru-RU'));
    } catch(e) {if (current === generation.current) setError(e instanceof Error ? e.message : 'Не удалось прочитать Excel.');}
    finally {working.current=false; setBusy(false); onBusy(false);}
  },[onData,onBusy]);
  const browse = async (next:Item[]) => {
    if (working.current) return;
    working.current=true;setBusy(true);setPicker(true);setError('');setItems([]);
    try {setItems(await listItems(next.at(-1)?.id));setTrail(next);}
    catch(e) {setError(e instanceof Error ? e.message : 'Не удалось открыть папку.');}
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
      <div className="cloud-toolbar"><div><b>OneDrive · {account}</b><p>{source ? source.name : 'Выберите книгу .xlsx с листом RUB Deals Rates'}</p>{checked && <small>Проверено в {checked} · обновление раз в минуту</small>}</div><div className="actions">
      <Button disabled={busy} onClick={()=>void browse([])}>{source ? 'Сменить Excel' : 'Выбрать Excel'}</Button>
      {source && <Button disabled={busy} onClick={()=>void refresh(source)} loading={busy}>Обновить</Button>}
      <Button disabled={busy} onClick={()=>void logout()}>Выйти</Button></div></div>
      {picker && <div className="file-picker"><div className="actions"><Button disabled={busy} onClick={()=>void browse([])}>Мой OneDrive</Button>{trail.length>0 && <Button disabled={busy} onClick={()=>void browse(trail.slice(0,-1))}>Назад</Button>}<span>{trail.map(i=>i.name).join(' / ')}</span><Button disabled={busy} onClick={()=>setPicker(false)}>Закрыть</Button></div>
      {busy ? <p role="status">Загрузка…</p> : items.length ? <ul>{items.map(item=><li key={item.id}><button disabled={busy} onClick={()=>void (item.folder ? browse([...trail,item]) : refresh(item))}><span aria-hidden="true">{item.folder ? '📁' : '▦'}</span> {item.name}{item.folder ? ' /' : ''}</button></li>)}</ul> : <p>В этой папке нет папок или файлов .xlsx.</p>}</div>}
    </>}
    {error && <div className="cloud-error" role="alert"><p>{error} {cache.current && 'Показаны данные последнего успешного чтения.'}</p>{account && <Button onClick={()=>void login()}>Войти снова</Button>}</div>}
    {source && !cache.current && !busy && !picker && <Button onClick={()=>void browse([])}>Выбрать другую книгу</Button>}
  </section>;
}
