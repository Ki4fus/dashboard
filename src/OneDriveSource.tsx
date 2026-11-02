import React from 'react';
import {Button, TextInput, DropdownMenu, Modal} from '@gravity-ui/uikit';
import {auth, initialized, signIn, listItems, loadItem, restoreSource, saveSource, searchSharedFiles, itemKey, OneDriveError, type Item} from './oneDrive';
import {createPortal} from 'react-dom';
import {AppIcon} from './Icons';
import {SharedCatalog} from './SharedCatalog';
import type {Data} from './data';

export function OneDriveSource({onData,onBusy,toolsElement,navigation,section,onSection,onSessionChange}: {onData:(data:Data|null)=>void;onBusy:(busy:boolean)=>void;toolsElement:HTMLElement|null;onSessionChange:(signedIn:boolean)=>void;navigation:{id:string;label:string}[];section:string;onSection:(id:string)=>void}) {
  const [ready,setReady] = React.useState(false);
  const [account,setAccount] = React.useState('');
  const [source,setSource] = React.useState<Item|null>(null);
  const [items,setItems] = React.useState<Item[]>([]);
  const [trail,setTrail] = React.useState<Item[]>([]);
  const [mode,setMode] = React.useState<'own'|'shared'>('own');
  const [query,setQuery] = React.useState('');
  const [menuOpen,setMenuOpen] = React.useState(false);
  const [confirmLogout,setConfirmLogout] = React.useState(false);
  const [catalogVersion,setCatalogVersion] = React.useState(0);
  const cancelLogout=React.useRef<HTMLButtonElement>(null);
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
      setChecked(new Date().toLocaleTimeString('en-GB'));
    } catch(e) {if (current === generation.current) {setError(e instanceof Error ? e.message : 'Could not read the Excel workbook.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);setRetryItem(item);}}
    finally {working.current=false; setBusy(false); onBusy(false);}
  },[onData,onBusy]);
  const browse = async (next:Item[]) => {
    if (working.current) return;
    working.current=true;setMode('own');setBusy(true);setPicker(true);setError('');setNeedsSignIn(false);setRetryItem(null);setItems([]);
    try {setItems(await listItems(next.at(-1)?.id));setTrail(next);}
    catch(e) {setError(e instanceof Error ? e.message : 'Could not open the folder.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);}
    finally {working.current=false;setBusy(false);}
  };
  const search = async (nextOffset=0) => {
    if (working.current) return;
    working.current=true;setBusy(true);setMode('shared');setPicker(true);setError('');setNeedsSignIn(false);setRetryItem(null);
    if (!nextOffset) {setItems([]);setCatalogVersion(value=>value+1);}
    const term=nextOffset ? searchedQuery : query;
    try {
      const result=await searchSharedFiles(term,nextOffset);
      setItems(previous=>nextOffset ? [...new Map([...previous,...result.items].map(item=>[itemKey(item),item])).values()] : result.items);
      setMore(result.more);setOffset(nextOffset+100);setSearchedQuery(term);
    } catch(e) {setError(e instanceof Error ? e.message : 'Could not search shared files.');setNeedsSignIn(e instanceof OneDriveError && e.requiresSignIn);}
    finally {working.current=false;setBusy(false);}
  };
  React.useEffect(()=>{
    let disposed=false;
    void initialized.then(()=>{
      if(disposed)return;
      setReady(true);setAccount(auth.getActiveAccount()?.username ?? '');onSessionChange(Boolean(auth.getActiveAccount()));
      const item=restoreSource(); if(item && auth.getActiveAccount()) {setSource(item);void refresh(item);}
    }).catch(()=>{if(!disposed){setReady(true);setError('Could not prepare Microsoft sign-in. Refresh the page.');}});
    return()=>{disposed=true;};
  },[refresh,onSessionChange]);
  React.useEffect(()=>{
    if(!source || !account)return;
    const check=()=>{if(document.visibilityState==='visible' && navigator.onLine)void refresh(source);};
    const timer=setInterval(check,60000);
    window.addEventListener('online',check);document.addEventListener('visibilitychange',check);
    return()=>{clearInterval(timer);window.removeEventListener('online',check);document.removeEventListener('visibilitychange',check);};
  },[source,account,refresh]);
  const login=async()=>{setError('');try{await signIn();}catch{setError('Could not start Microsoft sign-in. Try again.');}};
  const logout=async()=>{
    generation.current++;cache.current=undefined;onData(null);onSessionChange(false);setSource(null);setAccount('');setItems([]);setPicker(false);setError('');
    try{await auth.logoutRedirect({account:auth.getActiveAccount() ?? undefined});}catch{setError('Could not complete Microsoft sign-out. Refresh the page.');}
  };
  return <>
  {account && toolsElement && createPortal(<DropdownMenu defaultSwitcherClassName="icon-button" switcherWrapperClassName={account?undefined:'signed-out-menu'} size="l" open={menuOpen} onOpenToggle={setMenuOpen} icon={<AppIcon name="menu"/>} defaultSwitcherProps={{view:'flat','aria-label':'Open menu',title:'Menu'}} popupProps={{placement:'bottom-end'}} items={[
    navigation.map(item=>({text:item.label,action:()=>onSection(item.id),selected:section===item.id})),
    ...(account ? [[{text:'My OneDrive',iconStart:<AppIcon name="cloud" size={18}/>,disabled:busy,action:()=>void browse([])},{text:'Shared files',iconStart:<AppIcon name="folder" size={18}/>,disabled:busy,action:()=>void search()}],[{text:'Sign out',iconStart:<AppIcon name="logout" size={18}/>,disabled:busy,action:()=>setConfirmLogout(true)}]] : []),
  ]}/>,toolsElement)}
  <Modal open={confirmLogout} onOpenChange={setConfirmLogout} aria-labelledby="logout-title" initialFocus={cancelLogout} contentClassName="logout-dialog"><div className="dialog-symbol"><AppIcon name="logout" size={24}/></div><h2 id="logout-title">Sign out of the dashboard?</h2><p>Your current session data will be cleared. Sign in with Microsoft again to view the dashboard.</p><div className="dialog-actions"><Button ref={cancelLogout} size="l" onClick={()=>setConfirmLogout(false)}>Stay signed in</Button><Button size="l" view="action" onClick={()=>{setConfirmLogout(false);void logout();}}>Sign out</Button></div></Modal>
  <section className={`cloud-source ${!account ? 'welcome' : ''}`} aria-label="OneDrive connection">
    {!account ? <><h2>Your dashboard from OneDrive</h2><p>Sign in with your work Microsoft 365 account and select an Excel workbook. Data is processed on this device.</p><Button size="l" view="action" disabled={!ready} onClick={()=>void login()}>Sign in with Microsoft</Button></> : <>
      <div className="source-summary"><span className={`source-status ${busy?'checking':''}`} aria-hidden="true"/><div className="source-file"><span className="source-label">{source ? 'Excel from OneDrive' : 'Data source'}</span><b title={source?.name}>{source ? source.name : 'Select a workbook using the top-right menu'}</b></div><div className="source-account"><span className="source-label">Microsoft 365</span><span title={account}>{account}</span></div><div className="source-time"><span className="source-label">{busy?'Checking for updates…':checked ? 'Checked at '+checked : 'Checks every minute'}</span><span>{source ? 'Checks for changes every minute' : 'Excel is processed on this device'}</span></div>{source && <Button view="flat" className="icon-button" disabled={busy} onClick={()=>void refresh(source)} aria-label="Refresh" title="Refresh"><AppIcon name="refresh" size={18}/></Button>}</div>
      {picker && <div className="file-picker">
        <div className="picker-heading"><h2>{mode==='own'?'My OneDrive':'Shared files'}</h2><Button view="flat" className="icon-button" disabled={busy} onClick={()=>setPicker(false)} aria-label="Close file picker"><AppIcon name="close" size={18}/></Button></div>
        {mode==='own' && <div className="catalog-breadcrumbs"><button disabled={busy} onClick={()=>void browse([])}>My OneDrive</button>{trail.map((item,i)=><React.Fragment key={itemKey(item)}><span aria-hidden="true">/</span><button disabled={busy} onClick={()=>void browse(trail.slice(0,i+1))}>{item.name}</button></React.Fragment>)}{trail.length>0 && <Button disabled={busy} onClick={()=>void browse(trail.slice(0,-1))}>Back</Button>}</div>}
        {mode==='shared' && <form className="shared-search" onSubmit={e=>{e.preventDefault();void search();}}><TextInput size="l" placeholder="Workbook name or part of a name" controlProps={{'aria-label':'Shared workbook name'}} value={query} onUpdate={setQuery}/><Button type="submit" disabled={busy}>Search</Button></form>}
        {busy && <p role="status">Loading…</p>}
        {mode==='shared' && items.length>0 ? <SharedCatalog items={items} busy={busy} resetKey={catalogVersion} onSelect={item=>void refresh(item)}/> : mode==='own' && items.length>0 ? <ul className="catalog-list">{items.map(item=><li key={itemKey(item)}><button disabled={busy} onClick={()=>void (item.folder ? browse([...trail,item]) : refresh(item))}><span className={`catalog-icon ${item.folder?'':'file'}`}><AppIcon name={item.folder?'folder':'file'}/></span><span className="catalog-name">{item.name}</span>{item.folder && <AppIcon name="chevronRight" size={16}/>}</button></li>)}</ul> : !busy && <p>{mode==='own' ? 'No folders or .xlsx files here. Use Shared files in the menu to find colleagues\' workbooks.' : 'No workbooks found. Refine or clear your search. Recently shared workbooks may take time to appear.'}</p>}
        {mode==='shared' && more && <Button disabled={busy} onClick={()=>void search(offset)}>Show more</Button>}
      </div>}
    </>}
    {error && <div className="cloud-error" role="alert"><p>{error} {cache.current && 'Showing data from the last successful load.'}</p>{account && needsSignIn ? <Button disabled={busy} onClick={()=>void login()}>Sign in again</Button> : retryItem && <Button disabled={busy} onClick={()=>void refresh(retryItem)}>Retry download</Button>}</div>}
    {source && !cache.current && !busy && !picker && <Button onClick={()=>void browse([])}>Select another workbook</Button>}
  </section></>;
}
