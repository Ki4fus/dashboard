import {PublicClientApplication, InteractionRequiredAuthError, BrowserCacheLocation} from '@azure/msal-browser';
import {readDashboard} from './excel';
import type {Data} from './data';

const scopes = ['Files.Read'];
const redirectUri = new URL('./', window.location.href).href;
export const auth = new PublicClientApplication({
  auth: {clientId: '1c879ee5-215f-44d5-b3bf-db4c8431a21e', authority: 'https://login.microsoftonline.com/63c53981-5871-490a-8cf5-c6495a69d62c', redirectUri, postLogoutRedirectUri: redirectUri},
  cache: {cacheLocation: BrowserCacheLocation.SessionStorage},
});
export const initialized = (async () => {
  await auth.initialize();
  const result = await auth.handleRedirectPromise();
  if (result?.account) auth.setActiveAccount(result.account);
  else if (!auth.getActiveAccount()) auth.setActiveAccount(auth.getAllAccounts()[0] ?? null);
})();
export async function signIn() {
  await initialized;
  await auth.loginRedirect({scopes, account: auth.getActiveAccount() ?? undefined});
}
export function sourceKey() {return `netting-source:${auth.getActiveAccount()?.homeAccountId ?? ''}`;}
export type Item = {id: string; name: string; folder?: object; file?: object; eTag?: string; cTag?: string; size?: number; '@microsoft.graph.downloadUrl'?: string};
export function restoreSource(): Item | null {
  try {const item = JSON.parse(localStorage.getItem(sourceKey()) ?? 'null'); return item && typeof item.id === 'string' && typeof item.name === 'string' ? {id:item.id,name:item.name} : null;} catch {return null;}
}
export function saveSource(item: Item) {try {localStorage.setItem(sourceKey(), JSON.stringify({id:item.id,name:item.name}));} catch {/* Storage is optional. */}}
export async function graph<T>(path: string): Promise<T> {
  await initialized;
  const account = auth.getActiveAccount();
  if (!account) throw new Error('Войдите через Microsoft.');
  let token;
  try {token = await auth.acquireTokenSilent({scopes, account});}
  catch (e) {if (e instanceof InteractionRequiredAuthError) throw new Error('Сеанс Microsoft истёк. Нажмите «Войти снова».'); throw e;}
  const url = new URL(path.startsWith('https:') ? path : `https://graph.microsoft.com/v1.0${path}`);
  if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) throw new Error('Недопустимый адрес Microsoft Graph.');
  const response = await fetch(url, {headers:{Authorization:`Bearer ${token.accessToken}`},cache:'no-store'});
  if (!response.ok) {
    if (response.status === 401) throw new Error('Сеанс Microsoft истёк. Нажмите «Войти снова».');
    if (response.status === 403) throw new Error('Microsoft запретил чтение. Проверьте Files.Read и согласие вашей организации.');
    if (response.status === 404) throw new Error('Файл или папка недоступны. Выберите книгу заново в своём OneDrive.');
    if (response.status === 429) throw new Error('Microsoft временно ограничил запросы. Попробуйте обновить позже.');
    throw new Error(`Не удалось прочитать OneDrive (${response.status}).`);
  }
  return response.json();
}
export async function listItems(folder?: string): Promise<Item[]> {
  let path: string | undefined = `/me/drive/${folder ? `items/${encodeURIComponent(folder)}` : 'root'}/children?$select=id,name,folder,file&$top=200`;
  const items: Item[] = [];
  while (path) {
    const result: {value: Item[]; '@odata.nextLink'?: string} = await graph(path);
    items.push(...result.value); path = result['@odata.nextLink'];
  }
  return items.filter(item=>item.folder || (item.file && /\.xlsx$/i.test(item.name))).sort((a,b)=>Number(Boolean(b.folder))-Number(Boolean(a.folder)) || a.name.localeCompare(b.name));
}
export async function loadItem(item: Item, cached?: {tag: string; data: Data}): Promise<{tag:string;data:Data;changed:boolean}> {
  const meta = await graph<Item>(`/me/drive/items/${encodeURIComponent(item.id)}?$select=id,name,eTag,cTag,size,@microsoft.graph.downloadUrl`);
  const tag = meta.cTag ?? meta.eTag ?? '';
  if (cached && tag && tag === cached.tag) return {...cached,changed:false};
  const download = meta['@microsoft.graph.downloadUrl'];
  if (!download || new URL(download).protocol !== 'https:') throw new Error('Microsoft не вернул адрес загрузки Excel.');
  const response = await fetch(download, {cache:'no-store',credentials:'omit',referrerPolicy:'no-referrer'});
  if (!response.ok) throw new Error('Не удалось загрузить Excel. Повторите обновление.');
  const data = await readDashboard(await response.arrayBuffer(), meta.name);
  return {tag,data,changed:true};
}
