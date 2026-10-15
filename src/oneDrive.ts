import {PublicClientApplication, InteractionRequiredAuthError, BrowserCacheLocation} from '@azure/msal-browser';
import {readDashboard} from './excel';
import type {Data} from './data';

const scopes = ['Files.Read.All'];
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
export type Item = {id: string; name: string; driveId?: string; webUrl?: string; folder?: object; file?: object; eTag?: string; cTag?: string; size?: number; '@microsoft.graph.downloadUrl'?: string};
export function restoreSource(): Item | null {
  try {const item = JSON.parse(localStorage.getItem(sourceKey()) ?? 'null'); return item && typeof item.id === 'string' && typeof item.name === 'string' ? {id:item.id,name:item.name,...(typeof item.driveId === 'string' ? {driveId:item.driveId} : {})} : null;} catch {return null;}
}
export function saveSource(item: Item) {try {localStorage.setItem(sourceKey(), JSON.stringify({id:item.id,name:item.name,...(item.driveId ? {driveId:item.driveId} : {})}));} catch {/* Storage is optional. */}}
export class OneDriveError extends Error {
  constructor(message: string, public requiresSignIn = false) {super(message);this.name='OneDriveError';}
}
export async function graph<T>(path: string, body?: unknown): Promise<T> {
  await initialized;
  const account = auth.getActiveAccount();
  if (!account) throw new OneDriveError('Войдите через Microsoft.',true);
  let token;
  try {token = await auth.acquireTokenSilent({scopes, account});}
  catch (e) {if (e instanceof InteractionRequiredAuthError) throw new OneDriveError('Требуется подтверждение входа Microsoft. Нажмите «Войти снова».',true); throw e;}
  const url = new URL(path.startsWith('https:') ? path : `https://graph.microsoft.com/v1.0${path}`);
  if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) throw new Error('Недопустимый адрес Microsoft Graph.');
  const response = await fetch(url, {method: body ? 'POST' : 'GET',headers:{Authorization:`Bearer ${token.accessToken}`,...(body ? {'Content-Type':'application/json'} : {Prefer:'odata.include-annotations="*"'})},...(body ? {body:JSON.stringify(body)} : {}),cache:'no-store'});
  if (!response.ok) {
    if (response.status === 401) throw new OneDriveError('Требуется подтверждение входа Microsoft. Нажмите «Войти снова».',true);
    if (response.status === 403) throw new Error('Microsoft запретил чтение. Проверьте делегированное Files.Read.All и согласие вашей организации.');
    if (response.status === 404) throw new Error('Файл или папка недоступны. Проверьте доступ к книге или выберите её заново.');
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
export function itemKey(item: Item) {return `${item.driveId ?? 'me'}:${item.id}`;}
export async function loadItem(item: Item, cached?: {tag: string; data: Data}): Promise<{tag:string;data:Data;changed:boolean}> {
  let meta = await graph<Item>(`${item.driveId ? `/drives/${encodeURIComponent(item.driveId)}` : '/me/drive'}/items/${encodeURIComponent(item.id)}?$select=id,name,eTag,cTag,size,file,folder,remoteItem,@microsoft.graph.downloadUrl`);
  const tag = meta.cTag ?? meta.eTag ?? '';
  if (cached && tag && tag === cached.tag) return {...cached,changed:false};
  // Retry without $select when the download annotation is missing.
  // Do not assume missing metadata means that the user must sign in again.
  const path = `${item.driveId ? `/drives/${encodeURIComponent(item.driveId)}` : '/me/drive'}/items/${encodeURIComponent(item.id)}`;
  if (!meta['@microsoft.graph.downloadUrl']) meta = await graph<Item>(path);
  const remote = (meta as Item & {remoteItem?:{id?:string;parentReference?:{driveId?:string}}}).remoteItem;
  if (!meta['@microsoft.graph.downloadUrl'] && remote?.id && remote.parentReference?.driveId) {
    meta = await graph<Item>(`/drives/${encodeURIComponent(remote.parentReference.driveId)}/items/${encodeURIComponent(remote.id)}`);
  }
  const download = meta['@microsoft.graph.downloadUrl'];
  let validDownload = false;
  try {validDownload = typeof download === 'string' && new URL(download).protocol === 'https:';} catch {/* Invalid URLs must never be requested. */}
  if (!validDownload || !download) throw new OneDriveError('Microsoft вернул сведения о книге, но не предоставил адрес её скачивания. Повторите загрузку. Если ошибка сохраняется, проверьте, разрешено ли скачивание этой книги в OneDrive и политиками организации.');
  const response = await fetch(download, {cache:'no-store',credentials:'omit',referrerPolicy:'no-referrer'});
  if (!response.ok) throw new Error('Не удалось загрузить Excel. Повторите обновление.');
  const data = await readDashboard(await response.arrayBuffer(), meta.name);
  return {tag,data,changed:true};
}

// Search is delegated and security-trimmed by Microsoft to files the signed-in user can read.
export async function searchSharedFiles(query: string, from = 0): Promise<{items:Item[];more:boolean}> {
  const phrase = query.trim().replace(/["\\\r\n]/g, ' ').slice(0,200);
  type Resource = Partial<Item> & {parentReference?:{driveId?:string;sharepointIds?:{listItemUniqueId?:string}}};
  type Container = {hits?:{hitId:string;resource:Resource}[];moreResultsAvailable?:boolean};
  const result = await graph<{value:{hitsContainers?:Container[]}[]}>('/search/query', {requests:[{entityTypes:['driveItem'],query:{queryString:`filetype:xlsx${phrase ? ` AND "${phrase}"` : ''}`},from,size:100,fields:['id','name','webUrl','parentReference','file']}]});
  const containers = result.value.flatMap(value=>value.hitsContainers ?? []);
  const items:Item[] = [];
  for (const hit of containers.flatMap(container=>container.hits ?? [])) {
    const resource = hit.resource;
    const id = resource.id ?? resource.parentReference?.sharepointIds?.listItemUniqueId ?? hit.hitId;
    const driveId = resource.parentReference?.driveId;
    if (!id || !driveId || !resource.name || !/\.xlsx$/i.test(resource.name)) continue;
    items.push({id,name:resource.name,driveId,webUrl:resource.webUrl,file:{}});
  }
  return {items:[...new Map(items.map(item=>[itemKey(item),item])).values()],more:containers.some(container=>container.moreResultsAvailable)};
}
