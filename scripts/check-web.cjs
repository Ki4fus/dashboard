const {chromium} = require('playwright');
const ExcelJS = require('exceljs');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const {readDashboard} = require('../electron/excel.cjs');
(async()=>{
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('RUB Deals Rates');
 const day=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 sheet.getCell('A1').value=new Date(day+'T00:00:00Z');sheet.getCell('E1').value=80;sheet.getCell('L1').value='USD';sheet.getCell('U1').value=200;
 sheet.getCell('A2').value='TOTAL';sheet.getCell('Y2').value={formula:'100+23',result:123};sheet.getCell('U2').value={formula:'100+100',result:200};
 const bytes=Buffer.from(await book.xlsx.writeBuffer());
 const fixture='/private/tmp/netting-web-fixture.xlsx';await fs.writeFile(fixture,bytes);
 const expected=await readDashboard(fixture);
 const browser=await chromium.launch({executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:true});
 try {
 const page=await browser.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
 let tag='version-1',downloads=0,forbidden=false,invalid=false,pages=0,sharedReads=0;
 const authMock=`export class InteractionRequiredAuthError extends Error {};
 export const BrowserCacheLocation={SessionStorage:'sessionStorage'};
 export class PublicClientApplication {
 constructor(){this.account=sessionStorage.getItem('test-login')?{username:'user@example.com',homeAccountId:'test-user'}:null;}
 async initialize(){} async handleRedirectPromise(){return null;} getActiveAccount(){return this.account;} setActiveAccount(a){this.account=a;} getAllAccounts(){return this.account?[this.account]:[];}
 async loginRedirect(){sessionStorage.setItem('test-login','1');location.reload();} async acquireTokenSilent(){return {accessToken:'test-token'};}
 async logoutRedirect(){sessionStorage.removeItem('test-login');location.reload();}
 }`;
 await page.route(/\/[^/?]*azure[^/?]*msal[^/?]*\.js(?:\?|$)/,r=>r.fulfill({contentType:'application/javascript',body:authMock}));
 await page.route('https://graph.microsoft.com/**',async route=>{
  assert.equal(route.request().headers().authorization,'Bearer test-token');
  const url=new URL(route.request().url());let body;
  if(forbidden)return route.fulfill({status:403,body:'{}'});
  if(url.pathname.endsWith('/search/query')) {assert.equal(route.request().method(),'POST');const request=route.request().postDataJSON().requests[0];assert.deepEqual(request.entityTypes,['driveItem']);assert.ok(request.query.queryString.includes('RUB RATES LOG'));body={value:[{hitsContainers:[{moreResultsAvailable:false,hits:[{hitId:'book',resource:{id:'book',name:'RUB RATES LOG.xlsx',webUrl:'https://example.sharepoint.com/personal/colleague/Documents/RUB%20RATES%20LOG.xlsx',parentReference:{driveId:'colleague-drive'},file:{}}}]}]}]};}
  else if(url.pathname.endsWith('/root/children'))body={value:[{id:'folder',name:'Reports',folder:{}}]};
  else if(url.pathname.endsWith('/folder/children')) {pages++;body=url.searchParams.has('page')?{value:[{id:'book',name:'Report.xlsx',file:{}}]}:{value:[], '@odata.nextLink':'https://graph.microsoft.com/v1.0/me/drive/items/folder/children?page=2'};}
  else {if(url.pathname.includes('/drives/colleague-drive/'))sharedReads++;body={id:'book',name:url.pathname.includes('/drives/')?'RUB RATES LOG.xlsx':'Report.xlsx',cTag:tag,'@microsoft.graph.downloadUrl':'https://download.example.com/book.xlsx'};}
  await route.fulfill({contentType:'application/json',body:JSON.stringify(body)});
 });
 await page.route('https://download.example.com/**',async route=>{downloads++;assert.equal(route.request().headers().authorization,undefined);await route.fulfill({body:invalid?Buffer.from('invalid excel'):bytes});});
 await page.goto('http://localhost:5173/');
 await page.getByRole('button',{name:'Войти через Microsoft',exact:true}).click();
 await page.getByRole('button',{name:'Выбрать Excel',exact:true}).click();
 await page.getByRole('button',{name:/Reports/}).click();await page.getByRole('button',{name:/Report.xlsx/}).click();
 await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();
 assert.equal(pages,2);assert.equal(downloads,1);
 assert.ok((await page.locator('.dashboard-card.metric').first().textContent()).includes('123,00'));
 const parsed=await page.evaluate(async base64=>{const {readDashboard}=await import('/src/excel.ts');return readDashboard(Uint8Array.from(atob(base64),c=>c.charCodeAt(0)).buffer,'Report.xlsx');},bytes.toString('base64'));
 assert.deepEqual({sheet:parsed.sheet,points:parsed.points,operations:parsed.operations},expected,'Browser parser matches desktop including cached formulas');
 await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,1,'Unchanged version is not downloaded again');
 tag='version-2';await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,2);
 forbidden=true;await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.getByRole('alert').waitFor();assert.ok((await page.getByRole('alert').textContent()).includes('Показаны данные последнего'));assert.equal(await page.locator('.dashboard-card').count(),13);forbidden=false;
 for(const width of [390,768,1440]){
 await page.setViewportSize({width,height:900});await page.waitForFunction(()=>document.documentElement.scrollWidth<=innerWidth,{},{timeout:5000});
 const overflow=await page.evaluate(()=>[...document.querySelectorAll('body *')].filter(e=>e.getBoundingClientRect().right>innerWidth+1 && getComputedStyle(e).position!=='absolute').map(e=>[e.tagName,e.className,e.getBoundingClientRect().right]).slice(0,20)); if(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth))console.log('Overflow',overflow);
 assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false,`No page overflow at ${width}`);
 const columns=await page.locator('.dashboard-grid').evaluate(e=>getComputedStyle(e).gridTemplateColumns.split(' ').length);assert.equal(columns,width===390?1:width===768?2:3);
 if(width===390){await page.getByRole('button',{name:'Открыть меню',exact:true}).click();await page.getByRole('button',{name:'Total',exact:true}).click();assert.equal(await page.locator('.dashboard-card').count(),7);await page.getByRole('button',{name:'Открыть меню',exact:true}).click();await page.getByRole('button',{name:'Обзор показателей',exact:true}).click();}
 await page.screenshot({path:`/private/tmp/netting-web-${width}.png`,fullPage:true});
 }
 const saved=await page.evaluate(()=>Object.fromEntries(Object.entries(localStorage).filter(([k])=>k.startsWith('netting-source:'))));assert.deepEqual(JSON.parse(Object.values(saved)[0]),{id:'book',name:'Report.xlsx'});
 await page.reload();await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();assert.equal(downloads,3,'Reload restores source and reloads data into memory');
 await page.getByRole('button',{name:'Общие файлы',exact:true}).click();
 await page.getByRole('button',{name:/RUB RATES LOG.xlsx/}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(sharedReads,1);assert.equal(downloads,4,'Same item id in another drive must not reuse own-drive cache');
 const sharedSaved=await page.evaluate(()=>JSON.parse(localStorage.getItem('netting-source:test-user')));assert.deepEqual(sharedSaved,{id:'book',name:'RUB RATES LOG.xlsx',driveId:'colleague-drive'});
 await page.reload();await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();assert.equal(downloads,5);assert.equal(sharedReads,2,'Restored shared source reads colleague drive');
 invalid=true;tag='version-3';await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.getByRole('alert').waitFor();assert.equal(await page.locator('.dashboard-card').count(),13);
 await page.getByRole('button',{name:'Выйти',exact:true}).click();await page.getByRole('button',{name:'Войти через Microsoft',exact:true}).waitFor();assert.equal(await page.locator('.dashboard-card').count(),0);
 assert.deepEqual(errors,[]);console.log('Web checks passed: sign-in UI, paginated picker, formula parity, version checks, error retention, own/shared source restoration, cross-drive cache separation, read-only shared search, logout and 390/768/1440 layouts. Authentication is mocked; real Microsoft sign-in needs user verification.');
 } finally {await browser.close();await fs.unlink(fixture);}
})().catch(e=>{console.error(e);process.exitCode=1;});
