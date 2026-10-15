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
 const page=await browser.newPage({hasTouch:true});const errors=[];page.on('pageerror',e=>errors.push(e.message));
 let tag='version-1',downloads=0,forbidden=false,invalid=false,pages=0,sharedReads=0,missingDownload=false,remoteFull=false;
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
  if(url.pathname.endsWith('/search/query')) {assert.equal(route.request().method(),'POST');const request=route.request().postDataJSON().requests[0];assert.deepEqual(request.entityTypes,['driveItem']);assert.ok(request.query.queryString.includes('filetype:xlsx'));body={value:[{hitsContainers:[{moreResultsAvailable:false,hits:[{hitId:'book',resource:{id:'book',name:'RUB RATES LOG.xlsx',webUrl:'https://example.sharepoint.com/personal/colleague/Documents/RUB%20RATES%20LOG.xlsx',parentReference:{driveId:'colleague-drive'},file:{}}}]}]}]};}
  else if(url.pathname.endsWith('/root/children'))body={value:[{id:'folder',name:'Reports',folder:{}}]};
  else if(url.pathname.endsWith('/folder/children')) {pages++;body=url.searchParams.has('page')?{value:[{id:'book',name:'Report.xlsx',file:{}}]}:{value:[], '@odata.nextLink':'https://graph.microsoft.com/v1.0/me/drive/items/folder/children?page=2'};}
  else {const shared=url.pathname.includes('/drives/colleague-drive/');if(shared)sharedReads++;body={id:'book',name:url.pathname.includes('/drives/')?'RUB RATES LOG.xlsx':'Report.xlsx',cTag:tag,'@microsoft.graph.downloadUrl':'https://download.example.com/book.xlsx'};if(missingDownload || (shared && url.searchParams.has('$select')))delete body['@microsoft.graph.downloadUrl'];if(shared && remoteFull && !url.searchParams.has('$select')){delete body['@microsoft.graph.downloadUrl'];body.remoteItem={id:'actual',parentReference:{driveId:'origin-drive'}};}}
  await route.fulfill({contentType:'application/json',body:JSON.stringify(body)});
 });
 await page.route('https://download.example.com/**',async route=>{downloads++;assert.equal(route.request().headers().authorization,undefined);await route.fulfill({body:invalid?Buffer.from('invalid excel'):bytes});});
 const menuAction=async name=>{await page.getByRole('button',{name:'Открыть меню',exact:true}).click();await page.getByRole('menuitem',{name,exact:true}).click();};
 await page.goto('http://localhost:5173/');
 await page.getByRole('button',{name:'Войти через Microsoft',exact:true}).waitFor();
 for(const width of [390,768,1440]) {
  await page.setViewportSize({width,height:900});await page.evaluate(()=>document.fonts.ready);
  assert.equal(await page.locator('#dashboard-sidebar').count(),0,'No sidebar before authentication');
  assert.equal(await page.locator('.header-actions').count(),0,'No theme or menu before authentication');
  assert.equal(await page.getByRole('button').count(),1,'Only sign-in action is shown');
  assert.ok(await page.locator('.welcome').evaluate(e=>{const r=e.getBoundingClientRect();return Math.abs(r.x+r.width/2-innerWidth/2)<2&&Math.abs(r.y+r.height/2-innerHeight/2)<2;}),'Login card centered in viewport');
  await page.screenshot({path:`/private/tmp/netting-auth-${width}.png`});
 }
 await page.getByRole('button',{name:'Войти через Microsoft',exact:true}).click();
 await menuAction('Мой OneDrive');
 await page.getByRole('button',{name:/Reports/}).click();await page.getByRole('button',{name:/Report.xlsx/}).click();
 await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();
 assert.equal(pages,2);assert.equal(downloads,1);
 assert.ok((await page.locator('.dashboard-card.metric').first().textContent()).includes('123,00'));
 const calendar=page.getByRole('button',{name:'Выбрать дату: С даты',exact:true});
 await calendar.click();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor();await calendar.click();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});
 await calendar.click();await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).click();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});
 await calendar.click();await page.keyboard.press('Escape');await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});
 await calendar.click();await page.getByRole('button',{name:day.split('-').reverse().join('/'),exact:true}).click();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});assert.equal(await page.getByRole('textbox',{name:'С даты',exact:true}).inputValue(),day.split('-').reverse().join('/'));
 const handle=page.locator('.drag-handle').first();await handle.hover();assert.ok(await handle.evaluate(e=>{const a=e.getBoundingClientRect(),b=e.querySelector('svg').getBoundingClientRect();return Math.abs((a.left+a.width/2)-(b.left+b.width/2))<1&&Math.abs((a.top+a.height/2)-(b.top+b.height/2))<1;}),'Grip icon is centered');
 await page.getByRole('button',{name:'Тёмная тема',exact:true}).click();assert.equal(await page.getByRole('button',{name:'Светлая тема',exact:true}).count(),1);
 await page.getByRole('button',{name:'Открыть меню',exact:true}).click();await page.keyboard.press('Escape');await page.getByRole('menuitem',{name:'Выйти',exact:true}).waitFor({state:'hidden'});
 await page.evaluate(()=>document.fonts.ready);assert.equal(await page.evaluate(()=>document.fonts.check('14px "Inter Variable"','Дашборд')),true,'Local Cyrillic font loaded');
 const parsed=await page.evaluate(async base64=>{const {readDashboard}=await import('/src/excel.ts');return readDashboard(Uint8Array.from(atob(base64),c=>c.charCodeAt(0)).buffer,'Report.xlsx');},bytes.toString('base64'));
 assert.deepEqual({sheet:parsed.sheet,points:parsed.points,operations:parsed.operations},expected,'Browser parser matches desktop including cached formulas');
 await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,1,'Unchanged version is not downloaded again');
 tag='version-2';await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,2);
 forbidden=true;await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.getByRole('alert').waitFor();assert.ok((await page.getByRole('alert').textContent()).includes('Показаны данные последнего'));assert.equal(await page.locator('.dashboard-card').count(),13);forbidden=false;
 for(const width of [390,768,1440]){
 await page.setViewportSize({width,height:900});await page.waitForFunction(()=>document.documentElement.scrollWidth<=innerWidth,{},{timeout:5000}).catch(async error=>{console.log(await page.evaluate(()=>[...document.querySelectorAll('body *')].filter(e=>e.getBoundingClientRect().right>innerWidth+1).map(e=>({tag:e.tagName,cls:e.className,right:e.getBoundingClientRect().right})).slice(0,25)));await page.screenshot({path:'/private/tmp/netting-ui-overflow.png',fullPage:true});throw error;});
 const overflow=await page.evaluate(()=>[...document.querySelectorAll('body *')].filter(e=>e.getBoundingClientRect().right>innerWidth+1 && getComputedStyle(e).position!=='absolute').map(e=>[e.tagName,e.className,e.getBoundingClientRect().right]).slice(0,20)); if(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth))console.log('Overflow',overflow);
 assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false,`No page overflow at ${width}`);
 const topButtons=page.locator('.header-actions button');
 assert.equal(await topButtons.count(),2);
 const topBoxes=await topButtons.evaluateAll(nodes=>nodes.map(e=>{const r=e.getBoundingClientRect();return {x:r.x,y:r.y,w:r.width,h:r.height};}));
 assert.ok(Math.abs(topBoxes[0].y-topBoxes[1].y)<1&&Math.abs(topBoxes[0].h-topBoxes[1].h)<1,'Theme and menu share a baseline and height');
 for(const button of await page.locator('button.icon-button,button.calendar-button,button.drag-handle,button.card-rename').all()) {
  if(!await button.isVisible())continue;await button.hover();
  assert.ok(await button.evaluate(e=>{const icon=e.querySelector('svg');if(!icon)return true;const a=e.getBoundingClientRect(),b=icon.getBoundingClientRect();return Math.abs(a.x+a.width/2-b.x-b.width/2)<1&&Math.abs(a.y+a.height/2-b.y-b.height/2)<1;}),'Button icon centered on hover');
 }
 for(const button of await page.locator('button.g-button:not(.icon-button)').all()) {
  if(!await button.isVisible())continue;
  assert.ok(await button.evaluate(e=>{const text=e.querySelector('.g-button__text');if(!text)return true;const a=e.getBoundingClientRect(),b=text.getBoundingClientRect();return Math.abs(a.x+a.width/2-b.x-b.width/2)<1&&Math.abs(a.y+a.height/2-b.y-b.height/2)<1;}),'Button text wrapper centered');
 }
 const columns=await page.locator('.dashboard-grid').evaluate(e=>getComputedStyle(e).gridTemplateColumns.split(' ').length);assert.equal(columns,width===390?1:width===768?2:3);
 if(width===390){await page.getByRole('button',{name:/^Переименовать:/}).first().tap();await page.getByRole('dialog',{name:'Название карточки',exact:true}).waitFor();await page.getByRole('button',{name:'Отмена',exact:true}).tap();await page.getByRole('dialog',{name:'Название карточки',exact:true}).waitFor({state:'hidden'});await calendar.tap();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor();await calendar.tap();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});await calendar.tap();await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).tap();await page.getByRole('dialog',{name:'Календарь: С даты'}).waitFor({state:'hidden'});await menuAction('Total');assert.equal(await page.locator('.dashboard-card').count(),7);await menuAction('Обзор показателей');}
 await page.screenshot({path:`/private/tmp/netting-web-${width}.png`,fullPage:true});
 }
 assert.equal(await page.getByRole('button',{name:/^Переименовать:/}).count(),13,'Every card can be renamed');
 const metricTitle='TOTAL · за период, USD',customMetric='Прибыль за период';
 const chartTitle='Динамика TOTAL',customChart='Результат по дням';
 await page.getByRole('button',{name:`Переименовать: ${metricTitle}`,exact:true}).click();
 const renameDialog=page.getByRole('dialog',{name:'Название карточки',exact:true});await renameDialog.waitFor();
 const nameInput=renameDialog.getByRole('textbox',{name:'Название карточки',exact:true});
 await nameInput.fill('   ');await renameDialog.getByRole('button',{name:'Сохранить',exact:true}).click();await renameDialog.getByText('Введите название карточки.',{exact:true}).waitFor();
 await nameInput.fill(customMetric);await nameInput.press('Enter');await renameDialog.waitFor({state:'hidden'});
 await page.locator('.sidebar-nav').getByRole('button',{name:'Total',exact:true}).click();assert.equal(await page.getByRole('button',{name:`Переименовать: ${customMetric}`,exact:true}).count(),1,'Name is shared across sections');
 await page.locator('.sidebar-nav').getByRole('button',{name:'Обзор показателей',exact:true}).click();
 await page.getByRole('button',{name:`Переименовать: ${chartTitle}`,exact:true}).click();await renameDialog.getByRole('textbox',{name:'Название карточки',exact:true}).fill(customChart);await renameDialog.getByRole('button',{name:'Сохранить',exact:true}).click();await renameDialog.waitFor({state:'hidden'});
 assert.equal(await page.getByRole('heading',{name:customChart,exact:true}).count(),1,'Panel heading uses custom title');
 const saved=await page.evaluate(()=>Object.fromEntries(Object.entries(localStorage).filter(([k])=>k.startsWith('netting-source:'))));assert.deepEqual(JSON.parse(Object.values(saved)[0]),{id:'book',name:'Report.xlsx'});
 await page.reload();await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();assert.equal(downloads,3,'Reload restores source and reloads data into memory');
 assert.equal(await page.getByRole('heading',{name:customChart,exact:true}).count(),1,'Panel name survives reload');
 assert.equal(await page.getByRole('button',{name:`Как рассчитывается: ${customMetric}`,exact:true}).count(),1,'Tooltip remains available under renamed title');
 await page.getByRole('button',{name:`Переименовать: ${customMetric}`,exact:true}).click();await renameDialog.getByRole('textbox',{name:'Название карточки',exact:true}).fill('Не сохранять');await page.keyboard.press('Escape');await renameDialog.waitFor({state:'hidden'});assert.equal(await page.getByRole('button',{name:`Переименовать: ${customMetric}`,exact:true}).count(),1,'Cancel leaves previous name intact');
 for(const title of [customMetric,customChart]){await page.getByRole('button',{name:`Переименовать: ${title}`,exact:true}).click();await renameDialog.getByRole('button',{name:'Вернуть исходное название',exact:true}).click();await renameDialog.waitFor({state:'hidden'});}
 assert.equal(await page.getByRole('heading',{name:chartTitle,exact:true}).count(),1,'Original panel title restored');
 await menuAction('Общие файлы');
 await page.getByRole('button',{name:/OneDrive · colleague/}).click();await page.getByRole('button',{name:'Документы',exact:true}).click();await page.getByRole('button',{name:'Назад',exact:true}).click();await page.getByRole('button',{name:'Документы',exact:true}).click();
 await page.getByRole('button',{name:/RUB RATES LOG.xlsx/}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(sharedReads,2,'Missing selected annotation triggers full metadata retry');assert.equal(downloads,4,'Same item id in another drive must not reuse own-drive cache');
 const sharedSaved=await page.evaluate(()=>JSON.parse(localStorage.getItem('netting-source:test-user')));assert.deepEqual(sharedSaved,{id:'book',name:'RUB RATES LOG.xlsx',driveId:'colleague-drive'});
 await page.reload();await page.getByRole('heading',{name:'Данные по дням',exact:true}).waitFor();assert.equal(downloads,5);assert.equal(sharedReads,4,'Restored shared source retries full metadata in colleague drive');
 invalid=true;tag='version-3';await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.getByRole('alert').waitFor();assert.equal(await page.locator('.dashboard-card').count(),13);
 missingDownload=true;invalid=false;tag='version-4';await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.getByText(/не предоставил адрес её скачивания/).waitFor();assert.equal(await page.getByRole('button',{name:'Войти снова',exact:true}).count(),0,'Missing download URL is not an authentication failure');
 missingDownload=false;remoteFull=true;await page.getByRole('button',{name:'Повторить загрузку',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(await page.getByRole('alert').count(),0,'Retry resolves remoteItem against origin drive');
 await menuAction('Выйти');await page.getByRole('dialog',{name:'Выйти из дашборда?'}).waitFor();await page.getByRole('button',{name:'Остаться',exact:true}).click();assert.equal(await page.locator('.dashboard-card').count(),13,'Cancel logout preserves data');
 await menuAction('Выйти');await page.getByRole('dialog',{name:'Выйти из дашборда?'}).waitFor();await page.keyboard.press('Escape');await page.getByRole('dialog',{name:'Выйти из дашборда?'}).waitFor({state:'hidden'});assert.equal(await page.locator('.dashboard-card').count(),13);
 await menuAction('Выйти');await page.getByRole('button',{name:'Выйти',exact:true}).click();await page.getByRole('button',{name:'Войти через Microsoft',exact:true}).waitFor();assert.equal(await page.locator('.dashboard-card').count(),0);
 assert.deepEqual(errors,[]);console.log('Web checks passed: card rename/validation/cancellation/reset/persistence across sections and reload, centered sign-in-only screen, aligned header and all icon/text buttons, mouse/touch calendar dismissal, menu, logout confirmation and cancellation, folder catalog, local Cyrillic font, centered grip, 390/768/1440 layouts, Excel parsing, loading and refresh. Microsoft authentication is mocked.');
 } finally {await browser.close();await fs.unlink(fixture);}
})().catch(e=>{console.error(e);process.exitCode=1;});
