const {chromium} = require('playwright');
const ExcelJS = require('exceljs');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const {readDashboard} = require('../electron/excel.cjs');
(async()=>{
 const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('RUB Deals Rates');
 const day=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 sheet.getCell('A1').value=new Date(day+'T00:00:00Z');sheet.getCell('E1').value=80;sheet.getCell('L1').value='USD';sheet.getCell('U1').value=200;
 sheet.getCell('A2').value='TOTAL';sheet.getCell('Y2').value={formula:'100+23',result:123};sheet.getCell('T2').value=50;sheet.getCell('U2').value={formula:'100+100',result:200};
 sheet.getCell('A3').value='TOTAL'; // An ambiguous block exercises removal of the old warning summary.
 sheet.spliceRows(2,0,[],[]);
 for(const [row,currency,value] of [[2,'EUR',100],[3,'GBP',50]]){sheet.getCell(`A${row}`).value=new Date(day+'T00:00:00Z');sheet.getCell(`L${row}`).value=currency;sheet.getCell(`U${row}`).value=value;}
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
 const menuAction=async name=>{await page.getByRole('button',{name:'Open menu',exact:true}).click();await page.getByRole('menuitem',{name,exact:true}).click();};
 await page.addInitScript(()=>localStorage.setItem('netting.dashboard.card-order.v1',JSON.stringify(['rate-change','fixed-today','total-period','total-latest'])));
 await page.goto('http://localhost:5173/');
 await page.getByRole('button',{name:'Sign in with Microsoft',exact:true}).waitFor();
 for(const width of [390,768,1440]) {
  await page.setViewportSize({width,height:900});await page.evaluate(()=>document.fonts.ready);
  assert.equal(await page.locator('#dashboard-sidebar').count(),0,'No sidebar before authentication');
  assert.equal(await page.locator('.header-actions').count(),0,'No theme or menu before authentication');
  assert.equal(await page.getByRole('button').count(),1,'Only sign-in action is shown');
  assert.ok(await page.locator('.welcome').evaluate(e=>{const r=e.getBoundingClientRect();return Math.abs(r.x+r.width/2-innerWidth/2)<2&&Math.abs(r.y+r.height/2-innerHeight/2)<2;}),'Login card centered in viewport');
  await page.screenshot({path:`/private/tmp/netting-auth-${width}.png`});
 }
 await page.getByRole('button',{name:'Sign in with Microsoft',exact:true}).click();
 await menuAction('My OneDrive');
 await page.getByRole('button',{name:/Reports/}).click();await page.getByRole('button',{name:/Report.xlsx/}).click();
 await page.getByRole('heading',{name:'Daily data',exact:true}).waitFor();
 assert.equal(pages,2);assert.equal(downloads,1);assert.equal(await page.locator('aside,#dashboard-sidebar').count(),0,'No sidebar after sign-in');assert.equal(await page.locator('body').evaluate(e=>/[А-Яа-яЁё]/.test(e.innerText)),false,'Dashboard labels are in English');
 assert.ok((await page.locator('.dashboard-card.metric').first().textContent()).includes('173.00'));
 assert.equal(await page.locator('[data-card-id="total-fx-period"] strong').textContent(),'50.00');
 assert.equal(await page.locator('[data-card-id="total-combined-period"] strong').textContent(),'173.00');
 assert.equal(await page.locator('[data-card-id="total-period"] strong').textContent(),'123.00','P&L period retains column Y');
 assert.equal(await page.locator('.donut-segment').count(),3);
 const eur=page.locator('button.currency-entry').filter({has:page.locator('.currency-name',{hasText:'EUR'})});
 await eur.click();assert.equal(await eur.getAttribute('aria-pressed'),'true');assert.equal(await page.locator('.donut-center>span').textContent(),'EUR');assert.equal(await page.locator('.donut-segment.is-muted').count(),2);await eur.click();
 assert.equal(await page.locator('.donut-center strong').textContent(),'350');
 await page.locator('.metric-groups').screenshot({path:'/private/tmp/netting-depth-light.png'});
 await page.locator('[data-card-id="currencies"]').screenshot({path:'/private/tmp/netting-donut-light.png'});
 const expectedGroups=[['total-combined-today','total-latest','total-fx-today'],['fixed-today'],['total-period','fixed-period','total-fx-period','total-combined-period'],['total-average','total-change'],['rate-today','rate-period','rate-change']];
 const groupCards=()=>page.locator('.metric-group').evaluateAll(groups=>groups.map(group=>[...group.querySelectorAll('.dashboard-card')].map(card=>card.dataset.cardId)));
 assert.deepEqual(await groupCards(),expectedGroups,'Five metric blocks follow the requested hierarchy');
 assert.equal(await page.locator('.notice').count(),0,'No filename/Loaded-at/error-count summary, including with invalid totals');
 const calendar=page.getByRole('button',{name:'Choose date: From',exact:true});
 await calendar.click();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor();await calendar.click();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});
 await calendar.click();await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).click();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});
 await calendar.click();await page.keyboard.press('Escape');await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});
 await calendar.click();await page.getByRole('button',{name:day.split('-').reverse().join('/'),exact:true}).click();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});assert.equal(await page.getByRole('textbox',{name:'From',exact:true}).inputValue(),day.split('-').reverse().join('/'));
 const handle=page.locator('.drag-handle').first();await handle.hover();assert.ok(await handle.evaluate(e=>{const a=e.getBoundingClientRect(),b=e.querySelector('svg').getBoundingClientRect();return Math.abs((a.left+a.width/2)-(b.left+b.width/2))<1&&Math.abs((a.top+a.height/2)-(b.top+b.height/2))<1;}),'Grip icon is centered');
 await page.getByRole('button',{name:'Dark theme',exact:true}).click();assert.equal(await page.getByRole('button',{name:'Light theme',exact:true}).count(),1);await page.locator('.metric-groups').screenshot({path:'/private/tmp/netting-depth-dark.png'});await page.locator('[data-card-id="currencies"]').screenshot({path:'/private/tmp/netting-donut-dark.png'});
 await page.getByRole('button',{name:'Open menu',exact:true}).click();await page.keyboard.press('Escape');await page.getByRole('menuitem',{name:'Sign out',exact:true}).waitFor({state:'hidden'});
 await page.evaluate(()=>document.fonts.ready);assert.equal(await page.evaluate(()=>document.fonts.check('14px "Inter Variable"','Dashboard')),true,'Local English font loaded');
 const parsed=await page.evaluate(async base64=>{const {readDashboard}=await import('/src/excel.ts');return readDashboard(Uint8Array.from(atob(base64),c=>c.charCodeAt(0)).buffer,'Report.xlsx');},bytes.toString('base64'));
 assert.deepEqual({sheet:parsed.sheet,points:parsed.points,operations:parsed.operations},expected,'Browser parser matches desktop including cached formulas');
 await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,1,'Unchanged version is not downloaded again');
 tag='version-2';await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(downloads,2);
 forbidden=true;await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.getByRole('alert').waitFor();assert.ok((await page.getByRole('alert').textContent()).includes('Showing data from the last successful'));assert.equal(await page.locator('.dashboard-card').count(),17);forbidden=false;await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));
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
 const columns=await page.locator('[data-group-id="rates"] .group-grid').evaluate(e=>getComputedStyle(e).gridTemplateColumns.split(' ').length);assert.equal(columns,width===390?1:width===768?2:3);assert.deepEqual(await groupCards(),expectedGroups);
 const periodColumns=await page.locator('[data-group-id="period"] .group-grid').evaluate(e=>getComputedStyle(e).gridTemplateColumns.split(' ').length);assert.equal(periodColumns,width===390?1:2);
 if(width===390){await page.getByRole('button',{name:/^Rename:/}).first().tap();await page.getByRole('dialog',{name:'Card name',exact:true}).waitFor();await page.getByRole('button',{name:'Cancel',exact:true}).tap();await page.getByRole('dialog',{name:'Card name',exact:true}).waitFor({state:'hidden'});await calendar.tap();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor();await calendar.tap();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});await calendar.tap();await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).tap();await page.getByRole('dialog',{name:'Calendar: From'}).waitFor({state:'hidden'});await menuAction('P&L');assert.equal(await page.locator('.dashboard-card').count(),11);await menuAction('Overview');}
 await page.screenshot({path:`/private/tmp/netting-web-${width}.png`,fullPage:true});
 }
 assert.equal(await page.getByRole('button',{name:/^Rename:/}).count(),17,'Every card can be renamed');
 const metricTitle='P&L · Period (USD)',customMetric='Period profit';
 const chartTitle='Trend of P&L · Today (USD)',customChart='Daily results';
 await page.getByRole('button',{name:`Rename: ${metricTitle}`,exact:true}).click();
 const renameDialog=page.getByRole('dialog',{name:'Card name',exact:true});await renameDialog.waitFor();
 const nameInput=renameDialog.getByRole('textbox',{name:'Card name',exact:true});
 await nameInput.fill('   ');await renameDialog.getByRole('button',{name:'Save',exact:true}).click();await renameDialog.getByText('Enter a card name.',{exact:true}).waitFor();
 await nameInput.fill(customMetric);await nameInput.press('Enter');await renameDialog.waitFor({state:'hidden'});
 await menuAction('P&L');assert.equal(await page.getByRole('button',{name:`Rename: ${customMetric}`,exact:true}).count(),1,'Name is shared across sections');
 await menuAction('Overview');
 await page.getByRole('button',{name:`Rename: ${chartTitle}`,exact:true}).click();await renameDialog.getByRole('textbox',{name:'Card name',exact:true}).fill(customChart);await renameDialog.getByRole('button',{name:'Save',exact:true}).click();await renameDialog.waitFor({state:'hidden'});
 assert.equal(await page.getByRole('heading',{name:customChart,exact:true}).count(),1,'Panel heading uses custom title');
 const saved=await page.evaluate(()=>Object.fromEntries(Object.entries(localStorage).filter(([k])=>k.startsWith('netting-source:'))));assert.deepEqual(JSON.parse(Object.values(saved)[0]),{id:'book',name:'Report.xlsx'});
 await page.reload();await page.getByRole('heading',{name:'Daily data',exact:true}).waitFor();assert.equal(downloads,3,'Reload restores source and reloads data into memory');
 assert.equal(await page.getByRole('heading',{name:customChart,exact:true}).count(),1,'Panel name survives reload');
 assert.equal(await page.getByRole('button',{name:`How it is calculated: ${customMetric}`,exact:true}).count(),1,'Tooltip remains available under renamed title');
 await page.getByRole('button',{name:`Rename: ${customMetric}`,exact:true}).click();await renameDialog.getByRole('textbox',{name:'Card name',exact:true}).fill('Do not save');await page.keyboard.press('Escape');await renameDialog.waitFor({state:'hidden'});assert.equal(await page.getByRole('button',{name:`Rename: ${customMetric}`,exact:true}).count(),1,'Cancel leaves previous name intact');
 for(const title of [customMetric,customChart]){await page.getByRole('button',{name:`Rename: ${title}`,exact:true}).click();await renameDialog.getByRole('button',{name:'Restore default name',exact:true}).click();await renameDialog.waitFor({state:'hidden'});}
 assert.equal(await page.getByRole('heading',{name:chartTitle,exact:true}).count(),1,'Original panel title restored');
 await menuAction('Shared files');
 await page.getByRole('button',{name:/OneDrive · colleague/}).click();await page.getByRole('button',{name:'Documents',exact:true}).click();await page.getByRole('button',{name:'Back',exact:true}).click();await page.getByRole('button',{name:'Documents',exact:true}).click();
 await page.getByRole('button',{name:/RUB RATES LOG.xlsx/}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(sharedReads,2,'Missing selected annotation triggers full metadata retry');assert.equal(downloads,4,'Same item id in another drive must not reuse own-drive cache');
 const sharedSaved=await page.evaluate(()=>JSON.parse(localStorage.getItem('netting-source:test-user')));assert.deepEqual(sharedSaved,{id:'book',name:'RUB RATES LOG.xlsx',driveId:'colleague-drive'});
 await page.reload();await page.getByRole('heading',{name:'Daily data',exact:true}).waitFor();assert.equal(downloads,5);assert.equal(sharedReads,4,'Restored shared source retries full metadata in colleague drive');
 const rateGroup=page.locator('[data-group-id="rates"]');
 const beforeRates=await rateGroup.locator('.dashboard-card').evaluateAll(nodes=>nodes.map(e=>e.dataset.cardId));
 const beforeGroups=await groupCards();
 await rateGroup.locator('.drag-handle').first().focus();await page.keyboard.press('Space');await page.keyboard.press('ArrowRight');await page.keyboard.press('Space');
 const movedRates=await rateGroup.locator('.dashboard-card').evaluateAll(nodes=>nodes.map(e=>e.dataset.cardId));assert.notDeepEqual(movedRates,beforeRates,'Cards can reorder inside a block');assert.deepEqual((await groupCards()).slice(0,4),beforeGroups.slice(0,4),'Other blocks keep their order');
 await page.reload();await page.getByRole('heading',{name:'Daily data',exact:true}).waitFor();assert.deepEqual(await rateGroup.locator('.dashboard-card').evaluateAll(nodes=>nodes.map(e=>e.dataset.cardId)),movedRates,'Block layout survives reload');
 invalid=true;tag='version-3';await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.getByRole('alert').waitFor();assert.equal(await page.locator('.dashboard-card').count(),17);
 missingDownload=true;invalid=false;tag='version-4';await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.getByText(/without a download URL/).waitFor();assert.equal(await page.getByRole('button',{name:'Sign in again',exact:true}).count(),0,'Missing download URL is not an authentication failure');
 missingDownload=false;remoteFull=true;await page.getByRole('button',{name:'Retry download',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(await page.getByRole('alert').count(),0,'Retry resolves remoteItem against origin drive');
 await menuAction('Sign out');await page.getByRole('dialog',{name:'Sign out of the dashboard?'}).waitFor();await page.getByRole('button',{name:'Stay signed in',exact:true}).click();assert.equal(await page.locator('.dashboard-card').count(),17,'Cancel logout preserves data');
 await menuAction('Sign out');await page.getByRole('dialog',{name:'Sign out of the dashboard?'}).waitFor();await page.keyboard.press('Escape');await page.getByRole('dialog',{name:'Sign out of the dashboard?'}).waitFor({state:'hidden'});assert.equal(await page.locator('.dashboard-card').count(),17);
 await menuAction('Sign out');await page.getByRole('button',{name:'Sign out',exact:true}).click();await page.getByRole('button',{name:'Sign in with Microsoft',exact:true}).waitFor();assert.equal(await page.locator('.dashboard-card').count(),0);
 assert.deepEqual(errors,[]);console.log('Web checks passed: five ordered metric blocks, within-block keyboard reorder and persistence, removed summary banner, card rename/validation/cancellation/reset/persistence across sections and reload, centered sign-in-only screen, aligned header and all icon/text buttons, mouse/touch calendar dismissal, menu, logout confirmation and cancellation, folder catalog, local font, centered grip, 390/768/1440 layouts, Excel parsing, loading and refresh. Microsoft authentication is mocked.');
 } finally {await browser.close();await fs.unlink(fixture);}
})().catch(e=>{console.error(e);process.exitCode=1;});
