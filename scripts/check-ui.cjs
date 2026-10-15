const {_electron: electron} = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const {readDashboard} = require('../electron/excel.cjs');
(async()=>{
 const profile=await fs.mkdtemp('/private/tmp/netting-ui-');
 const data={...await readDashboard('/Users/sufi/Downloads/RUB RATES LOG-2.xlsx'),name:'Example.xlsx',updated:new Date().toISOString()};
 const launchEnv={...process.env,NETTING_TEST_PROFILE:profile};delete launchEnv.ELECTRON_RUN_AS_NODE;
 const app=await electron.launch({args:[path.resolve('scripts/ui-fixture.cjs')],env:launchEnv});
 try{
 await app.evaluate(({BrowserWindow,ipcMain},{root,data})=>{
  ipcMain.handle('excel:restore',()=>data);ipcMain.handle('excel:refresh',()=>data);ipcMain.handle('excel:open',()=>data);
  const win=new BrowserWindow({width:1440,height:1000,webPreferences:{preload:root+'/electron/preload.cjs',contextIsolation:true,sandbox:true}});
  win.loadFile(root+'/dist/index.html');
 },{root:process.cwd(),data});
 const page=await app.firstWindow();const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).waitFor();
 const titles=()=>page.locator('.dashboard-grid .drag-handle').evaluateAll(nodes=>nodes.map(n=>n.getAttribute('aria-label')));
 await page.locator('.dashboard-card[aria-busy="false"]').first().waitFor();
 const initial=await titles();assert.equal(initial.length,13);
 assert.equal(await page.getByRole('button',{name:/Как рассчитывается:/}).count(),13);
 await page.getByRole('button',{name:'Как рассчитывается: TOTAL · за период, USD',exact:true}).hover();await page.getByRole('tooltip').waitFor();assert.ok((await page.getByRole('tooltip').textContent()).includes('Y'));await page.mouse.move(1,1);await page.getByRole('tooltip').waitFor({state:'hidden'});
 const fromInput=page.getByRole('textbox',{name:'С даты',exact:true}),toInput=page.getByRole('textbox',{name:'По дату',exact:true});assert.equal(await toInput.inputValue(),'13/10/2026');await fromInput.fill('12/10/2026');await fromInput.blur();assert.equal(await page.locator('.table-scroll tbody tr').count(),2);await fromInput.fill('31/02/2026');await fromInput.blur();await page.getByText('Введите существующую дату: дд/мм/гггг',{exact:true}).waitFor();assert.equal(await page.locator('.table-scroll tbody tr').count(),2);await fromInput.fill('');await toInput.fill('13/10/2026');await toInput.blur();

 assert.equal(await page.locator('.powered-by').textContent(),'Powered by LSD');
 const handles=page.locator('.dashboard-grid .drag-handle');
 const a=await handles.nth(0).boundingBox(), b=await page.locator(".dashboard-grid .dashboard-card").nth(4).boundingBox();
 await page.mouse.move(a.x+a.width/2,a.y+a.height/2);await page.mouse.down();await page.mouse.move(a.x+15,a.y+15,{steps:5});await page.mouse.move(b.x+b.width/2,b.y+b.height/2,{steps:20});
 await page.waitForTimeout(100);console.log('drag state',await page.locator('.is-dragging').count(),await titles());await page.screenshot({path:'/private/tmp/netting-drag.png'});assert.notDeepEqual(await titles(),initial,'Cards must move while pointer is down');await page.mouse.up();
 const changed=await titles();await page.reload();await page.getByRole('heading',{name:'Netting Dashboard',exact:true}).waitFor();assert.deepEqual(await titles(),changed,'Order survives reload');
 const current=await titles();await handles.nth(0).focus();await page.keyboard.press('Space');await page.keyboard.press('ArrowRight');await page.keyboard.press('Escape');assert.deepEqual(await titles(),current,'Escape cancels sorting');
 const credit=await page.locator('.powered-by').textContent();assert.equal(credit,'Powered by LSD');
 for(const [value,css] of [[90,'change-negative'],[100,'change-neutral'],[110,'change-positive']]){
  const sample={...data,points:[{date:'2026-10-12',row:1,total:100,fixedUsdAmount:100,issue:null,fixedIssue:null},{date:'2026-10-13',row:2,total:value,fixedUsdAmount:value,issue:null,fixedIssue:null}],operations:[{date:'2026-10-12',row:1,currency:'USD',rate:100,fixedUsdAmount:100,hasAmount:true},{date:'2026-10-13',row:2,currency:'USD',rate:value,fixedUsdAmount:value,hasAmount:true}]};
  await app.evaluate(({ipcMain},sample)=>{ipcMain.removeHandler('excel:refresh');ipcMain.handle('excel:refresh',async()=>{await new Promise(r=>setTimeout(r,800));return sample;});},sample);
  const before=await page.locator('.dashboard-card').first().boundingBox();await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.locator('.card-skeleton').first().waitFor();assert.equal(await page.locator('.card-skeleton').count(),13);const during=await page.locator('.dashboard-card').first().boundingBox();assert.equal(before.height,during.height);
  await page.locator('.card-skeleton').first().waitFor({state:'detached'});assert.equal(await page.locator('strong.'+css).count(),2);
 }
 await app.evaluate(({ipcMain},data)=>{ipcMain.removeHandler('excel:refresh');ipcMain.handle('excel:refresh',()=>data);},data);await page.getByRole('button',{name:'Обновить',exact:true}).click();await page.locator('.card-skeleton').first().waitFor({state:'detached'});
 for(const width of [1440,900]){await app.evaluate(({BrowserWindow},width)=>BrowserWindow.getAllWindows()[0].setSize(width,1000),width);await page.waitForTimeout(150);const overflow=await page.locator('.currency-layout').evaluate(root=>{const parent=root.closest('.dashboard-card').getBoundingClientRect();return [...root.querySelectorAll('.currency-entry,.currency-chart')].some(e=>{const r=e.getBoundingClientRect();return r.left<parent.left||r.right>parent.right;});});assert.equal(overflow,false,'Currency values stay inside card');}
 await page.getByRole('button',{name:'Тёмная тема',exact:true}).click();
 const overviewOrder=await titles();await page.getByRole('button',{name:'Total',exact:true}).click();assert.equal(await handles.count(),7);assert.equal(await page.getByRole('heading',{name:'Динамика TOTAL',exact:true}).count(),1);assert.equal(await page.getByRole('heading',{name:'Разбивка по валютам',exact:true}).count(),0);assert.equal(await page.getByRole('heading',{name:'Данные по дням',exact:true}).count(),1);assert.equal(await page.getByRole('heading',{name:'Проверьте даты в Excel',exact:true}).count(),1);
 await page.getByRole('button',{name:'Fixed USD Amount',exact:true}).click();assert.equal(await handles.count(),5);assert.equal(await page.getByRole('heading',{name:'Разбивка по валютам',exact:true}).count(),1);assert.equal(await page.getByRole('heading',{name:'Динамика TOTAL',exact:true}).count(),0);assert.equal(await page.getByRole('heading',{name:'Данные по дням',exact:true}).count(),1);assert.equal(await page.getByRole('heading',{name:'Проверьте даты в Excel',exact:true}).count(),1);
 await page.getByRole('button',{name:'Обзор показателей',exact:true}).click();assert.deepEqual(await titles(),overviewOrder,'Overview order survives switching sections');
 assert.deepEqual(errors,[]);await page.screenshot({path:'/private/tmp/netting-dashboard-preview.png',fullPage:true});console.log('UI checks passed: live grid reorder, saved order, keyboard cancellation, charts, credit, skeleton sizing, change colors, responsive currency layout, navigation filters, overview order retention, Gravity tooltips, dd/mm/yyyy filters and invalid dates.');
 }finally{await app.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
