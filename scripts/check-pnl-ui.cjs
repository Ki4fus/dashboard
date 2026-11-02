const {_electron:electron}=require('playwright');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
(async()=>{
 const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Dubai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 const previous=new Date(new Date(today+'T00:00:00Z').getTime()-86400000).toISOString().slice(0,10);
 const point=(date,total,fxPnl,row)=>({date,total,fxPnl,row,fixedUsdAmount:50,issue:null,fxIssue:null,fixedIssue:null});
 const data={name:'Pnl fixture.xlsx',sheet:'RUB Deals Rates',updated:new Date().toISOString(),operations:[],points:[point(previous,999,888,3),point(today,120.25,30.5,5)]};
 const profile=await fs.mkdtemp('/private/tmp/netting-pnl-ui-');const env={...process.env,NETTING_TEST_PROFILE:profile};delete env.ELECTRON_RUN_AS_NODE;
 const app=await electron.launch({args:[path.resolve('scripts/ui-fixture.cjs')],env});
 try{
 await app.evaluate(({BrowserWindow,ipcMain},{root,data})=>{for(const action of ['restore','open','refresh'])ipcMain.handle('excel:'+action,()=>data);const win=new BrowserWindow({width:1440,height:1000,webPreferences:{preload:root+'/electron/preload.cjs',contextIsolation:true,sandbox:true}});win.loadFile(root+'/dist/index.html');},{root:process.cwd(),data});
 const page=await app.firstWindow();const errors=[];page.on('pageerror',e=>errors.push(e.message));await page.locator('.dashboard-card[aria-busy="false"]').first().waitFor();
 const card=title=>page.locator('.dashboard-card').filter({has:page.getByRole('button',{name:'How it is calculated: '+title,exact:true})});
 assert.equal(await card('Our P&L · Today (USD)').locator('strong').textContent(),'120.25');
 assert.equal(await card('FX P&L · Today (USD)').locator('strong').textContent(),'30.50');
 assert.equal(await card('P&L+FX· Today (USD)').locator('strong').textContent(),'150.75');
 assert.equal(await card('Fixed Amt · Period (USD)').count(),1);assert.equal(await card('Fixed Amt · Today (USD)').count(),1);
 await page.getByRole('button',{name:'How it is calculated: FX P&L · Today (USD)',exact:true}).hover();await page.getByRole('tooltip').waitFor();assert.ok((await page.getByRole('tooltip').textContent()).includes('T'));await page.mouse.move(1,1);
 await page.getByRole('textbox',{name:'From',exact:true}).fill(previous.split('-').reverse().join('/'));await page.getByRole('textbox',{name:'To',exact:true}).fill(previous.split('-').reverse().join('/'));
 assert.equal(await card('Our P&L · Today (USD)').locator('strong').textContent(),'120.25');assert.equal(await card('P&L+FX· Today (USD)').locator('strong').textContent(),'150.75');
 assert.equal(await page.locator('aside,#dashboard-sidebar').count(),0,'Sidebar is removed');
 for(const [our,fx,expected] of [[120.25,null,'120.25'],[null,30.5,'30.50'],[null,null,'0.00']]) {
  const next={...data,points:[point(previous,999,888,3),point(today,our,fx,5)]};
  await app.evaluate(({ipcMain},next)=>{ipcMain.removeHandler('excel:refresh');ipcMain.handle('excel:refresh',()=>next);},next);
  await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));
  assert.equal(await card('P&L+FX· Today (USD)').locator('strong').textContent(),expected,'Each dash counts as zero in combined P&L');
 }
 await app.evaluate(({ipcMain},next)=>{ipcMain.removeHandler('excel:refresh');ipcMain.handle('excel:refresh',()=>next);},{...data,points:[point(previous,999,888,3)]});
 await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));assert.equal(await card('Our P&L · Today (USD)').locator('strong').textContent(),'—');assert.equal(await card('P&L+FX· Today (USD)').locator('strong').textContent(),'0.00');
 await app.evaluate(({ipcMain},next)=>{ipcMain.removeHandler('excel:refresh');ipcMain.handle('excel:refresh',()=>next);},data);await page.getByRole('button',{name:'Refresh',exact:true}).click();await page.waitForFunction(()=>!document.querySelector('[aria-busy="true"]'));
 assert.equal(await page.locator('body').evaluate(e=>/[А-Яа-яЁё]/.test(e.innerText)),false,'Rendered dashboard is in English');
 await page.getByRole('button',{name:'Open menu',exact:true}).click();await page.getByRole('menuitem',{name:'Our P&L',exact:true}).click();assert.equal(await card('FX P&L · Today (USD)').count(),1);assert.equal(await card('P&L+FX· Today (USD)').count(),1);
 await page.getByRole('button',{name:'Open menu',exact:true}).click();await page.getByRole('menuitem',{name:'Fixed Amount',exact:true}).click();assert.equal(await card('FX P&L · Today (USD)').count(),0);assert.equal(await card('Fixed Amt · Today (USD)').count(),1);
 await page.evaluate(()=>localStorage.setItem('netting.dashboard.card-titles.v1',JSON.stringify({'total-latest':'Our P&L · Latest Day (USD)','fixed-period':'Fixed USD · Period','fixed-today':'Fixed USD · Today','total-chart':'Динамика TOTAL','currencies':'Разбивка по валютам','rate-today':'My own rate'})));await page.reload();await card('Our P&L · Today (USD)').waitFor();assert.equal(await card('Fixed Amt · Today (USD)').count(),1);assert.equal(await page.getByRole('heading',{name:'Trend of our P&L · Today (USD)',exact:true}).count(),1);assert.equal(await page.getByRole('heading',{name:'Currency breakdown',exact:true}).count(),1);assert.equal(await card('My own rate').count(),1,'Custom card names remain unchanged');assert.deepEqual(errors,[]);
 console.log('P&L UI passed: requested names, T and Y+T, actual today independent of filter, tooltip, menus and saved-title migration.');
 }finally{await app.close();}
})().catch(e=>{console.error(e);process.exitCode=1});
