const {app, BrowserWindow, dialog, ipcMain} = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const {readDashboard} = require('./excel.cjs');
let selectedFile;
const settingsPath = () => path.join(app.getPath('userData'), 'source.json');
async function restoreSource() {
  const candidates = [settingsPath(), ...['TOTAL Dashboard', 'executive-total-dashboard'].map(name => path.join(app.getPath('appData'), name, 'source.json'))];
  for (const candidate of candidates) {
    try {
      const settings = JSON.parse(await fs.readFile(candidate, 'utf8'));
      if (typeof settings.file === 'string' && path.isAbsolute(settings.file)) {
        selectedFile = settings.file;
        return;
      }
    } catch (error) {
      if (error.code !== 'ENOENT') throw new Error('Не удалось прочитать настройки источника. Выберите Excel заново.');
    }
  }
}
ipcMain.handle('excel:restore', async () => {
  await restoreSource();
  return selectedFile ? readSelected() : null;
});
async function readSelected() {
  if (!selectedFile) throw new Error('Сначала выберите Excel.');
  return {...await readDashboard(selectedFile), name: path.basename(selectedFile), updated: new Date().toISOString()};
}
ipcMain.handle('excel:open', async () => {
  const result = await dialog.showOpenDialog({properties: ['openFile'], filters: [{name: 'Excel', extensions: ['xlsx']}]});
  if (result.canceled) return null;
  const next = result.filePaths[0];
  const data = await readDashboard(next);
  await fs.mkdir(app.getPath('userData'), {recursive: true});
  await fs.writeFile(settingsPath() + '.tmp', JSON.stringify({file: next}), {mode: 0o600});
  await fs.rename(settingsPath() + '.tmp', settingsPath());
  selectedFile = next;
  return {...data, name: path.basename(next), updated: new Date().toISOString()};
});
ipcMain.handle('excel:refresh', readSelected);
function createWindow() {
  const window = new BrowserWindow({width: 1440, height: 960, minWidth: 900, minHeight: 650, title: 'Netting Dashboard', webPreferences: {preload: path.join(__dirname, 'preload.cjs'), nodeIntegration: false, contextIsolation: true, sandbox: true}});
  window.webContents.setWindowOpenHandler(() => ({action: 'deny'}));
  window.webContents.on('will-navigate', event => event.preventDefault());
  if (app.isPackaged) window.loadFile(path.join(__dirname, '../dist/index.html'));
  else window.loadURL('http://localhost:5173');
}
app.whenReady().then(() => {createWindow(); app.on('activate', () => {if (!BrowserWindow.getAllWindows().length) createWindow();});});
app.on('window-all-closed', () => {if (process.platform !== 'darwin') app.quit();});
