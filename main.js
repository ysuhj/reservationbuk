// 북가좌초 특별실 PC 앱 (Electron)
// GitHub Pages에 올린 웹앱을 창으로 띄웁니다. 웹앱을 고치면 PC 앱에도 바로 반영됩니다.
const { app, BrowserWindow, shell, Menu } = require('electron');
const path = require('path');

// GitHub Pages 주소로 바꾸세요
const APP_URL = 'https://깃허브사용자명.github.io/specialroom/';

let win = null;

function createWindow() {
  win = new BrowserWindow({
    width: 1320,
    height: 900,
    minWidth: 400,
    minHeight: 500,
    title: '북가좌초 특별실',
    backgroundColor: '#123a2f',
    autoHideMenuBar: true,
    icon: path.join(__dirname, 'icon.png'),
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true }
  });

  win.loadURL(APP_URL);

  // 시트 열기 등 바깥 링크는 기본 브라우저로
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith(APP_URL)) { e.preventDefault(); shell.openExternal(url); }
  });

  // 인터넷이 끊겼을 때 안내 화면
  win.webContents.on('did-fail-load', (e, code, desc, url, isMain) => {
    if (isMain && code !== -3) win.loadFile(path.join(__dirname, 'offline.html'), { query: { u: APP_URL } });
  });
}

// 두 번 실행하면 기존 창을 앞으로
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (win) { if (win.isMinimized()) win.restore(); win.focus(); }
  });
  app.whenReady().then(() => {
    Menu.setApplicationMenu(Menu.buildFromTemplate([
      { label: '보기', submenu: [
        { label: '새로고침', accelerator: 'F5', click: () => win && win.loadURL(APP_URL) },
        { role: 'zoomIn', label: '크게' }, { role: 'zoomOut', label: '작게' }, { role: 'resetZoom', label: '원래 크기' },
        { type: 'separator' }, { role: 'togglefullscreen', label: '전체 화면' }
      ] }
    ]));
    createWindow();
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
