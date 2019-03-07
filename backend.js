const { app } = require('electron');
const frontendDir = __dirname + '/frontend/';
const iconFile = __dirname + '/icon.png';

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	let debugMode = false;

	function loadMainWin() {
		mainWindow = createWin(700, 360, false);
		mainWindow.loadURL(frontendDir + 'main.html');
		mainWindow.focus();

		mainWindow.on('closed', function(e) {
			if (mainWindow2 === null) {
				mainWindow = null;
				if (process.platform != 'darwin') {
					app.quit();
				}
			} else {
				e.preventDefault();
			}
		});
	}

	function createWin(width, height, maximizable, winFile) {
		const { BrowserWindow } = require('electron');
		let retData;
		retData = new BrowserWindow({
			width: width,
			height: height,
			minWidth: width,
			minHeight: height,
			title: "",
			icon: iconFile,
			frame: false,
			resize: (maximizable ? true : false),
			maximizable: maximizable,
			backgroundColor: '#060621',
			webPreferences: { webSecurity: false, nodeIntegration: true }
		});
		if (!maximizable) {
			retData.setMaximumSize(width, height);
		}
		if (debugMode) {
			retData.webContents.openDevTools();
		}
		return retData;
	}

	function loadIpc() {
		const ipc = require('electron').ipcMain;
		ipc.on('remote', (event, data) => {
			switch (data) {
				case "exit":
					if (mainWindow2 != null) {
						mainWindow2.destroy();
					}
					if (process.platform != 'darwin') {
						app.quit();
					}
					break;
				case "select1":
					mainWindow2 = createWin(300, 300, false);
					mainWindow2.loadURL(frontendDir + 'control.html');
					break;
				case "select2":
					mainWindow2 = createWin(1200, 680, true);
					mainWindow2.loadURL(frontendDir + 'index.html');
					break;
				default:
					mainWindow.destroy();
					break;
			}
			if (/^select/.test(data)) {
				mainWindow2.on('close', function(e) {
					e.preventDefault();
					mainWindow2.webContents.send('remote', {
						msg: 'exit'
					});
				});
				mainWindow.destroy();
			}
		});
	}
	loadMainWin();
	loadIpc();
});

app.on('window-all-closed', (e) => {
	if (process.platform != 'darwin')
		app.quit();
});
