const { app } = require('electron');

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	const { BrowserWindow } = require('electron');
	mainWindow = new BrowserWindow({
		width: 700,
		height: 360,
		title: "",
		icon: __dirname + '/icon.png',
		resize: false,
		frame: false,
		maximizable: false,
		backgroundColor: '#060621',
		webPreferences: { webSecurity: false, nodeIntegration: true }
	});

	mainWindow.loadURL('file://' + __dirname + '/frontend/main.html');
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

	loadIpc();

	function loadEvent() {
		mainWindow2.on('close', function(e) {
			e.preventDefault();
			mainWindow2.webContents.send('remote', {
				msg: 'exit'
			});
		});
	}

	function loadIpc() {
		const ipc = require('electron').ipcMain;
		ipc.on('remote', (event, data) => {
			if (data == "exit") {
				if (mainWindow2 != null) {
					mainWindow2.destroy();
				}
				if (process.platform != 'darwin') {
					app.quit();
				}
				return;
			}
			if (data == "select1") {
				mainWindow2 = new BrowserWindow({
					width: 300,
					height: 150,
					minWidth: 300,
					minHeight: 150,
					title: "",
					icon: __dirname + '/icon.png',
					resize: false,
					frame: false,
					maximizable: false,
					backgroundColor: '#060621'
				});
				mainWindow2.loadURL('file://' + __dirname + '/frontend/control.html');
				loadEvent();
				mainWindow.destroy();

				//For debugging:
				//mainWindow2.webContents.openDevTools();
			}
			if (data == "select2") {
				mainWindow2 = new BrowserWindow({
					width: 1200,
					height: 680,
					minWidth: 1200,
					minHeight: 680,
					title: "",
					icon: __dirname + '/icon.png',
					resize: true,
					frame: false,
					backgroundColor: '#060621'
				});
				mainWindow2.loadURL('file://' + __dirname + '/frontend/index.html');
				loadEvent();
				mainWindow.destroy();
			}
		});
	}
});

app.on('window-all-closed', (e) => {
	if (process.platform != 'darwin')
		app.quit();
});
