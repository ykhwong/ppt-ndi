const { app } = require('electron');
const frontendDir = __dirname + '/frontend/';
const iconFile = __dirname + '/icon.png';

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	let debugMode = false;

	function init() {
		let ret = loadArg();
		if (!ret) {
			loadMainWin();
		}
		loadIpc();
	}

	function loadArg() {
		let matched=false;
		for (i = 0, len = process.argv.length; i < len; i++) {
			let val = process.argv[i];
			if (/--(h|help)/i.test(val)) {
				const path = require('path');
				let out = "PPT NDI\n";
				out += " " + path.basename(process.argv[0]) + " [--slideshow] [--classic] [--bg]\n";
				out += "   [--slideshow] : SlidShow Mode\n";
				out += "     [--classic] : Classic Mode\n";
				out += "          [--bg] : Run SlideShow Mode as background\n";
				console.log(out);
				app.quit();
			}
			if (/--bg/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(300, 300, false, 'control.html');
				mainWindow2.hide();
				break;
			}
			if (/--slideshow/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(300, 330, false, 'control.html');
				break;
			}
			if (/--classic/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(1200, 680, true, 'index.html');
				break;
			}
		}
		return matched;
	}

	function loadMainWin() {
		mainWindow = createWin(700, 360, false, 'main.html');
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
		if (debugMode) {
			maximizable = true;
		}
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
		retData.loadURL(frontendDir + winFile);
		retData.focus();
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
					mainWindow2 = createWin(300, 330, false, 'control.html');
					break;
				case "select2":
					mainWindow2 = createWin(1200, 680, true, 'index.html');
					break;
				case "onTop":
					mainWindow2.setAlwaysOnTop(true);
					break;
				case "onTopOff":
					mainWindow2.setAlwaysOnTop(false);
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
	
	init();

	//console.log(process.argv);
});

app.on('window-all-closed', (e) => {
	if (process.platform != 'darwin')
		app.quit();
});
