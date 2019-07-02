const { app, Menu, Tray } = require('electron');
const frontendDir = __dirname + '/frontend/';
const iconFile = __dirname + '/icon.png';
let tray = null;

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	let mainWindow3 = null;
	let debugMode = false;
	let startAsTray = false;

	function init() {
		let ret;
		let configPath;
		const configFile = 'config.js';
		const fs = require("fs-extra");
		mainWindow3 = createWin(400, 345, false, 'config.html', false);

		configPath = configFile;
		if (!fs.existsSync(configPath)) {
			const appDataPath = process.env.APPDATA + "/PPT-NDI";
			configPath = appDataPath + "/" + configFile;
		}
		if (fs.existsSync(configPath)) {
			startAsTray = JSON.parse(fs.readFileSync(configPath)).startAsTray;
		} else {
			// Do nothing
		}
		ret = loadArg();
		if (!ret) {
			loadMainWin(!startAsTray);
		}
		loadIpc();

		tray = new Tray(iconFile);
		const contextMenu = Menu.buildFromTemplate([
			{ label: '&Hide', click() {
				if (mainWindow2 != null) {
					mainWindow2.hide();
				} else if (mainWindow != null) {
					mainWindow.hide();
				}
			}},
			{ label: '&Configure', click() {
				mainWindow3.show();
			}},
			{ label: 'E&xit', click() {
				if (mainWindow2 != null) {
					mainWindow2.destroy();
				}
				if (process.platform != 'darwin') {
					app.quit();
				}
			}}
		]);
		tray.setToolTip('PPT-NDI');
		tray.setContextMenu(contextMenu);
		tray.on('double-click', () => {
			if (mainWindow2 != null) {
				mainWindow2.show();
			} else if (mainWindow != null) {
				mainWindow.show();
			}
		});
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
				mainWindow2 = createWin(300, 300, false, 'control.html', false);
				addMainWin2closeHandler();
				break;
			}
			if (/--slideshow/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(300, 330, false, 'control.html', !startAsTray);
				addMainWin2closeHandler();
				break;
			}
			if (/--classic/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(1200, 680, true, 'index.html', !startAsTray);
				addMainWin2closeHandler();
				break;
			}
		}
		return matched;
	}

	function loadMainWin(showWin) {
		mainWindow = createWin(700, 360, false, 'main.html', showWin);
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

	function createWin(width, height, maximizable, winFile, showWin) {
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
			webPreferences: {
				webSecurity: false,
				nodeIntegration: true,
				nodeIntegrationInWorker: true
			}
		});
		if (!maximizable) {
			retData.setMaximumSize(width, height);
		}
		if (debugMode) {
			retData.webContents.openDevTools();
		}
		retData.loadURL(frontendDir + winFile);
		if (!showWin) {
			retData.hide();
		} else {
			retData.focus();
		}
		return retData;
	}

	function addMainWin2closeHandler() {
		if (mainWindow2 !== null) {
			mainWindow2.on('close', function(e) {
				e.preventDefault();
				mainWindow2.webContents.send('remote', {
					msg: 'exit'
				});
			});
		}
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
					mainWindow2 = createWin(300, 330, false, 'control.html', true);
					addMainWin2closeHandler();
					mainWindow.destroy();
					break;
				case "select2":
					mainWindow2 = createWin(1200, 680, true, 'index.html', true);
					addMainWin2closeHandler();
					mainWindow.destroy();
					break;
				case "showConfig":
					mainWindow3.show();
					break;
				case "hideConfig":
					mainWindow3.hide();
					break;
				case "onTop":
					mainWindow2.setAlwaysOnTop(true);
					break;
				case "onTopOff":
					mainWindow2.setAlwaysOnTop(false);
					break;
				case "reflectConfig":
					mainWindow2.webContents.send('remote', { msg: 'reload' });
					break;
				default:
					console.log("Unhandled function - loadIpc()");
					mainWindow.destroy();
					break;
			}
		});
	}
	
	init();
});

app.on('window-all-closed', (e) => {
	if (process.platform != 'darwin') {
		app.quit();
	}
});
