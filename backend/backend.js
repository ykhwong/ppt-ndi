const { app, Menu, Tray, screen, globalShortcut } = require('electron');
const multipleInstance = !app.requestSingleInstanceLock();
const frontendDir = __dirname + '/../frontend/';
const debugMode = false;
let iconFile;
let tray = null;
app.disableHardwareAcceleration();
app.allowRendererProcessReuse = true;

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	let mainWindow3 = null;
	let monitorWin = null;
	let rendererWin = null;
	let startAsTray = false;
	let isMainWinShown = false;
	let isMainWin2shown = false;
	let remoteLib = {};
	let remoteVar = {};
	let lastImageArgs = null;
	let loopPaused = false;

	require('@electron/remote/main').initialize();

	const winData = {
		"mode" : {
			"width" : 600,
			"height" : 330,
			"dest" : "client-mode.html"
		},
		"slideshow" : {
			"width" : 277,
			"height" : 430,
			"dest" : "client-slideshow.html"
		},
		"classic" : {
			"width" : 960,
			"height" : 700,
			"dest" : "client-classic.html"
		},
		"config" : {
			"width" : 320,
			"height" : 465,
			"dest" : "client-config.html"
		},
		"monitor" : {
			"width" : 0,
			"height" : 0,
			"dest" : "monitor.html"
		},
		"renderer": {
			"width" : 1024,
			"height" : 768,
			"dest" : "renderer.html"
		}
	}

	function destroyWin(win) {
		if (win != null && !win.isDestroyed()) {
			win.destroy();
		}
	}

	function refreshTray() {
		let isVisible = true;
		let hideShowItem;
		if (tray === null) {
			tray = new Tray(iconFile);
		}
		if (mainWindow2 != null) {
			isVisible = isMainWin2shown;
		} else if (mainWindow != null) {
			isVisible = isMainWinShown;
		}
		hideShowItem = isVisible ? '&Hide' : '&Show';
		const contextMenu = Menu.buildFromTemplate([
			{ label: hideShowItem, click() {
				if (mainWindow2 != null) {
					if (isVisible) {
						mainWindow2.hide();
					} else {
						mainWindow2.show();
					}
				} else if (mainWindow != null) {
					if (isVisible) {
						mainWindow.hide();
					} else {
						mainWindow.show();
					}
				}
				refreshTray();
			}},
			{ label: '&Configure', click() {
				mainWindow3.show();
			}},
			{ label: 'E&xit', click() {
				loopPaused = true;
				destroyWin(mainWindow2);
				destroyWin(mainWindow3);
				destroyWin(monitorWin);
				destroyWin(rendererWin);
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
			refreshTray();
		});
	}

	function init() {
		let ret;
		let configPath;
		const configFile = 'config.js';
		const fs = require("fs-extra");

		if (process.platform === 'win32') {
			iconFile = __dirname + '/img/icon.ico';
		} else {
			iconFile = __dirname + '/img/icon.png';
		}

		mainWindow3 = createWin(winData.config.width, winData.config.height, false, winData.config.dest, false, false);
		mainWindow3.on('close', function (event) {
			event.preventDefault();
			mainWindow3.hide();
		});
		mainWindow3.setAlwaysOnTop(true);

		configPath = configFile;
		if (!fs.existsSync(configPath)) {
			const appDataPath = (process.env.APPDATA || (process.platform === 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + "/.local/share")) + "/PPT-NDI";
			configPath = appDataPath + "/" + configFile;
		}
		if (fs.existsSync(configPath)) {
			startAsTray = JSON.parse(fs.readFileSync(configPath)).startAsTray;
		} else {
			// Do nothing
		}
		ret = loadArg();
		if (!ret) {
			if (process.platform === 'win32') {
				loadMainWin(!startAsTray);
			} else {
				mainWindow2 = createWin(winData.classic.width, winData.classic.height, true, winData.classic.dest, !startAsTray, false);
				addMainWin2handler(!startAsTray);
				registerFocusInfo(mainWindow2);
			}
		}
		loadIpc();
		sendLoop();
		refreshTray();
		monitorWin = createWin(winData.monitor.width, winData.monitor.height, false, winData.monitor.dest, false, true);
		rendererWin = createWin(winData.renderer.width, winData.renderer.height, false, winData.renderer.dest, debugMode ? true : false, true);
		monitorWin.setAlwaysOnTop(true);
		monitorWin.on('close', function (event) {
			event.preventDefault();
			monitorWin.hide();
		});
	}

	function loadArg() {
		let matched=false;
		for (i = 0, len = process.argv.length; i < len; i++) {
			let val = process.argv[i];
			if (/^(-h|--help|\/\?)/i.test(val)) {
				const path = require('path');
				let out = "PPT NDI\n";
				out += " " + path.basename(process.argv[0]) + " [--slideshow] [--classic]\n";
				out += "   [--slideshow] : SlidShow Mode\n";
				out += "     [--classic] : Classic Mode\n";
				console.log(out);
				process.exit(0);
			} else if (/^--slideshow/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(winData.slideshow.width, winData.slideshow.height, false, winData.slideshow.dest, !startAsTray, false);
				addMainWin2handler(!startAsTray);
				break;
			} else if (/^--classic/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(winData.classic.width, winData.classic.height, true, winData.classic.dest, !startAsTray, false);
				addMainWin2handler(!startAsTray);
				registerFocusInfo(mainWindow2);
				break;
			} else if (/^(-[^-]|--\S)/.test(val)) {
				console.log("Unknown switch: " + val);
				process.exit(0);
			}
		}
		return matched;
	}

	function loadMainWin(showWin) {
		mainWindow = createWin(winData.mode.width, winData.mode.height, false, winData.mode.dest, showWin, false);
		mainWindow.on('closed', function(e) {
			if (mainWindow2 === null) {
				mainWindow = null;
				destroyWin(mainWindow3);
				if (process.platform != 'darwin') {
					app.quit();
				}
			} else {
				e.preventDefault();
			}
		});
		mainWindow.on('show', () => {
			isMainWinShown = true;
		});
		mainWindow.on('hide', () => {
			isMainWinShown = false;
		});
		if (showWin) {
			mainWindow.show();
		}
	}

	function createWin(width, height, maximizable, winFile, showWin, isTransparent) {
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
			webPreferences: {
				webSecurity: false,
				nodeIntegration: true,
				enableRemoteModule: true,
				contextIsolation: false
			},
			transparent : isTransparent,
			backgroundColor: isTransparent?'#00051336':'#060621'
		});
		
		require("@electron/remote/main").enable(retData.webContents);
		if (!maximizable) {
			retData.setMaximumSize(width, height);
		}
		if (debugMode) {
			retData.webContents.openDevTools();
		}

		retData.loadURL('file://' + frontendDir + winFile);
		if (!showWin) {
			retData.hide();
		} else {
			retData.focus();
		}
		return retData;
	}

	function addMainWin2handler(showWin) {
		if (mainWindow2 !== null) {
			mainWindow2.on('close', function(e) {
				loopPaused = true;
				e.preventDefault();
				mainWindow2.webContents.send('remote', {
					msg: 'exit'
				});
			});
		}
		mainWindow2.on('show', () => {
			isMainWin2shown = true;
		});
		mainWindow2.on('hide', () => {
			isMainWin2shown = false;
		});
		if (showWin) {
			mainWindow2.show();
		}
	}

	function registerFocusInfo(myWin) {
		if (myWin == null) {
			return;
		}
		myWin.on('focus', () => {
			myWin.webContents.send('remote', { msg: 'focused' });
		});
		myWin.on('blur', () => {
			myWin.webContents.send('remote', { msg: 'blurred' });
		});
	}

	function sleep(ms){
		return new Promise(resolve=>{
			setTimeout(resolve,ms)
		})
	}

	async function sendLoop() {
		let sleepCnt = 0;
		while (true) {
			if (
				loopPaused ||
				typeof remoteVar.configData === 'undefined' ||
				!remoteVar.configData.highPerformance
			) {
				await sleep(1000);
				continue;
			}
			if (sleepCnt === 0) {
				sleepCnt = 1000;
			}
			await sleep(sleepCnt);
			if (lastImageArgs !== null) {
				if (lastImageArgs[1] === false) {
					sleepCnt = 100;
					remoteVar.lib.send( ...lastImageArgs );
				}
			} else {
				sleepCnt = 1000;
			}
		}
	}

	function loadIpc() {
		const ipc = require('electron').ipcMain;

		ipc.on('remote', (event, data) => {
			switch (data.name) {
				case "exit":
					loopPaused = true;
					destroyWin(mainWindow2);
					destroyWin(mainWindow3);
					destroyWin(monitorWin);
					destroyWin(rendererWin);
					if (process.platform != 'darwin') {
						app.quit();
					}
					break;
				case "select1":
					mainWindow2 = createWin(winData.slideshow.width, winData.slideshow.height, false, winData.slideshow.dest, true, false);
					addMainWin2handler(true);
					destroyWin(mainWindow);
					break;
				case "select2":
					mainWindow2 = createWin(winData.classic.width, winData.classic.height, true, winData.classic.dest, true, false);
					addMainWin2handler(true);
					registerFocusInfo(mainWindow2);
					destroyWin(mainWindow);
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
					if (mainWindow != null && !mainWindow.isDestroyed()) {
						mainWindow.webContents.send('remote', { msg: 'reload' });
					}
					if (mainWindow2 !== null && !mainWindow2.isDestroyed()) {
						mainWindow2.webContents.send('remote', { msg: 'reload' });
					}
					globalShortcut_proc();
					break;
				case "passConfigData":
					remoteVar.configData = data.details;
					break;
				default:
					console.log("Unhandled function - loadIpc(): " + data);
					destroyWin(mainWindow);
					break;
			}
		});

		ipc.on('renderer', (event, data) => {
			switch (data.name) {
				case "notifyError":
				case "notifyLoaded":
				case "notifyCanceled":
					mainWindow2.webContents.send('renderer', data);
					break;
			}

			switch (data.func) {
				case "load":
				case "cancel":
					rendererWin.webContents.send('renderer', data);
					break;
			}
		});

		ipc.on('monitor', (event, data) => {
			switch (data.func) {
				case "update":
					monitorWin.webContents.send('monitor', data);
					break;
				case "get":
					event.returnValue = screen.getAllDisplays();
					/*
					[
					  {
						id: 2528732444,
						bounds: { x: 0, y: 0, width: 1920, height: 1080 },
						workArea: { x: 0, y: 0, width: 1920, height: 1040 },
						accelerometerSupport: 'unknown',
						monochrome: false,
						colorDepth: 24,
						colorSpace: '{primaries:BT709, transfer:IEC61966_2_1, matrix:RGB, range:FULL}',
						depthPerComponent: 8,
						size: { width: 1920, height: 1080 },
						workAreaSize: { width: 1920, height: 1040 },
						scaleFactor: 1,
						rotation: 0,
						internal: false,
						touchSupport: 'unknown'
					  }
					]
					*/
					break;
				case "assign":
					if (data.monitorNo >= 1) {
						let disps = screen.getAllDisplays();
						let disp = disps[data.monitorNo - 1];
						
						if (disp) {
							monitorWin.setBounds({
								x: disp.bounds.x,
								y: disp.bounds.y
							});
							monitorWin.width = disp.bounds.width;
							monitorWin.height = disp.bounds.height;
						}
					}
					break;
				case "turnOn":
					monitorWin.show();
					monitorWin.setFullScreen(true);
					monitorWin.setResizable(false);
					break;
				case "turnOff":
					monitorWin.hide();
					break;
				case "transparentOn":
					monitorWin.setBackgroundColor('#00051336');
					monitorWin.webContents.send('monitor', data);
					break;
				case "transparentOff":
					monitorWin.setBackgroundColor('black');
					monitorWin.webContents.send('monitor', data);
					break;
				case "monitorBlack":
				case "monitorWhite":
				case "monitorTrans":
					monitorWin.webContents.send('monitor', data);
					break;
				default:
					break;
			}
		});

		function globalShortcut_proc() {
			const mainKey = 'Ctrl+Shift+';
			if (typeof remoteVar.configData === 'undefined') return;
			globalShortcut.unregisterAll();

			if ( remoteVar.configData.hotKeys.prev ) {
				globalShortcut.register(mainKey + remoteVar.configData.hotKeys.prev, () => {
					mainWindow2.webContents.send('remote', { msg: 'gotoPrev' });
				});
			}
			if ( remoteVar.configData.hotKeys.next ) {
				globalShortcut.register(mainKey + remoteVar.configData.hotKeys.next, () => {
					mainWindow2.webContents.send('remote', { msg: 'gotoNext' });
				});
			}
			if ( remoteVar.configData.hotKeys.transparent ) {
				globalShortcut.register(mainKey + remoteVar.configData.hotKeys.transparent, () => {
					mainWindow2.webContents.send('remote', { msg: 'update_trn' });
				});
			}
			if ( remoteVar.configData.hotKeys.black ) {
				globalShortcut.register(mainKey + remoteVar.configData.hotKeys.black, () => {
					mainWindow2.webContents.send('remote', { msg: 'update_black' });
				});
			}
			if ( remoteVar.configData.hotKeys.white ) {
				globalShortcut.register(mainKey + remoteVar.configData.hotKeys.white, () => {
					mainWindow2.webContents.send('remote', { msg: 'update_white' });
				});
			}
		}

		ipc.on('status', (event, data) => {
			let ret = null;
			switch (data.item) {
				case "multipleInstance":
					ret = multipleInstance;
					break;
				default:
					break;
			}
			event.returnValue = ret;
		});

		ipc.on('require', (event, data) => {
			/*
				data.lib : string
				data.func : string
				data.on : string
				data.args : array
			*/
			let ret = -1;
			switch (data.lib) {
				case "ffi":
					if (data.func === null && data.args === null) {
						remoteLib.ffi = require("ffi-napi");
						try {
							if (process.platform === 'win32') {
								remoteVar.lib = remoteLib.ffi.Library(
									app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/PPTNDI.dll', {
									'init': [ 'int', [] ],
									'destroy': [ 'int', [] ],
									'send': [ 'int', [ "string", "bool" ] ]
								});
							} else if (process.platform === 'darwin') {
								remoteVar.lib = remoteLib.ffi.Library(
									app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/PPTNDI.dylib', {
									'init': [ 'int', [] ],
									'destroy': [ 'int', [] ],
									'send': [ 'int', [ "string", "bool" ] ]
								});
							} else {
								console.error('Unsupported platform ' + process.platform);
							}
							ret = remoteVar.lib;
						} catch(e) {
							console.log("remoteLib failed: " + e);
						}
					}
					if (data.func === "init") {
						let ret = -1;
						if (typeof remoteVar.lib !== 'undefined') {
							ret = remoteVar.lib.init();
						}
					} else if (data.func === "destroy") {
						let ret = -1;
						if (typeof remoteVar.lib !== 'undefined') {
							ret = remoteVar.lib.destroy();
						}
					} else if (data.func === "send") {
						let ret = -1;
						if (typeof remoteVar.lib !== 'undefined') {
							lastImageArgs = data.args;
							ret = remoteVar.lib.send( ...data.args );
						}
					}
					break;
				case "electron-globalShortcut":
					ret = globalShortcut_proc();
					break;
				default:
					break;
			}
			if (ret < 0 || ret > 0) {
				event.returnValue = ret;
			} else {
				event.returnValue = 0;
			}
		});

	}

	init();
});

app.on('window-all-closed', (e) => {
	loopPaused = true;
	app.quit();
});
