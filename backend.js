const { app, Menu, Tray } = require('electron');
const frontendDir = __dirname + '/frontend/';
let iconFile;
let tray = null;
app.disableHardwareAcceleration();

app.on('ready', function() {
	let mainWindow = null;
	let mainWindow2 = null;
	let mainWindow3 = null;
	let debugMode = false;
	let startAsTray = false;
	let isMainWinShown = false;
	let isMainWin2shown = false;
	let remoteLib = {};
	let remoteVar = {};
	let lastImageArgs = null;
	let loopPaused = false;

	const winData = {
		"home" : {
			"width" : 600,
			"height" : 330,
			"dest" : "main.html"
		},
		"control" : {
			"width" : 277,
			"height" : 430,
			"dest" : "control.html"
		},
		"classic" : {
			"width" : 1200,
			"height" : 700,
			"dest" : "index.html"
		},
		"config" : {
			"width" : 320,
			"height" : 345,
			"dest" : "config.html"
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
				if (mainWindow2 != null) {
					mainWindow2.destroy();
				}
				if (mainWindow3 != null) {
					mainWindow3.destroy();
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
			refreshTray();
		});
	}

	function init() {
		let ret;
		let configPath;
		const configFile = 'config.js';
		const fs = require("fs-extra");

		if (process.platform === 'win32') {
			iconFile = __dirname + '/icon.ico';
		} else {
			iconFile = __dirname + '/icon.png';
		}

		mainWindow3 = createWin(winData.config.width, winData.config.height, false, winData.config.dest, false);
		mainWindow3.on('close', function (event) {
			event.preventDefault();
			mainWindow3.hide();
		});

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
		sendLoop();
		refreshTray();
	}

	function loadArg() {
		let matched=false;
		for (i = 0, len = process.argv.length; i < len; i++) {
			let val = process.argv[i];
			if (/--(h|help)/i.test(val)) {
				const path = require('path');
				let out = "PPT NDI\n";
				out += " " + path.basename(process.argv[0]) + " [--slideshow] [--classic]\n";
				out += "   [--slideshow] : SlidShow Mode\n";
				out += "     [--classic] : Classic Mode\n";
				console.log(out);
				process.exit(0);
			}
			if (/--slideshow/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(winData.control.width, winData.control.height, false, winData.control.dest, !startAsTray);
				addMainWin2handler(!startAsTray);
				break;
			}
			if (/--classic/i.test(val)) {
				matched=true;
				mainWindow2 = createWin(winData.classic.width, winData.classic.height, true, winData.classic.dest, !startAsTray);
				addMainWin2handler(!startAsTray);
				registerFocusInfo(mainWindow2);
				break;
			}
		}
		return matched;
	}

	function loadMainWin(showWin) {
		mainWindow = createWin(winData.home.width, winData.home.height, false, winData.home.dest, showWin);
		mainWindow.on('closed', function(e) {
			if (mainWindow2 === null) {
				mainWindow = null;
				mainWindow3.destroy();
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
				nodeIntegration: true
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
					if (mainWindow2 != null) {
						mainWindow2.destroy();
					}
					if (mainWindow3 != null) {
						mainWindow3.destroy();
					}
					if (process.platform != 'darwin') {
						app.quit();
					}
					break;
				case "select1":
					mainWindow2 = createWin(winData.control.width, winData.control.height, false, winData.control.dest, true);
					addMainWin2handler(true);
					mainWindow.destroy();
					break;
				case "select2":
					mainWindow2 = createWin(winData.classic.width, winData.classic.height, true, winData.classic.dest, true);
					addMainWin2handler(true);
					registerFocusInfo(mainWindow2);
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
					if (mainWindow2 !== null) {
						mainWindow2.webContents.send('remote', { msg: 'reload' });
					}
					break;
				case "passConfigData":
					remoteVar.configData = data.details;
					break;
				case "passIgnoreIoHookVal":
					remoteVar.ignoreIoHook = data.details;
					break;
				default:
					console.log("Unhandled function - loadIpc(): " + data);
					mainWindow.destroy();
					break;
			}
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
						remoteLib.ffi = require("ffi");
						try {
							let { RTLD_NOW, RTLD_GLOBAL } = remoteLib.ffi.DynamicLibrary.FLAGS;
							remoteLib.ffi.DynamicLibrary(
								app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/Processing.NDI.Lib.x64.dll',
								RTLD_NOW | RTLD_GLOBAL
							);
							remoteVar.lib = remoteLib.ffi.Library(app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/PPTNDI.dll', {
								'init': [ 'int', [] ],
								'destroy': [ 'int', [] ],
								'send': [ 'int', [ "string", "bool" ] ]
							});
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
				case "iohook":
					if (data.on === null && data.args === null) {
						remoteLib.ioHook = require('iohook');
						ret = remoteLib.ioHook;
					} else if (data.func === "start") {
						ret = remoteLib.ioHook.start();
					} else if (data.on === "keyup") {
						remoteLib.ioHook.on('keyup', event => {
							if (typeof mainWindow2 === 'undefined' || mainWindow2 === null) return;
							if (event.shiftKey && event.ctrlKey) {
								let chr = String.fromCharCode( event.rawcode );
								if (chr === "") return;
								switch (chr) {
									case remoteVar.configData.hotKeys.prev:
										mainWindow2.webContents.send('remote', { msg: 'gotoPrev' });
										break;
									case remoteVar.configData.hotKeys.next:
										mainWindow2.webContents.send('remote', { msg: 'gotoNext' });
										break;
									case remoteVar.configData.hotKeys.transparent:
										mainWindow2.webContents.send('remote', { msg: 'update_trn' });
										break;
									case remoteVar.configData.hotKeys.black:
										mainWindow2.webContents.send('remote', { msg: 'update_black' });
										break;
									case remoteVar.configData.hotKeys.white:
										mainWindow2.webContents.send('remote', { msg: 'update_white' });
										break;
									default:
										break;
								}
							}
						});
						ret = 1;
					} else if (data.on === "keydown") {
						remoteLib.ioHook.on('keydown', event => {
							if (typeof mainWindow2 === 'undefined' || mainWindow2 === null || remoteVar.ignoreIoHook) return;
							if (((event.altKey || event.shiftKey || event.ctrlKey || event.metaKey) && event.keycode === 63) ||
							!(event.altKey || event.shiftKey || event.ctrlKey || event.metaKey) && (
							event.keycode === 63 || event.keycode === 3655 || event.keycode === 3663 ||
							event.keycode === 28 || event.keycode === 57 || event.keycode === 48 || event.keycode === 17 ||
							event.keycode === 49 || event.keycode === 25 || event.keycode === 60999 || event.keycode === 61007 ||
							event.keycode === 61003 || event.keycode === 61000 || event.keycode === 61005 || event.keycode === 61008
							)) {
								mainWindow2.webContents.send('remote', { msg: 'stdin_write_newline' });
							}
						});
						ret = 1;
					} else if (data.on === "mouseup") {
						remoteLib.ioHook.on('mouseup', event => {
							if (typeof mainWindow2 === 'undefined' || mainWindow2 === null || remoteVar.ignoreIoHook) return;
							loopPaused=false;
							mainWindow2.webContents.send('remote', { msg: 'stdin_write_newline' });
						});
						ret = 1;
					} else if (data.on === "mousewheel") {
						remoteLib.ioHook.on('mousewheel', event => {
							if (typeof mainWindow2 === 'undefined' || mainWindow2 === null || remoteVar.ignoreIoHook) return;
							mainWindow2.webContents.send('remote', { msg: 'stdin_write_newline' });
						});
						ret = 1;
					} else if (data.on === "mousedrag") {
						remoteLib.ioHook.on('mousedrag', event => {
							if (typeof mainWindow2 === 'undefined' || mainWindow2 === null || remoteVar.ignoreIoHook) return;
							if (loopPaused === false) {
								loopPaused=true;
							}
						});
						ret = 1;
					}
					break;
				default:
					break;
			}
			event.returnValue = ret;
		});

	}

	init();
});

app.on('window-all-closed', (e) => {
	if (process.platform != 'darwin') {
		loopPaused = true;
		app.quit();
	}
});
