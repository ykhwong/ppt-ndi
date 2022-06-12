$(document).ready(function() {
	const spawn = require( 'child_process' ).spawn;
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const cscript = require('./js/cscript').script.slideshow;
	const runtimeUrl = "https://aka.ms/vs/16/release/vc_redist.x64.exe";
	const vbsBg = cscript.vbsBg;
	const vbsNoBg = cscript.vbsNoBg;
	const vbsCheckSlide = cscript.vbsCheckSlide;
	const vbsDirectCmd = cscript.vbsDirectCmd;
	const appDataPath = (process.env.APPDATA || (process.platform == 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + "/.local/share")) + "/PPT-NDI";
	let ffi;
	let lib;
	let tmpDir = null;
	let preFile = "";
	let slideWidth = 0;
	let slideHeight = 0;
	let customSlideX = 0;
	let customSlideY = 0;
	let lastSignalTime = 0;
	let inTransition = false;
	let configData = {};
	let pin = true;
	let mustStop = false;
	let res; // vbsBg & vbsNoBg
	let res2; // vbsDirectCmd
	let res3; // vbsCheckSlide
	let duration = "";
	let effect = "";
	let slideIdx = "";
	let slideTranTimers = [];
	let curStatus = "";

	function alertMsg(myMsg) {
		const remote = require('@electron/remote');
		const {dialog} = require('@electron/remote');
		let currentWindow = remote.getCurrentWindow();
		let options;
		options = {
			type: 'info',
			message: myMsg,
			buttons: ["OK"]
		};
		dialog.showMessageBoxSync(currentWindow, options);
	}

	function setLangRsc() {
		setLangRscDiv("#show-checkerboard", "ui-slideshow/show-checkerboard", true, configData.lang);
		setLangRscDiv("#enable-slide-transition-effect", "ui-slideshow/enable-slide-transition-effect", true, configData.lang);
		setLangRscDiv("#include-background", "ui-slideshow/include-background", true, configData.lang);
		setLangRscDiv("#customRes", "ui-slideshow/customRes", true, configData.lang);
		setLangRscDiv("#setRes", "ui-slideshow/setRes", false, configData.lang);
		setLangRscDiv("#config", "ui-slideshow/config", false, configData.lang);
		setLangRscDiv("#prevTxt", "ui-slideshow/prevTxt", false, configData.lang);
		setLangRscDiv("#pinTxt", "ui-slideshow/pinTxt", false, configData.lang);

		switch (configData.lang) {
			case "ko":
				$("#slideRes").css("left", "100px");
				$("#pinTxt").css({
					"text-align": "center",
					"font-size": "10px",
					"left": "232px"
				});
				break;
			case "en":
			default:
				$("#config").css("width", "90px");
				$("#pinTxt").css({
					"text-align": "center",
					"font-size": "8px",
					"left": "230px"
				});
				break;
		}
	}

	function relocateTitlebarElements() {
		switch (process.platform) {
			case 'darwin':
				$("#closeImg").css({ left: "0px" });
				$("#closeImg").show();
				break;
			default:
				$("#closeImg").css({ right: "5px" });
				$("#closeImg").show();
				break;
		}
	}

	function runLib() {
		ffi = ipc.sendSync("require", { lib: "ffi", func: null, args: null });
		if ( ffi === -1 ) {

			let librariesFound = false;
			if (process.platform === 'win32') {
				alertMsg(getLangRsc("ui-slideshow/dll-init-failed", configData.lang));
				if (fs.existsSync("PPTNDI.DLL") && fs.existsSync("Processing.NDI.Lib.x64.dll")) {
					librariesFound = true;
				}
			} else if (process.platform === 'darwin') {
				alertMsg(getLangRsc("ui-slideshow/dylib-init-failed", configData.lang));
				if (fs.existsSync("PPTNDI.dylib")) {
					librariesFound = true;
				}
			}

			if (librariesFound) {
				const execRuntime = require('child_process').execSync;
				execRuntime("start " + runtimeUrl, (error, stdout, stderr) => { 
					callback(stdout);
				});
			}
			ipc.send('remote', { name: "exit" });
		}

		lib = ipc.sendSync("require", { lib: "ffi", func: "init", args: null });
		if (lib === 1) {
			alertMsg(getLangRsc("ui-slideshow/failed-to-create-listening-svr", configData.lang));
			ipc.send('remote', { name: "exit" });
			return;
		}
	}

	function stopSlideTransition() {
		for (let pp=2; pp<=9; pp++) {
			clearTimeout(slideTranTimers[pp]);
		}
		mustStop = true;
	}

	function sendColorNDI(color) {
		const now = new Date().getTime();
		const PNG = require('pngjs').PNG;

		let file;
		let png;
		let buffer;
		let colorInfo = {};
		let mWidth = (slideWidth === 0 ? 1920 : slideWidth);
		let mHeight = (slideHeight === 0 ? 1080 : slideHeight);

		switch (color) {
			case "black":
				file = tmpDir + "/SlideBlack.png";
				colorInfo = { r: 0, g: 0, b: 0, alpha: 255 }
				break;
			case "white":
				file = tmpDir + "/SlideWhite.png";
				colorInfo = { r: 255, g: 255, b: 255, alpha: 255 };
				break;
			case "tran":
				file = tmpDir + "/SlideTran.png";
				colorInfo = { r: 255, g: 255, b: 255, alpha: 0 };
				break;
		}

		png = new PNG({
			width: mWidth,
			height: mHeight,
			filterType: -1
		});

		for (let y = 0; y < png.height; y++) {
			for (let x = 0; x < png.width; x++) {
				let idx = (png.width * y + x) << 2;
				png.data[idx  ] = colorInfo.r;
				png.data[idx+1] = colorInfo.g;
				png.data[idx+2] = colorInfo.b;
				png.data[idx+3] = colorInfo.alpha;
			}
		}

		buffer = PNG.sync.write(png);
		fs.writeFileSync(file, buffer);
		$("#slidePreview").attr("src", file + "?" + now);
		ipc.sendSync("require", {
			lib: "ffi",
			func: "send",
			args: [ file, false ]
		});
	}

	function updateStat(cmd, details) {
		let msg = getLangRsc("ui-slideshow/status", configData.lang);
		curStatus = getLangRsc("ui-slideshow/status-ready", configData.lang);
		msg = msg + cmd;
		if (/\S/.test(details)) {
			msg = msg + "<br />" + details;
		}
		$("#tip").html(msg);
	}

	function sendNDI(file, data) {
		const now = new Date().getTime();
		const cmd = data.toString();
		let newSlideIdx;
		preFile = tmpDir + "/SlidePre.png";
		stopSlideTransition();
		if (/^PPTNDI: Sent /.test(cmd)) {
			let tmpCmd = cmd.replace(/^PPTNDI: Sent /, "");
			duration = tmpCmd.split(" ")[0].trim();
			effect = tmpCmd.split(" ")[1].trim();
			newSlideIdx = tmpCmd.split(" ")[2].trim();
		} else if(/^PPTNDI: White/.test(cmd)) {
			file = tmpDir + "/SlideWhite.png";
			newSlideIdx = "white";
			preFile = "";
		} else if(/^PPTNDI: Black/.test(cmd)) {
			file = tmpDir + "/SlideBlack.png"
			newSlideIdx = "black";
			preFile = "";
		} else if(/^PPTNDI: Done/.test(cmd)) {
			updateStat(
				getLangRsc("ui-slideshow/status-end-of-slideshow", configData.lang),
				""
			);
			return;
		//} else if(/^PPTNDI: Paused/.test(cmd)) {
		//	updateStat("PAUSED", "");
		//	return;
		} else if(/^PPTNDI: Ready/.test(cmd)) {
			updateStat(
				getLangRsc("ui-slideshow/status-ready", configData.lang),
				getLangRsc("ui-slideshow/status-ppt-start-slideshow", configData.lang)
			);
			return;
		} else if(/^PPTNDI: NoPPT/.test(cmd)) {
			updateStat(
				getLangRsc("ui-slideshow/status-error", configData.lang),
				getLangRsc("ui-slideshow/status-ppt-not-found", configData.lang)
			);
			return;
		} else {
			console.log(cmd);
			return;
		}

		updateStat(
			getLangRsc("ui-slideshow/status-ok", configData.lang),
			getLangRsc("ui-slideshow/status-request-complete", configData.lang)
		);
		if (/^PPTNDI: Sent /.test(cmd)) {
			let fd;
			try {
				fd = fs.openSync(file, 'r+');
			} catch (err) {
				if (err && err.code === 'EBUSY'){
					if (fd !== undefined) {
						fs.closeSync(fd);
					}
					sendNDI(file, data);
					return;
				}
			}
			if (fd !== undefined) {
				fs.closeSync(fd);
			}

			function getMeta(url, callback) {
				let img = new Image();
				img.src = url;
				img.onload = function() { callback(this.width, this.height); }
			}
			getMeta(
			  file + "?" + now,
			  function(width, height) { 
				slideWidth = width;
				slideHeight = height;
				$("#slideRes").html("( " + slideWidth + " x " + slideHeight + " )");
			  }
			);
			$("#slidePreview").attr("src", file + "?" + now);
		}

		if (slideIdx === newSlideIdx) {
			if (lastSignalTime >= (Date.now() - 500)) {
				return;
			}
		}
		console.log(cmd);
		slideIdx = newSlideIdx;
		lastSignalTime = Date.now();

		if (newSlideIdx === "black" || newSlideIdx === "white") {
			sendColorNDI(newSlideIdx);
		} else {
			if (preFile !== "") {
				let buf1;
				let buf2;
				if (fs.existsSync(preFile) && fs.existsSync(file)) {
					buf1 = fs.readFileSync(preFile);
					buf2 = fs.readFileSync(file);
					if (buf1.equals(buf2)) {
						return;
					}
				}
			}
			if ($("#slide_tran").is(":checked")) {
				if(!/^\s*0\s*$/.test(effect)) {
					if (fs.existsSync(preFile)) {
						mustStop = false;
						procTransition(file, data);
						return;
					}
				}
			}

			try {
				fs.copySync(file, preFile);
			} catch(e) {
				console.log("file could not be generated: "+ preFile);
			}
			ipc.sendSync("require", {
				lib: "ffi",
				func: "send",
				args: [ file, false ]
			});
		}
	}

	function handleHook(cmd) {
		switch (cmd) {
			case "prev":
			case "next":
				res2.stdin.write(cmd + "\n");
				res.stdin.write("\n");
				break;
			case "tran":
				setTimeout(function() {
					sendColorNDI("tran");
				}, 500);
				break;
			case "black":
			case "white":
				if (slideWidth === 0 || curStatus === getLangRsc("ui-slideshow/status-ready", configData.lang)) {
					sendColorNDI(cmd);
				} else {
					res2.stdin.write(cmd + "\n");
					res.stdin.write("\n");
				}
				break;
			default:
				break;
		}
	}

	function registerGlobalShortcut() {
		ipc.send("require", { lib: "electron-globalShortcut", func: "control", args: null });
	}

	function procTransition(file, data) {
		const transLvl=9;
		inTransition = true;
		preFile = tmpDir + "/SlidePre.png";

		try {
			for (let i=2; i<=transLvl; i++) {
				fs.unlinkSync(tmpDir + "/t" + i.toString() + ".png");
			}
		} catch(e) {
		}

		function sendSlides(i) {
			console.log(i);
			if (mustStop) {
				inTransition = false;
				return;
			}
			function setLast() {
				if (mustStop) {
					inTransition = false;
					return;
				}
				slideTranTimers[10] = setTimeout(function() {
					ipc.sendSync("require", {
						lib: "ffi",
						func: "send",
						args: [ tmpDir + "/Slide.png", false ]
					});
					if (fs.existsSync(file)) {
						try {
							fs.copySync(file, preFile);
						} catch(e) {
							console.log("file could not be generated: "+ preFile);
						}
					}
					inTransition = false;
				}, 10 * parseFloat(duration) * 50);
			}

			slideTranTimers[i] = setTimeout(function() {
				ipc.sendSync("require", {
					lib: "ffi",
					func: "send",
					args: [ tmpDir + "/t" + i.toString() + ".png", false ]
				});
			}, i * parseFloat(duration) * 50);
			if (i === transLvl) {
				const now = new Date().getTime();
				setLast();
				$("#slidePreview").attr("src", file + "?" + now);
			}
		}

		function doTrans() {
			const mergeImages = require('merge-images');
			stopSlideTransition();
			mustStop = false;

			for (let i=2; i <= transLvl; i++) {
				let now = new Date().getTime();
				mergeImages([
					{ src: preFile + "?" + now, opacity: 1 - (0.1 * i) },
					{ src: file + "?" + now, opacity: 0.1 * i }
				])
				.then(b64 => {
					let b64data = b64.replace(/^data:image\/png;base64,/, "");
					try {
						fs.writeFileSync(tmpDir + "/t" + i.toString() + ".png", b64data, 'base64');
						if (i === 8) {
							for (let i2=2; i2<=transLvl; i2++) {
								sendSlides(i2);
							}
						}
					} catch(e) {
					}
				});
			};
		}
		doTrans();
	}

	function init() {
		const remote = require('@electron/remote');
		let file;
		let vbsDir;
		let vbsDir2;
		let newVbsContent;
		let now = new Date().getTime();
		let multipleInstance = false;
		try {
			process.chdir(remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, ""));
		} catch(e) {
		}
		relocateTitlebarElements();
		$.ajaxSetup({
			async: false
		});
		reflectConfig();
		runLib();

		if (process.platform === 'darwin') {
			tmpDir = process.env.TMPDIR + '/ppt_ndi';
		} else { // win32
			tmpDir = process.env.PROGRAMDATA + '/PPT-NDI/temp';
		}
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir, { recursive: true });
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir, { recursive: true });
		vbsDir = tmpDir + '/wb.vbs';
		vbsDir2 = tmpDir + '/wb2.vbs';
		vbsDir3 = tmpDir + '/wb3.vbs';
		file = tmpDir + "/Slide.png";

		newVbsContent = vbsNoBg;
		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			alertMsg(getLangRsc("ui-slideshow/failed-to-access-tempdir", configData.lang));
			return;
		}
		try {
			fs.writeFileSync(vbsDir2, vbsDirectCmd, 'utf-8');
		} catch(e) {
		}
		try {
			fs.writeFileSync(vbsDir3, vbsCheckSlide, 'utf-8');
		} catch(e) {
			alertMsg(getLangRsc("ui-slideshow/failed-to-access-tempdir", configData.lang));
			return;
		}
		if (fs.existsSync(vbsDir)) {
			let resX = 0;
			let resY = 0;
			if (customSlideX == 0 || customSlideY == 0 || !/\S/.test(customSlideX) || !/\S/.test(customSlideY)) {
				resX = 0;
				resY = 0;
			} else {
				resX = customSlideX;
				resY = customSlideY;
			}

			res = spawn( 'cscript.exe', [ vbsDir, tmpDir, resX, resY, "//NOLOGO", '' ] );
			res.stdout.on('data', function(data) {
				sendNDI(file, data);
			});
		} else {
			alertMsg(getLangRsc("ui-slideshow/failed-to-parse-presentation", configData.lang));
			return;
		}
		if (fs.existsSync(vbsDir2)) {
			res2 = spawn( 'cscript.exe', [ vbsDir2, "//NOLOGO", '' ] );
		}
		if (fs.existsSync(vbsDir3)) {
			res3 = spawn( 'cscript.exe', [ vbsDir3, "//NOLOGO", '' ] );
		} else {
			alertMsg(getLangRsc("ui-slideshow/failed-to-parse-presentation", configData.lang));
			return;
		}

		res3.stdout.on('data', function(data) {
			console.log(data.toString());
			let curSlideStat = data.toString().replace(/^Status: /, "");
			if (/^\s*OFF\s*$/.test(curSlideStat)) {
				// Ready
				updateStat(
					getLangRsc("ui-slideshow/status-ready", configData.lang),
					getLangRsc("ui-slideshow/open-ppt-file", configData.lang)
				);
			} else if (/^\s*0\s*$/.test(curSlideStat)) {
				// Not found
				// updateStat("-", "");
			} else {
				// ON
				res.stdin.write("\n");
			}
		});

		// Enable Always On Top by default
		ipc.send('remote', { name: "onTop" });
		$("#pin").attr("src", "./img/pin_green.png");
		pin = true;

		// Enable Slide Checkerboard by default
		$("#slidePreview").css('background-image', "url('./img/trans_slide.png')");

		registerGlobalShortcut();

		$("#resWidth").val("0");
		$("#resHeight").val("0");
		$("#resWidth, #resHeight").click(function() {
			$(this).val("");
		});
		$("#setRes").click(function() {
			let resX = $("#resWidth").val();
			let resY = $("#resHeight").val();
			if (/^\d+$/.test(resX) && /^\d+$/.test(resY)) {
				res.stdin.write("setRes " + resX + "x" + resY + "\n");
				customSlideX = parseInt(resX, 10);
				customSlideY = parseInt(resY, 10);
			}
		});
		multipleInstance = ipc.sendSync("status", { item: "multipleInstance" });
		if (multipleInstance) {
			alertMsg(getLangRsc("ui-slideshow/no-support-multiple-instances", configData.lang));
			cleanupForExit();
		}
		
		reflectCache(false);
	}

	function cleanupForTemp() {
		if (fs.existsSync(tmpDir)) {
			fs.removeSync(tmpDir);
		}
	}

	function reflectConfig() {
		const configFile = 'config.js';
		let configPath = "";
		const remote = require('@electron/remote');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		if (!fs.existsSync(configPath)) {
			configPath = appDataPath + "/" + configFile;
		}
		if (fs.existsSync(configPath)) {
			$.getJSON(configPath, function(json) {
				configData.hotKeys = json.hotKeys;
				//configData.highPerformance = json.highPerformance;
				configData.lang = json.lang;
				setLangRsc();
				ipc.send('remote', { name: "passConfigData", details: configData });
			});
		} else {
			// Do nothing
		}
	}

	function reflectCache(saveOnly) {
		const configFile = 'cache_control.js';
		let configPath = "";
		const remote = require('@electron/remote');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		const cacheData = {
			"showCheckerboard": $("#trans_checker").is(":checked"),
			"enableSlideTransition": $("#slide_tran").is(":checked"),
			"includeBackground": $("#bk").is(":checked"),
			"alwaysontop": pin
		};
		if (!fs.existsSync(configPath)) {
			configPath = appDataPath + "/" + configFile;
		}
		
		if (saveOnly || !fs.existsSync(configPath)) {
			fs.writeFileSync(configPath, JSON.stringify(cacheData));
			return;
		}

		$.getJSON(configPath, function(json) {
			if (
				(json.showCheckerboard && !cacheData.showCheckerboard) ||
				(!json.showCheckerboard && cacheData.showCheckerboard)
			) {
				$('#trans_checker').trigger("click");
			}

			if (
				(json.enableSlideTransition && !cacheData.enableSlideTransition) ||
				(!json.enableSlideTransition && cacheData.enableSlideTransition)
			) {
				$('#slide_tran').trigger("click");
			}

			if (
				(json.includeBackground && !cacheData.includeBackground) ||
				(!json.includeBackground && cacheData.includeBackground)
			) {
				$('#bk').trigger("click");
			}

			if (json.alwaysontop) {
				ipc.send('remote', { name: "onTop" });
				$("#pin").attr("src", "./img/pin_green.png");
				pin = true;
			} else {
				ipc.send('remote', { name: "onTopOff" });
				$("#pin").attr("src", "./img/pin_grey.png");
				pin = false;
			}
		});
	}

	function cleanupForExit() {
		ipc.sendSync("require", { lib: "ffi", func: "destroy", args: null });
		cleanupForTemp();
		reflectCache(true);
		ipc.send('remote', { name: "exit" });
	}

	function registerEvents() {
		ipc.on('remote' , function(event, data){
			switch (data.msg) {
				case "exit":
					cleanupForExit();
					break;
				case "reload":
					reflectConfig();
					break;
				case "stdin_write_newline":
					res.stdin.write("\n");
					break;
				case "gotoPrev":
					handleHook("prev");
					break;
				case "gotoNext":
					handleHook("next");
					break;
				case "update_trn":
					handleHook("tran");
					break;
				case "update_black":
					handleHook("black");
					break;
				case "update_white":
					handleHook("white");
					break;
			}
			return;
		});

		$('#closeImg').click(function() {
			cleanupForExit();
		});

		$('#bk').click(function() {
			let newVbsContent;
			let vbsDir = tmpDir + '/wb.vbs';
			let file = tmpDir + "/Slide.png";
			if ($("#bk").is(":checked")) {
				newVbsContent = vbsBg;
				try {
					fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
				} catch(e) {
					alertMsg(getLangRsc("ui-slideshow/failed-to-access-tempdir", configData.lang));
					return;
				}
			} else {
				newVbsContent = vbsNoBg;
				try {
					fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
				} catch(e) {
					alertMsg(getLangRsc("ui-slideshow/failed-to-access-tempdir", configData.lang));
					return;
				}
			}
			res.stdin.pause();
			res.kill();
			res = null;
			if (fs.existsSync(vbsDir)) {
				let resX = 0;
				let resY = 0;
				if (customSlideX == 0 || customSlideY == 0 || !/\S/.test(customSlideX) || !/\S/.test(customSlideY)) {
					resX = 0;
					resY = 0;
				} else {
					resX = customSlideX;
					resY = customSlideY;
				}
				res = spawn( 'cscript.exe', [ vbsDir, tmpDir, resX, resY, "//NOLOGO", '' ] );
				res.stdout.on('data', function(data) {
					sendNDI(file, data);
				});
			} else {
				alertMsg(getLangRsc("ui-slideshow/failed-to-parse-presentation", configData.lang));
				return;
			}
		});

		$('#trans_checker').click(function() {
			if ($("#trans_checker").is(":checked")) {
				$("#slidePreview").css('background-image', "url('./img/trans_slide.png')");
			} else {
				$("#slidePreview").css('background-image', "url('./img/null_slide.png')");
			}
		});
		
		$('#pin').click(function() {
			if (pin) {
				ipc.send('remote', { name: "onTopOff" });
				$("#pin").attr("src", "./img/pin_grey.png");
				pin = false;
			} else {
				ipc.send('remote', { name: "onTop" });
				$("#pin").attr("src", "./img/pin_green.png");
				pin = true;
			}
		});
		
		$('#config').click(function() {
			ipc.send('remote', { name: "showConfig" });
		});
	}

	init();
	registerEvents();
});
