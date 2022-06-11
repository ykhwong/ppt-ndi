$(document).ready(function() {
	const { config } = require('process');
	const remote = require('@electron/remote');
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const cscript = require('./js/cscript').script.classic;
	const runtimeUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe";
	const vbsBg = cscript.vbsBg;
	const vbsNoBg = cscript.vbsNoBg;
	const vbsQuickEdit = cscript.vbsQuickEdit;
	const appDataPath = (process.env.APPDATA || (process.platform === 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + "/.local/share")) + "/PPT-NDI";
	let ffi;
	let lib;
	let maxSlideNum = 0;
	let prevSlide = 1;
	let currentSlide = 1;
	let currentWindow = remote.getCurrentWindow();
	let slideWidth = 0;
	let slideHeight = 0;
	let customSlideX = 0;
	let customSlideY = 0;
	let spawnpid = 0;
	let belowImgWidth = 180;
	let belowImgHeight = 100;
	let hiddenSlides = [];
	let advanceSlides = {};
	let slideEffects = {};
	let advanceTimeout = null;
	let configData = {};
	let blkBool = false;
	let whtBool = false;
	let trnBool = false;
	let mustStop = false;
	let isLoaded = false;
	let isCancelTriggered = false;
	let numTypBuf = "";
	let tmpDir = "";
	let preTmpDir = "";
	let pptPath = "";
	let modePostFix = "";
	let modeUseBg = false;
	let mode1options = "";
	let mode2options = "";
	let renderMode = "";
	let pptTimestamp = 0;
	let repo;
	let slideTranTimers = [];
	let multipleMonitors;
	let loadBackgroundInit = false;
	let glLayout = null;

	function prepare() {
		try {
			process.chdir(remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, ""));
		} catch(e) {
		}

		$.ajaxSetup({
			async: false
		});
		reflectConfig();
		ffi = ipc.sendSync("require", { lib: "ffi", func: null, args: null });
		if ( ffi === -1 ) {		
			let librariesFound = false;
			if (process.platform === 'win32') {
				alertMsg(getLangRsc("ui-classic/dll-init-failed", configData.lang));
				if (fs.existsSync("PPTNDI.DLL") && fs.existsSync("Processing.NDI.Lib.x64.dll")) {
					librariesFound = true;
				}
			} else if (process.platform === 'darwin') {
				alertMsg(getLangRsc("ui-classic/dylib-init-failed", configData.lang));
				if (fs.existsSync("PPTNDI.dylib")) {
					librariesFound = true;
				}
			} else {
				alertMsg(getLangRsc("ui-classic/dll-init-failed", configData.lang));
				librariesFound = false;
			}

			if (librariesFound && process.platform === 'win32') {
				const execRuntime = require('child_process').execSync;
				execRuntime("start " + runtimeUrl, (error, stdout, stderr) => { 
					callback(stdout); 
				});
			}
			ipc.send('remote', { name: "exit" });
		}

		lib = ipc.sendSync("require", { lib: "ffi", func: "init", args: null });
		if (lib === 1) {
			alertMsg(getLangRsc("ui-classic/failed-to-create-listening-svr", configData.lang));
			ipc.send('remote', { name: "exit" });
			return;
		}

		$("#resWidth").val("0");
		$("#resHeight").val("0");
	}

	function alertMsg(myMsg) {
		const {dialog} = require('@electron/remote');
		let options;
		options = {
			type: 'info',
			message: "" + myMsg,
			buttons: ["OK"]
		};
		dialog.showMessageBoxSync(currentWindow, options);
	}

	function setLangRsc() {
		setLangRscDiv("#edit_pptx", "ui-classic/edit_pptx", false, configData.lang);
		//setLangRscDiv("#reload", "ui-classic/reload", false, configData.lang);
		//setLangRscDiv("#config", "ui-classic/config", false, configData.lang);
		setLangRscDiv("#show-checkerboard", "ui-classic/show-checkerboard", false, configData.lang);
		setLangRscDiv("#enable-slide-transition-effect", "ui-classic/enable-slide-transition-effect", false, configData.lang);
		setLangRscDiv("#include-background", "ui-classic/include-background", false, configData.lang);
		setLangRscDiv("#resolution", "ui-classic/resolution", false, configData.lang);
		setLangRscDiv("#resDefault", "ui-classic/resDefault", false, configData.lang);
		setLangRscDiv("#resCustom", "ui-classic/resCustom", false, configData.lang);
		setLangRscDiv("#setRes", "ui-classic/setRes", false, configData.lang);
		setLangRscDiv("#monitorControlText", "ui-classic/monitorControlText", false, configData.lang);
		setLangRscDiv("#monitorAlphaText", "ui-classic/monitorAlphaText", false, configData.lang);
		setLangRscDiv("#setMonitor", "ui-classic/setMonitor", false, configData.lang);
		setLangRscDiv("#monitorText", "ui-classic/monitor", false, configData.lang);
		//$("#currentSlideText").attr("data-img-label", getLangRsc("ui-classic/currentSlideText", configData.lang));
		//$("#nextSlideText").attr("data-img-label", getLangRsc("ui-classic/nextSlideText", configData.lang));
		setLangRscDiv("#loadingTxt", "ui-classic/loadingTxt", false, configData.lang);
		setLangRscDiv("#cancel", "ui-classic/cancel", false, configData.lang);
	}

	function stopSlideTransition() {
		for (let pp=2; pp<=9; pp++) {
			clearTimeout(slideTranTimers[pp]);
		}
		mustStop = true;
	}

	function hideCancelBox() {
		$("#fullblack, .cancelBox").hide();
	}

	function createNullSlide() {
		const now = new Date().getTime();
		const PNG = require('pngjs').PNG;
		let png;
		let buffer;
		function getMeta(url, callback) {
			let img = new Image();
			img.src = url;
			img.onload = function() { callback(this.width, this.height); }
		}
		function createSli(redVal, greenVal, blueVal, alphaVal, fileVal) {
			png = new PNG({
				width: slideWidth,
				height: slideHeight,
				filterType: -1
			});

			for (let y = 0; y < png.height; y++) {
				for (let x = 0; x < png.width; x++) {
					let idx = (png.width * y + x) << 2;
					png.data[idx  ] = redVal;
					png.data[idx+1] = greenVal;
					png.data[idx+2] = blueVal;
					png.data[idx+3] = alphaVal;
				}
			}
			buffer = PNG.sync.write(png);
			fs.writeFileSync(tmpDir + fileVal, buffer);
		}
		getMeta(
			tmpDir + "/Slide1.png" + "?" + now,
			function(width, height) { 
				slideWidth = width;
				slideHeight = height;
				$("#slide_res").html(slideWidth + " x " + slideHeight);

				createSli(255, 255, 255, 0, "/Slide0.png");
				createSli(255, 255, 255, 255, "/SlideWhite.png");
				createSli(0, 0, 0, 255, "/SlideBlack.png");
			}
		);
	}

	function updateCurNext(curSli, nextSli) {
		$("select").find('option[value="Current"]').data('img-src', curSli);
		$("select").find('option[value="Next"]').data('img-src', nextSli);
		initImgPicker();
		ipc.send("monitor", {
			file: curSli,
			workerinit: loadBackgroundInit,
			mode: modePostFix,
			modeusebg: modeUseBg,
			func: "update"
		});
	}

	function updateScreen(noTran) {
		let curSli, nextSli;
		let nextNum;
		let re, rpc;
		clearTimeout(advanceTimeout);
		if(!repo) {
			return;
		}
		rpc = tmpDir + modePostFix + "/Slide";
		curSli = rpc + currentSlide.toString() + '.png';
		nextNum = currentSlide;
		nextNum++;

		if (nextNum > maxSlideNum) {
			nextNum = 1;
		}
		if (hiddenSlides.length == 0 || maxSlideNum == hiddenSlides.length) {
			nextSli = rpc + nextNum.toString() + '.png';
		} else {
			let cnts = 0;
			while (1) {
				if (!hiddenSlides.includes(nextNum + cnts)) {
					nextNum += cnts;
					nextSli = rpc + nextNum.toString() + '.png';
					break;
				}
				cnts++;
			}
		}

		if (
		    $("#use_slide_transition").is(":checked") && !(whtBool || trnBool || blkBool) && !noTran &&
		    ! (Object.entries(slideEffects).length === 0 && slideEffects.constructor === Object) &&
		    slideEffects[currentSlide.toString()].effectName !== "0"
		) {
			let duration = slideEffects[currentSlide.toString()].duration;
			const prevSli = rpc + prevSlide.toString() + '.png';
			const transLvl=9;
			try {
				for (let i=2; i<=transLvl; i++) {
					fs.unlinkSync(tmpDir + modePostFix + "/t" + i.toString() + ".png");
				}
			} catch(e) {
			}
			function sendSlides(i) {
				if (mustStop) {
					updateCurNext(curSli, nextSli);
					return;
				}
				function setLast() {
					if (mustStop) {
						updateCurNext(curSli, nextSli);
						return;
					}
					slideTranTimers[10] = setTimeout(function() {
						ipc.sendSync("require", {
							lib: "ffi",
							func: "send",
							args: [
								tmpDir + modePostFix + "/Slide" + currentSlide.toString() + ".png",
								false
							]
						});
						updateCurNext(curSli, nextSli);
					}, 10 * parseFloat(duration) * 50);
				}
				slideTranTimers[i] = setTimeout(function() {
					ipc.sendSync("require", {
						lib: "ffi",
						func: "send",
						args: [
							tmpDir + modePostFix + "/t" + i.toString() + ".png",
							true
						]
					});
					if ( i % 2 === 0 ) {
						let now = new Date().getTime();
						$("img.image_picker_image:first").attr("src", tmpDir + modePostFix + "/t" + i.toString() + ".png?" + now);
					}
				}, i * parseFloat(duration) * 50);
				if (i === transLvl) {
					setLast();
				}
			}

			function doTrans() {
				const mergeImages = require('merge-images');
				stopSlideTransition();
				mustStop = false;

				for (let i=2; i<=transLvl; i++) {	
					mergeImages([
						{ src: prevSli, opacity: 1 - (0.1 * i) },
						{ src: curSli, opacity: 0.1 * i }
					])
					.then(b64 => {
						let b64data = b64.replace(/^data:image\/png;base64,/, "");
						try {
							fs.writeFileSync(tmpDir + modePostFix + "/t" + i.toString() + ".png", b64data, 'base64');
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

		} else {
			stopSlideTransition();
			updateCurNext(curSli, nextSli);
			ipc.sendSync("require", {
				lib: "ffi",
				func: "send",
				args: [
					curSli,
					false
				]
			});
		}
		$("#slide_cnt").html("SLIDE " + currentSlide + " / " + maxSlideNum);
		blkBool = false;
		whtBool = false;
		trnBool = false;
		if (/^(\d|\.)+$/.test(advanceSlides[currentSlide])) {
			advanceTimeout = setTimeout(function() {
				gotoNext();
			}, parseFloat(advanceSlides[currentSlide]) * 1000);
		}
	}

	function askReloadFile(ele, msg, description) {
		const {dialog} = require('@electron/remote');
		let options;
		let response;
		let defaultMsg = getLangRsc("ui-classic/do-you-want-to-continue", configData.lang);
		let defaultDetail = getLangRsc("ui-classic/changes-require-reload", configData.lang);
		
		options = {
			type: 'question',
			buttons: ['Yes', 'No'],
			defaultId: 1,
			message: "" + (/\S/.test(msg)?msg:defaultMsg),
			detail: (/\S/.test(description)?description:defaultDetail)
		};

		response = dialog.showMessageBoxSync(currentWindow, options);
		if (response === 0) { // Yes
			loadPPTX(pptPath);
		} else { // No
			if (ele !== undefined && /\S/.test(ele)) {
				$(ele).prop("checked", !$(ele).prop("checked"));
			}
		}
	}

	function registerEvents() {
		$("#monitor_trans").click(function() {
			if (!isLoaded) {
				return;
			}

			if (!loadBackgroundInit) {
				$(this).prop("checked", !$(this).prop("checked"));
			}
		});

		$("#with_background").click(function() {
			if (!isLoaded) {
				return;
			}

			if (loadBackgroundInit) {
				if (maxSlideNum > 0) {
					if (modePostFix == "") {
						modePostFix = "/mode2";
						$("#slides_grp").html(mode2options);
					} else {
						modePostFix = "";
						$("#slides_grp").html(mode1options);
					}
					modeUseBg != modeUseBg;
					selectSlide(currentSlide.toString());
				}
			} else {
				$(this).prop("checked", !$(this).prop("checked"));
			}
		});

		ipc.on('renderer' , function(event, data){
			switch (data.name) {
				case "notifyError":
					cleanupForTemp(false);
					tmpDir = preTmpDir;
					// Error
					hideCancelBox();
					break;
				case "notifyLoaded":
					// Done
					loadPPTX_PostProcess(data.pptFile);
					hideCancelBox();
					break;
				case "notifyCanceled":
					// Canceled
					cleanupForTemp(false);
					tmpDir = preTmpDir;
					hideCancelBox();		
					break;
			}
		});
		$("#load_pptx").click(function() {
			const {dialog} = require('@electron/remote');
			$("#fullblack").show();

			dialog.showOpenDialog(currentWindow,{
				properties: ['openFile'],
				filters: [
					{name: getLangRsc("ui-classic/open-file-ppt-presentation", configData.lang), extensions: ['pptx', 'ppt']},
					{name: getLangRsc("ui-classic/open-file-all-files", configData.lang), extensions: ['*']}
				]
			}).then(result => {
				loadPPTX(result.filePaths[0], 0, 0);
			}).catch(err => {
				$("#fullblack").hide();
			});
		});

		$("#reload").click(function() {
			if (maxSlideNum > 0) {
				askReloadFile(null, "", "");
			}
		});

		$("#edit_pptx").click(function() {
			if (pptPath !== "") {
				const { exec } = require('child_process');
				exec('"' + pptPath + '"', (err, stdout, stderr) => {
					if (err) {
						return;
					}
				});
			}
		});

		$('#prev').click(function() {
			gotoPrev();
		});

		$('#next').click(function() {
			gotoNext();
		});

		$('#blk').click(function() {
			updateBlkWhtTrn("black");
		});

		$('#wht').click(function() {
			updateBlkWhtTrn("white");
		});

		$('#trn').click(function() {
			updateBlkWhtTrn("trn");
		});

		$(document).keydown(function(e) {
			let realNum = 0;
			if (e.ctrlKey || e.shiftKey || e.altKey || e.metaKey) {
				numTypBuf = "";
				if (e.ctrlKey) {
					e.preventDefault();
					e.stopPropagation();
				}
				return;
			}
			if ($(":text").is(":focus")) {
				return;
			}
			$("#below").trigger('click');
			if(e.which >= 48 && e.which <= 57) {
				// 0 through 9
				realNum = e.which - 48;
				numTypBuf += realNum.toString();
			} else if (e.which >= 96 && e.which <= 105) {
				// 0 through 9 (keypad)
				realNum = e.which - 96;
				numTypBuf += realNum.toString();
			} else if (e.which === 13) {
				// Enter
				if (numTypBuf == "") {
					gotoNext();
				} else {
					realNum = parseInt(numTypBuf, 10);
					selectSlide(realNum);
				}
				numTypBuf = "";
			} else if (e.which === 32 || e.which === 39 || e.which === 40 || e.which === 78 || e.which === 34) {
				// Spacebar, right arrow, down, N or page down
				numTypBuf = "";
				gotoNext();
			} else if(e.which === 37 || e.which === 8 || e.which === 38 || e.which === 80 || e.which === 33) {
				// Left arrow, backspace, up, P or page up
				numTypBuf = "";
				gotoPrev();
			} else if(e.which === 36) {
				// Home
				numTypBuf = "";
				if (hiddenSlides.length === 0 || maxSlideNum === hiddenSlides.length) {
					selectSlide('1');
				} else {
					for (i = 1; i <= maxSlideNum; i++) {
						if (!hiddenSlides.includes(i)) {
							selectSlide(i.toString());
							break;
						}
					}
				}
			} else if(e.which === 35) {
				// End
				numTypBuf = "";
				if (hiddenSlides.length === 0 || maxSlideNum === hiddenSlides.length) {
					selectSlide(maxSlideNum.toString());
				} else {
					for (i = maxSlideNum; i >= 1; i--) {
						if (!hiddenSlides.includes(i)) {
							selectSlide(i.toString());
							break;
						}
					}
				}

			} else if(e.which === 66) {
				// B
				numTypBuf = "";
				updateBlkWhtTrn("black");
			} else if(e.which === 84) {
				// T
				numTypBuf = "";
				updateBlkWhtTrn("trn");
			} else if(e.which === 87) {
				// W
				numTypBuf = "";
				updateBlkWhtTrn("white");
			} else if(e.which === 189 || e.which === 109) {
				// -
				makePreviewSmaller();
			} else if(e.which === 187 || e.which === 107) {
				// +
				makePreviewBigger();
			}
		});

		$('#smaller').click(function() {
			makePreviewSmaller();
		});

		$('#bigger').click(function() {
			makePreviewBigger();
		});

		$('.button, .checkbox').keydown(function(e){
			if (e.which == 13 || e.which == 32) {
				// Enter or spacebar
				e.preventDefault();
				e.stopPropagation();
				gotoNext();
			}
		});

		ipc.on('remote' , function(event, data){
			switch (data.msg) {
				case "exit":
					cleanupForExit();
					break;
				case "reload":
					reflectConfig();
					break;
				case "focused":
					let stats;
					let tmpPptTimestamp = 0;
					if (pptTimestamp === 0) {
						return;
					}
					try {
						stats = fs.statSync(pptPath);
						tmpPptTimestamp = stats.mtimeMs;
						if (pptTimestamp === tmpPptTimestamp) {
						} else {
							pptTimestamp = tmpPptTimestamp;
							askReloadFile("", getLangRsc("ui-classic/ask-reload", configData.lang), "");
						}
					} catch(e) {
					}
					break;
				case "blurred":
					// we don't care here
					break;
				case "gotoPrev":
					gotoPrev();
					break;
				case "gotoNext":
					gotoNext();
					break;
				case "update_trn":
					updateBlkWhtTrn("trn");
					break;
				case "update_black":
					updateBlkWhtTrn("black");
					break;
				case "update_white":
					updateBlkWhtTrn("white");
					break;
				default:
					break;
			}
			return;
		});

		$('#minimize').click(function() {
			remote.BrowserWindow.getFocusedWindow().minimize();
		});

		$('#max_restore').click(function() {
			if(currentWindow.isMaximized()) {
				remote.BrowserWindow.getFocusedWindow().unmaximize();
			} else {
				remote.BrowserWindow.getFocusedWindow().maximize();
			}
		});

		$('#cancel').click(function() {
			isCancelTriggered = true;
			cancelLoad();
		});

		$('#trans_checker').click(function() {
			if ($("#trans_checker").is(":checked")) {
				$(".right img").css('background-image', "url('./img/trans_slide.png')");
			} else {
				$(".right img").css('background-image', "url('./img/null_slide.png')");
			}
		});

		currentWindow.on('maximize', function (){
			$("#max_restore").attr("src", "./img/restore.png");
		});

		currentWindow.on('unmaximize', function (){
			$("#max_restore").attr("src", "./img/max.png");
		});
		
		$('#exit').click(function() {
			cleanupForExit();
		});

		document.addEventListener('dragover',function(event){
			event.preventDefault();
			return false;
		},false);
		
		document.addEventListener('drop',function(event){
			event.preventDefault();
			return false;
		},false);

		window.addEventListener("keydown", function(e) {
			if([32, 37, 38, 39, 40].indexOf(e.keyCode) > -1) {
				if (!$(":text").is(":focus")) {
					e.preventDefault();
				}
			}
		}, false);

		$("select").change(function() {
			if (repo == null) {
				repo = $(this);
			}
		});

		$("#resWidth, #resHeight").click(function() {
			$(this).val("");
		});

		$("#setRes").click(function() {
			let resX = $("#resWidth").val();
			let resY = $("#resHeight").val();
			if (/^\d+$/.test(resX) && /^\d+$/.test(resY)) {
				customSlideX = resX;
				customSlideY = resY;
				if (maxSlideNum > 0) {
					// slideWidth : slideHeight = customSlideX : customSlideY
					if (slideHeight*customSlideX/slideWidth !== parseInt(customSlideY, 10)) {
						alertMsg(getLangRsc("ui-classic/original-aspect-ratio-not-match", configData.lang));
					}
					askReloadFile(null, "", "");
				}
			}
		});

		$("#setMonitor").click(function() {
			if (!isLoaded) {
				return;
			}

			if (!loadBackgroundInit) {
				$(this).prop("checked", !$(this).prop("checked"));
				return;
			}

			let idx = $("#monitorList").prop('selectedIndex');
			if (idx === 0) {
				disableMonitor();
				return;
			}
			assignMonitor($("#monitorList").prop('selectedIndex'));
			$("#monitor_trans").is(":checked") ? enableMonitorTransparent() : disableMonitorTransparent();
			enableMonitor();
		});

		$('#monitor_trans').click(function() {
			if ($("#monitor_trans").is(":checked")) {
				enableMonitorTransparent();
			} else {
				disableMonitorTransparent();
			}
		});

		$("#listRes").click(function() {
			$(".resText").hide();
			$("#listResList").val("0x0");
			$("#listResList").show();
		});

		$("#listResList").change(function() {
			let resVal = $(this).val();
			if (resVal === "custom") {
				$("#listResList").hide();
				$(".resText").show();
			} else {
				$("#resWidth").val(resVal.replace(/x.*/, ""));
				$("#resHeight").val(resVal.replace(/.*x/, ""));
			}
		});

		$("#config").click(function() {
			ipc.send('remote', { name: "showConfig" });
		});

		$(".resText").hide();
		$("#listResList").show();

		$(window).mousedown(function(event) {
			if ( process.platform === 'darwin' ) {
				return;
			}
			if ( event.which === 3 ) {
				if ( /image_picker_image/.test($(event.target).attr('class')) ) {
					let lSlideNo = $(event.target).attr('src').replace(/.*Slide(\d+)\.png\s*$/i, "$1");

					if ( /^\d+$/.test(lSlideNo) ) {
						const { Menu, MenuItem } = require('@electron/remote');
						const menu = new Menu();
						menu.append(new MenuItem ({
							label: getLangRsc("ui-classic/quick-edit", configData.lang),
								click() {
									let vbsDir;
									let file = pptPath;
									let tmpDir2 = tmpDir + "/mode2";
									vbsDir = tmpDir2 + '/vbQuickEdit.vbs';

									try {
										fs.writeFileSync(vbsDir, vbsQuickEdit, 'utf-8');
									} catch(e) {
										return;
									}
									const spawn = require( 'child_process' ).spawn;
									spawn( 'cscript.exe', [ "//NOLOGO", "//E:jscript", vbsDir, file, lSlideNo, '' ] );
								}
							}));
						if ( /CURRENT|NEXT/.test($(event.target).parent().text()) ) {
							menu.append(new MenuItem ({
								label: ($("#trans_checker").is(":checked")) ? getLangRsc("ui-classic/hide-checkerboard", configData.lang) : getLangRsc("ui-classic/show-checkerboard", configData.lang),
									click() {
										$('#trans_checker').trigger("click");
									}
								}));
						}
						menu.popup();
					} else {
						const { Menu, MenuItem } = require('@electron/remote');
						const menu = new Menu();
						if ( /CURRENT|NEXT/.test($(event.target).parent().text()) ) {
							menu.append(new MenuItem ({
								label: ($("#trans_checker").is(":checked")) ? getLangRsc("ui-classic/hide-checkerboard", configData.lang) : getLangRsc("ui-classic/show-checkerboard", configData.lang),
									click() {
										$('#trans_checker').trigger("click");
									}
								}));
						}
						menu.popup();
					}
				}
			}
		});
	}

	function initImgPicker() {
		$(".right select").imagepicker({
			hide_select: true,
			show_label: true,
			selected:function(select, picker_option, event) {
				prevSlide = currentSlide;
				currentSlide=$('.selected').text();
				updateScreen(false);
			}
		});
		if ($("#trans_checker").is(":checked")) {
			$(".right img").css('background-image', "url('./img/trans_slide.png')");
		}
		$("img.image_picker_image:first, img.image_picker_image:eq(1)").click(function() {
			gotoNext();
		});
		$(window).trigger('resize');
		applyPreviewSize();
	}

	function cancelLoad() {
		if (renderMode !== "Internal") {
			const kill  = require('tree-kill');
			kill(spawnpid);
		} else {
			ipc.send("renderer", {
				func: "cancel"
			});
		}
	}

	function loadBackgroundWorker() {
		let vbsDir, res;
		let re = new RegExp("\\.(ppt|pptx)\$", "i");
		let resX = 0;
		let resY = 0;
		let file = pptPath;
		loadBackgroundInit = false;
		modePostFix = "";
		$("#with_bg_slider, #monitor_trans_switch, #setMonitor").css("filter", "sepia(1)");
		if (re.exec(file)) {
			let tmpDir2 = tmpDir + "/mode2";
			let newVbsContent;
			const spawn = require( 'child_process' ).spawn;
			if (fs.existsSync(tmpDir2)) {
				fs.removeSync(tmpDir2);
			}
			if (!fs.existsSync(tmpDir2)) {
				fs.mkdirSync(tmpDir2, { recursive: true });
			}
			vbsDir = tmpDir2 + '/wb.vbs';
			if ($("#with_background").is(":checked")) {
				newVbsContent = vbsNoBg;
				modeUseBg = false;
			} else {
				newVbsContent = vbsBg;
				modeUseBg = true;
			}

			try {
				fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
			} catch(e) {
				loadBackgroundInit = false;
				$("#with_bg_slider, #monitor_trans_switch, #setMonitor").css("filter", "sepia(0)");
				return;
			}

			if (customSlideX == 0 || customSlideY == 0 || !/\S/.test(customSlideX) || !/\S/.test(customSlideY)) {
				resX = 0;
				resY = 0;
			} else {
				resX = customSlideX;
				resY = customSlideY;
			}
			res = spawn( 'cscript.exe', [ "//NOLOGO", "//E:jscript", vbsDir, file, tmpDir2, resX, resY, '' ] );
			res.stderr.on('data', (data) => {
				loadBackgroundInit = false;
				$("#with_bg_slider, #monitor_trans_switch, #setMonitor").css("filter", "sepia(0)");
				return;
			});
			res.on('close', (code) => {
				let fileArr = [];
				let options;
				for (let i=1; i<=maxSlideNum; i++) {
					fileArr.push(i.toString());
				}
				fileArr.sort((a, b) => a - b).forEach(file2 => {
					let rpc = file2;
					let isHidden = false;
					options += '<option data-img-label="' + rpc + '"';

					for (i = 0, len = hiddenSlides.length; i < len; i++) { 
						let num = hiddenSlides[i];
						if (/^\d+$/.test(num)) {
							if (num == parseInt(rpc, 10)) {
								options += ' data-img-class="hiddenSlide" ';
								isHidden = true;
								break;
							}
						}
					}
					if (!isHidden && slideEffects[rpc] && ( slideEffects[rpc].effectName !== "0" )) {
						options += ' data-img-class="transSlide" ';
					}

					options += ' data-img-src="' + tmpDir2 + '/Slide' + rpc + '.png" value="' + rpc + '">' + "\n";
					$("select").find('option[value="Current"]').prop('img-src', tmpDir + "/Slide1.png");
					if (!fs.existsSync(tmpDir + "/Slide2.png")) {
						$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide1.png");
					} else {
						$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide2.png");
					}
				});
				mode2options = options;

				loadBackgroundInit = true;
				$("#with_bg_slider, #monitor_trans_switch, #setMonitor").css("filter", "sepia(0)");
			});
		}
	}

	function setTmpDir() {
		if (process.platform === 'darwin') {
			tmpDir = process.env.TMPDIR + '/ppt_ndi';
		} else if (process.platform === 'win32') {
			tmpDir = process.env.PROGRAMDATA + '/PPT-NDI/temp';
		} else {
			tmpDir = '/tmp/ppt_ndi';
		}
	}

	function loadPPTX_Renderer_Internal(file) {
		let now = new Date().getTime();
		isCancelTriggered = false;
		preTmpDir = tmpDir;
		setTmpDir();
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir, { recursive: true });
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir, { recursive: true });
		let opts = {
			file: file,
			outDir : tmpDir,
			resX : customSlideX,
			resY : customSlideY
		}
		$(".cancelBox").show();

		ipc.send("renderer", {
			func: "load",
			options: opts
		});
	}

	function loadPPTX_PostProcess(file) {
		let fileArr = [];
		let options = "";
		let newMaxSlideNum = 0;
		let stats;
		let path = require("path");
		if (tmpDir === "") {
			return;
		}
		fs.readdirSync(tmpDir).forEach(file2 => {
			re = new RegExp("^Slide(\\d+)\\.png\$", "i");
			if (re.exec(file2)) {
				let rpc = file2.replace(re, "\$1");
				fileArr.push(rpc);
				newMaxSlideNum++;
			}
		});
		if (isCancelTriggered) { hideCancelBox(); return; }
		if (fileArr === undefined || fileArr.length == 0) {
			maxSlideNum = 0;
			cleanupForTemp(false);
			tmpDir = preTmpDir;
			alertMsg(getLangRsc("ui-classic/ppt-not-loaded", configData.lang));
			hideCancelBox();
			return;
		}
		hiddenSlides = [];
		if (fs.existsSync(tmpDir + "/hidden.dat")) {
			const hs = fs.readFileSync(tmpDir + "/hidden.dat", { encoding: 'utf8' });
			hiddenSlides = hs.split("\n");
		}

		advanceSlides = {};
		if (fs.existsSync(tmpDir + "/advance.dat")) {
			const as = fs.readFileSync(tmpDir + "/advance.dat", { encoding: 'utf8' });
			const tmpAdvanceSlides = as.split(/\r\n|\n/);
			for (let i=0; i < tmpAdvanceSlides.length; i++) {
				let sNum = tmpAdvanceSlides[i].split("\t")[0];
				let sSec = tmpAdvanceSlides[i].split("\t")[1];
				advanceSlides[sNum] = sSec;
			}
		}

		if (isCancelTriggered) return;
		hiddenSlides = hiddenSlides.filter(n => n);
		for (i = 0, len = hiddenSlides.length; i < len; i++) { 
			hiddenSlides[i] = parseInt(hiddenSlides[i], 10);
		}

		slideEffects = {};
		if (fs.existsSync(tmpDir + "/slideEffect.dat")) {
			const hs = fs.readFileSync(tmpDir + "/slideEffect.dat", { encoding: 'utf8' });
			const lines = hs.split(/(\r|\n)+/);
			for (i = 0; i < lines.length; i++) {
				let ls = lines[i].split(",");
				let obj = {
					"effectName" : ls[1],
					"duration" : ls[2]
				};
				slideEffects[ls[0].toString()] = obj;
			}
		}
		if (isCancelTriggered) { hideCancelBox(); return; }
		fileArr.sort((a, b) => a - b).forEach(file2 => {
			let rpc = file2;
			let isHidden = false;
			options += '<option data-img-label="' + rpc + '"';

			for (i = 0, len = hiddenSlides.length; i < len; i++) { 
				let num = hiddenSlides[i];
				if (/^\d+$/.test(num)) {
					if (num == parseInt(rpc, 10)) {
						options += ' data-img-class="hiddenSlide" ';
						isHidden = true;
						break;
					}
				}
			}
			if (!isHidden && slideEffects[rpc] && ( slideEffects[rpc].effectName !== "0" )) {
				options += ' data-img-class="transSlide" ';
			}

			options += ' data-img-src="' + tmpDir + '/Slide' + rpc + '.png" value="' + rpc + '">' + "\n";
			$("select").find('option[value="Current"]').prop('img-src', tmpDir + "/Slide1.png");
			if (!fs.existsSync(tmpDir + "/Slide2.png")) {
				$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide1.png");
			} else {
				$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide2.png");
			}
		});
		$("#slides_grp").html(options);
		mode1options = options;
		hideCancelBox();
		maxSlideNum = newMaxSlideNum;
		createNullSlide();

		if (configData.startWithTheFirstSlideSelected === true) {
			if (hiddenSlides.length === 0 || maxSlideNum === hiddenSlides.length) {
				selectSlide('1');
			} else {
				for (i = 1; i <= maxSlideNum; i++) {
					if (!hiddenSlides.includes(i)) {
						selectSlide(i.toString());
						break;
					}
				}
			}
		} else {
			let selectedDiv = "ul.thumbnails.image_picker_selector li .thumbnail.selected";
			let tmpSrc = $("img.image_picker_image:first").attr('src');
			initImgPicker();
			$("img.image_picker_image:first").attr('src', tmpSrc);
			$(selectedDiv).css("background", "rgb(0, 0, 0, 0)");
			currentSlide = 0;
			$("img.image_picker_image:eq(1)").attr("src", "./img/null_slide.png");
			$("#below .thumbnail:first").click(function() {
				selectSlide('1');
				$(this).off('click');
			});
		}
		
		if (isLoaded) {
			cleanupForTemp(true);
		}
		isLoaded = true;
		pptPath = file;
		stats = fs.statSync(pptPath);
		pptTimestamp = stats.mtimeMs;
		$("#ppt_filename").html(path.basename(pptPath));
		blkBool = false;
		whtBool = false;
		trnBool = false;
		if (renderMode !== "Internal") {
			loadBackgroundWorker();
		} else {
			$("#with_bg_slider, #monitor_trans_switch, #setMonitor").css("filter", "sepia(1)");
		}
	}

	function loadPPTX_Renderer_PPT(file) {
		isCancelTriggered = false;
		let vbsDir, res;
		let fileArr = [];
		let options = "";
		let resX = 0;
		let resY = 0;
		let now = new Date().getTime();
		let newVbsContent;
		const spawn = require( 'child_process' ).spawn;
		preTmpDir = tmpDir;
		setTmpDir();
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir, { recursive: true });
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir, { recursive: true });
		vbsDir = tmpDir + '/wb.vbs';

		if ($("#with_background").is(":checked")) {
			newVbsContent = vbsBg;
		} else {
			newVbsContent = vbsNoBg;
		}

		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			cleanupForTemp(false);
			tmpDir = preTmpDir;
			alertMsg(getLangRsc("ui-classic/failed-to-access-tempdir", configData.lang));
			hideCancelBox();
			return;
		}

		if (customSlideX == 0 || customSlideY == 0 || !/\S/.test(customSlideX) || !/\S/.test(customSlideY)) {
			resX = 0;
			resY = 0;
		} else {
			resX = customSlideX;
			resY = customSlideY;
		}
		
		res = spawn( 'cscript.exe', [ "//NOLOGO", "//E:jscript", vbsDir, file, tmpDir, resX, resY, '' ] );
		spawnpid = res.pid;

		$(".cancelBox").show();
		res.stderr.on('data', (data) => {
			let myMsg = getLangRsc("ui-classic/failed-to-parse-presentation", configData.lang);
			maxSlideNum = 0;
			cleanupForTemp(false);
			tmpDir = preTmpDir;
			if (!fs.existsSync(file)) {
				alertMsg(myMsg + getLangRsc("ui-classic/file-moved-or-deleted", configData.lang));
			} else if (maxSlideNum > 0) {
				alertMsg(myMsg + getLangRsc("ui-classic/check-the-config", configData.lang));
			} else {
				alertMsg(myMsg + getLangRsc("ui-classic/make-sure-ppt-installed", configData.lang));
			}
			hideCancelBox();
			return;
		});
		res.on('close', (code) => {
			loadPPTX_PostProcess(file);
		});
	}

	function loadPPTX(file) {
		if (file === undefined) {
			hideCancelBox();
			return false;
		}
		if (! /\.(ppt|pptx)$/i.test(file)) {
			if (/\S/.test(file)) {
				alertMsg(getLangRsc("ui-classic/only-allowed-filename", configData.lang));
			}
			hideCancelBox();
			return false;
		}

		$("#fullblack").show();
		renderMode = configData.renderer;
		if (configData.renderer === "Microsoft PowerPoint") {
			loadPPTX_Renderer_PPT(file);
		} else if (configData.renderer === "Internal") {
			loadPPTX_Renderer_Internal(file);
		} else {
			hideCancelBox();
		}
	}

	function selectSlide(num) {
		blkBool = false;
		whtBool = false;
		trnBool = false;
		if (num == 0) {
			return;
		}
		if ( num > maxSlideNum ) {
			num = maxSlideNum;
		}
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').prop('selected',true);
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').change();
		prevSlide = currentSlide;
		currentSlide = num;

		let selected = $('.selected:eq( 0 )');
		if (selected.length) {
			$("#below").stop().animate(
			{ scrollTop: selected.position().top + $("#below").scrollTop() },
			  500, 'swing', function() {
			  });
		}
		updateScreen(false);
		applyPreviewSize();
	}

	function gotoPrev() {
		let curSli;
		let re;
		if (!repo || maxSlideNum === 0) {
			return;
		}
		curSli = currentSlide;
		if (hiddenSlides.length == 0 || maxSlideNum == hiddenSlides.length) {
			curSli--;
			if (curSli == 0) {
				curSli = maxSlideNum;
			}
		} else {
			while (true) {
				curSli--;
				if (curSli == 0) {
					curSli = maxSlideNum;
				}
				if (!hiddenSlides.includes(curSli)) {
					break;
				}
			}
		}
		selectSlide(curSli.toString());
	}

	function gotoNext() {
		let curSli;
		let re;
		if (!repo || maxSlideNum === 0) {
			return;
		}
		curSli = currentSlide;
		if (hiddenSlides.length == 0 || maxSlideNum == hiddenSlides.length) {
			curSli++;
			if (curSli > maxSlideNum) {
				curSli = 1;
			}
		} else {
			while (true) {
				curSli++;
				if (curSli > maxSlideNum) {
					curSli = 1;
				}
				if (!hiddenSlides.includes(curSli)) {
					break;
				}
			}
		}
		selectSlide(curSli.toString());
	}

	function updateBlkWhtTrn(color) {
		if (maxSlideNum === 0) {
			return;
		}
		let dirTo = "";
		switch (color) {
			case "black":
				whtBool = false;
				trnBool = false;
				if (blkBool) {
					blkBool = false;
					updateScreen(true);
					return;
				} else {
					blkBool = true;
					dirTo = tmpDir + "/SlideBlack.png";
					multipleMonitors = ipc.send("monitor", {
						func: "monitorBlack"
					});
				}
				break;
			case "white":
				blkBool = false;
				trnBool = false;
				if (whtBool) {
					whtBool = false;
					updateScreen(true);
					return;
				} else {
					whtBool = true;
					dirTo = tmpDir + "/SlideWhite.png";
					multipleMonitors = ipc.send("monitor", {
						func: "monitorWhite"
					});
				}
				break;
			case "trn":
				blkBool = false;
				whtBool = false;
				if (trnBool) {
					trnBool = false;
					updateScreen(true);
					return;
				} else {
					trnBool = true;
					dirTo = tmpDir + "/Slide0.png";
					color = "null";
					multipleMonitors = ipc.send("monitor", {
						func: "monitorTrans"
					});
				}
				break;
			default:
				break;
		}

		if (!fs.existsSync(dirTo)) {
			dirTo = __dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/img/" + color + "_slide.png";
		}
		$("img.image_picker_image:first").attr('src', dirTo);

		ipc.sendSync("require", {
			lib: "ffi",
			func: "send",
			args: [
				dirTo,
				false
			]
		});
	}

	function resetPreviewSize() {
		belowImgWidth = 180;
		belowImgHeight = 100;
	}

	function applyPreviewSize() {
		$("#below img").css("width", belowImgWidth + "px");
		$("#below img").css("height", belowImgHeight + "px");
	}

	function makePreviewSmaller() {
		belowImgWidth -= 5;
		belowImgHeight -= 2.7;
		if (belowImgWidth < 0 || belowImgHeight < 0) {
			resetPreviewSize();
		}
		applyPreviewSize();
	}

	function makePreviewBigger() {
		belowImgWidth += 5;
		belowImgHeight += 2.7;
		if (
		$("#below").height() < $("#below .thumbnail:first").height() ||
		$("#below").width() < $("#below .thumbnail:first").width()
		) {
			belowImgWidth -= 5;
			belowImgHeight -= 2.7;
		}
		applyPreviewSize();
	}

	function checkTime(i) {
		if (i < 10) {
			i = "0" + i;
		}
		return i;
	}

	function startCurrentTime() {
		let today = new Date();
		let h = today.getHours();
		let m = today.getMinutes();
		let s = today.getSeconds();
		let t;
		m = checkTime(m);
		s = checkTime(s);
		$('#current_time').html(h + ":" + m + ":" + s);
		t = setTimeout(startCurrentTime, 500);
	}

	function cleanupForTemp(usePreTmp) {
		let dir = "";
		if (usePreTmp) {
			dir = preTmpDir;
		} else {
			dir = tmpDir;
		}
		if (dir === "") {
			return;
		}
		if (fs.existsSync(dir)) {
			fs.removeSync(dir);
		}
	}

	function cleanupForExit() {
		ipc.sendSync("require", { lib: "ffi", func: "destroy", args: null });
		cleanupForTemp(false);
		reflectCache(true);
		ipc.send('remote', { name: "exit" });
	}

	function registerGlobalShortcut() {
		ipc.send("require", { lib: "electron-globalShortcut", func: "client", args: null });
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
				configData.startWithTheFirstSlideSelected = json.startWithTheFirstSlideSelected;
				configData.highPerformance = false;
				configData.renderer = json.renderer;
				configData.lang = json.lang;
				setLangRsc();
				updateMonitorList();
				ipc.send('remote', { name: "passConfigData", details: configData });
			});
		} else {
			// Do nothing
		}
	}

	function reflectCache(saveOnly) {
		const configFile = 'cache_client.js';
		let configPath = "";
		const remote = require('@electron/remote');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		const cacheData = {
			"showCheckerboard": $("#trans_checker").is(":checked"),
			"enableSlideTransition": $("#use_slide_transition").is(":checked"),
			"includeBackground": $("#with_background").is(":checked"),
			"monitorAlpha": $("#monitor_trans").is(":checked")
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
				$('#use_slide_transition').trigger("click");
			}

			if (
				(json.includeBackground && !cacheData.includeBackground) ||
				(!json.includeBackground && cacheData.includeBackground)
			) {
				$('#with_background').trigger("click");
			}

			if (
				(json.monitorAlpha && !cacheData.monitorAlpha) ||
				(!json.monitorAlpha && cacheData.monitorAlpha)
			) {
				$('#monitor_trans').trigger("click");
			}
		});
	}

	function getMultipleMonitors() {
		multipleMonitors = ipc.sendSync("monitor", {
			func: "get"
		});
		return multipleMonitors;
	}

	function assignMonitor(num) {
		ipc.send("monitor", {
			func: "assign",
			monitorNo: num
		});
	}

	function enableMonitorTransparent() {
		ipc.send("monitor", {
			file: tmpDir + "/Slide" + currentSlide.toString() + '.png',
			workerinit: loadBackgroundInit,
			mode: modePostFix,
			modeusebg: modeUseBg,
			func: "transparentOn"
		});
	}

	function disableMonitorTransparent() {
		ipc.send("monitor", {
			file: tmpDir + "/Slide" + currentSlide.toString() + '.png',
			workerinit: loadBackgroundInit,
			mode: modePostFix,
			modeusebg: modeUseBg,
			func: "transparentOff"
		});
	}

	function enableMonitor() {
		ipc.send("monitor", {
			func: "turnOn"
		});
	}

	function disableMonitor() {
		ipc.send("monitor", {
			func: "turnOff"
		});
	}

	function updateMonitorList() {
		$('#monitorList').html($('<option>', {
			value: "none",
			text: getLangRsc("ui-classic/slideshow-monitor-none", configData.lang)
		}));
		for (let i=0; i<getMultipleMonitors().length; i++) {
			let monNum = i + 1;
			$('#monitorList').append($('<option>', {
				value: monNum,
				text: getLangRsc("ui-classic/slideshow-monitor-monitor", configData.lang) + monNum
			}));
		}
	}

	function glInit() {
		let contentPerTab = [/*{
				type: 'component',
				componentName: 'area',
				componentState: { label: 'A' },
				title: 'Area',
				width: 15,
				height: 100,
				reorderEnabled: false,
				isClosable: false
			},*/
			{
				type: 'column',
				content:[
				{
					type: 'component',
					componentName: 'ndiViews',
					componentState: { label: 'B' },
					title: 'NDI Views',
					height: 50,
					reorderEnabled: false,
					isClosable: false
				},
				{
					type: 'component',
					componentName: 'slides',
					componentState: { label: 'C' },
					title: 'Slides',
					reorderEnabled: false,
					isClosable: false
				}
				]
			}];

		let glConfig = {
			settings: {
				showPopoutIcon: false,
				showMaximiseIcon: false,
				showCloseIcon: false
			},
			dimensions: {
				borderWidth: 5,
				minItemHeight: 10,
				minItemWidth: 10,
				headerHeight: 20,
				dragProxyWidth: 300,
				dragProxyHeight: 200
			},
			content: [{
				isClosable: false,
				type: "stack",
				title: "area",
				content: [{
					id: "tab1",
					name: "tab1",
					title: "-",
					reorderEnabled: false,
					type: "row",
					componentName: "test",
					isClosable: false,
					content: contentPerTab
				}/*, {
					title: "+",
					type: "row",
					reorderEnabled: false,
					componentName: "test2",
					isClosable: false
				}*/
				]
			}]
		};

		glLayout = new GoldenLayout( glConfig );

		//glLayout.registerComponent( 'area', function( container, componentState ){
		//	container.getElement().html( 'html text' );
		//	container.on('resize', function (tab) {
		//		if (container.width < 50 ) {
		//			container.setSize( 100, container.height );
		//		}
		//	});
		//});

		glLayout.registerComponent( 'ndiViews', function( container, componentState ){
			container.getElement().html( `
			<table class="right" border="0">
			<tbody>
			<tr id="rightTop">
				<td width="50%">
					<select class="image-picker show-html">
						<optgroup label="Screen" id="screen_grp">
						<option id="currentSlideText" data-img-label="CURRENT" data-img-src="./img/null_slide.png" value="Current" disabled>Current</option>
						</optgroup>
					</select>
				</td>
				<td width="40%" style="vertical-align: top;">
					<select class="image-picker show-html">
						<optgroup label="Screen" id="screen_grp">
						<option id="nextSlideText" data-img-label="NEXT" data-img-src="./img/null_slide.png" value="Next" disabled>Next</option>
						</optgroup>
					</select>
					<div id="buttons">
						<button type="button" class="button" id="prev">&nbsp;&lt;&nbsp;</button>
						<button type="button" class="button" id="next">&nbsp;&gt;&nbsp;</button>
						<button type="button" class="button" id="blk">&nbsp;B&nbsp;</button>
						<button type="button" class="button" id="wht">&nbsp;W&nbsp;</button>
						<button type="button" class="button" id="trn">&nbsp;T&nbsp;</button>
						<button type="button" class="button" id="empty">&nbsp;</button>
						<button type="button" class="button" id="smaller">&nbsp;</button>
						<button type="button" class="button" id="bigger">&nbsp;</button>
					</div>
					<div id="slideInfo">
						<div id="slide_cnt" class="slideInfo">SLIDE 0 / 0</div>
						<div id="current_time" class="slideInfo">00:00 AM</div>
					</div>
				</td>
			</tr>
			</tbody>
			</table>
		` );
		});

		glLayout.registerComponent( 'slides', function( container, componentState ){
			const wrapBox = $('<div/>');
			const rightBox = $('<div/>').attr('class', 'right').attr('id', 'below');
			const selectBox = $('<select/>').attr('class', 'image-picker');
			selectBox.append($('<optgroup/>').attr('label', 'Slides').attr('id', 'slide_grp'));
			rightBox.append(selectBox);
			wrapBox.append(rightBox);

			container.getElement().html(wrapBox.html());
			container.on('resize', function () {
				if (container.height < 150 ) {
					container.setSize( container.width, 150 );
				}
			});
		});

		glLayout.init();

		for (let i = 0; i < $(".lm_title").length; i++) {
			let txtObj = $($(".lm_title")[i]);
			if ( txtObj.text() === "-" ) {
				txtObj.attr('id', 'ppt_filename');
				break;
			}
		}

		const cancelBox = $('<div>').attr('class', 'cancelBox');
		cancelBox
		.append($('<div/>').attr('id', 'loadingTxt'))
		.append($('<div/>').attr('id', 'cancel').attr('class', 'button'));
		$(".lm_goldenlayout")
		.append($('<div/>').attr('id', 'fullblack'))
		.append(cancelBox);
		setLangRsc();
	}

	prepare();
	glInit();
	registerEvents();
	initImgPicker();
	startCurrentTime();
	registerGlobalShortcut();
	reflectCache(false);
});
