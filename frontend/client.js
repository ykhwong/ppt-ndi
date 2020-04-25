const vbsBg = `
var objPPT;
var TestFile;
objPPT = new ActiveXObject("PowerPoint.Application");

function proc(ap) {
	var sl;
	var fn;
	for (var i=1; i<=ap.Slides.Count; i++) {
		sl = ap.Slides.Item(i);
		if (sl.SlideShowTransition.Hidden) {
			var objFileToWrite = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/hidden.dat",8,true);
			objFileToWrite.WriteLine(sl.SlideIndex);
			objFileToWrite.Close();
			objFileToWrite = null;
		}

		var objSlideEffect = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/slideEffect.dat",8,true);
		objSlideEffect.WriteLine(sl.SlideIndex + "," + sl.SlideShowTransition.EntryEffect + "," + sl.SlideShowTransition.Duration);
		objSlideEffect.Close();
		objSlideEffect = null;
		fn = WScript.arguments(1) + "/Slide" + sl.SlideIndex + ".png";
		if (WScript.arguments(2) === "0") {
			sl.Export(fn, "PNG");
		} else {
			sl.Export(fn, "PNG", WScript.arguments(2), WScript.arguments(3));
		}
	}
}

function main() {
	objPPT.DisplayAlerts = false;
	ap = objPPT.Presentations.Open(WScript.arguments(0), false, false, false);
	proc(ap);

	for (var i=0; i< objPPT.Presentations.length; i++) {
		var opres = objPPT.Presentations[i];
		TestFile = opres.FullName;
		break;
	}

	if (TestFile === "") {
		objPPT.quit;
	}
	objPPT = null;
	WScript.Echo("PPTNDI: Loaded");
}
main();
`

const vbsNoBg = `
var objPPT;
var TestFile;
var opres;
objPPT = new ActiveXObject("PowerPoint.Application");

function proc(ap) {
	var sl;
	var fn;
	var shGroup;
	var sngWidth;
	var sngHeight;

	sngWidth = ap.PageSetup.SlideWidth;
	sngHeight = ap.PageSetup.SlideHeight;

	for (var i=1; i<=ap.Slides.Count; i++) {
		sl = ap.Slides.Item(i);
		if (sl.SlideShowTransition.Hidden) {
			var objFileToWrite = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/hidden.dat",8,true);
			objFileToWrite.WriteLine(sl.SlideIndex);
			objFileToWrite.Close();
			objFileToWrite = null;
		}

		var objSlideEffect = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/slideEffect.dat",8,true);
		objSlideEffect.WriteLine(sl.SlideIndex + "," + sl.SlideShowTransition.EntryEffect + "," + sl.SlideShowTransition.Duration);
		objSlideEffect.Close();
		objSlideEffect = null;

		fn = WScript.arguments(1) + "/Slide" + sl.SlideIndex + ".png";
		var shp = sl.Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight );
		var shpGroup = sl.Shapes.Range();
		if (WScript.arguments(2) === "0") {
			shpGroup.Export(fn, 2, 0, 0, 1);
		} else {
			shpGroup.Export(fn, 2, Math.round(WScript.arguments(2) / 1.33333333), Math.round(WScript.arguments(3) / 1.33333333), 1);
		}
		shp.Delete();

		var fso = new ActiveXObject("Scripting.FileSystemObject");
		if (fso.FileExists(fn)) {
			var objFile = fso.GetFile(fn);
			if (objFile.size === 0) {
				for (var intShape = 1; i<=sl.Shapes.Count(); intShape++) {
					if (sl.Shapes(intShape).Type === 7) {
						sl.Shapes(intShape).Delete();
					}
				}
				var shp2 = sl.Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight);
				var shpGroup2 = sl.Shapes.Range();
				if (WScript.arguments(2) === "0") {
					shpGroup2.Export(fn, 2, 0, 0, 1);
				} else {
					shpGroup2.Export(fn, 2, Math.round(WScript.arguments(2) / 1.33333333), Math.round(WScript.arguments(3) / 1.33333333), 1);
				}
				shp2.Delete();
			}
		}

	}
}

function main() {
	objPPT.DisplayAlerts = false;
	ap = objPPT.Presentations.Open(WScript.arguments(0), false, false, false);
	proc(ap);

	for (var i=0; i< objPPT.Presentations.length; i++) {
		var opres = objPPT.Presentations[i];
		TestFile = opres.FullName;
		break;
	}

	if (TestFile === "") {
		objPPT.quit;
	}
	objPPT = null;
	WScript.Echo("PPTNDI: Loaded");
}
main();
`;

$(document).ready(function() {
	const { remote } = require('electron');
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const runtimeUrl = "https://aka.ms/vs/16/release/vc_redist.x64.exe";
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
	let hiddenSlides = [];
	let slideEffects = {};
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
	let pptTimestamp = 0;
	let repo;
	let slideTranTimers = [];
	let multipleMonitors;

	try {
		process.chdir(remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, ""));
	} catch(e) {
	}

	ffi = ipc.sendSync("require", { lib: "ffi", func: null, args: null });
	if ( ffi === -1 ) {
		const initFailMsg = "DLL init failed.";
		alertMsg("DLL init failed.");
		if (fs.existsSync("PPTNDI.DLL") && fs.existsSync("Processing.NDI.Lib.x64.dll")) {
			const execRuntime = require('child_process').execSync;
			execRuntime("start " + runtimeUrl, (error, stdout, stderr) => { 
				callback(stdout); 
			});
		}
		ipc.send('remote', { name: "exit" });
	}

	lib = ipc.sendSync("require", { lib: "ffi", func: "init", args: null });
	if (lib === 1) {
		alertMsg('Failed to create a listening server!');
		ipc.send('remote', { name: "exit" });
		return;
	}

	function alertMsg(myMsg) {
		const {dialog} = require('electron').remote;
		let options;
		options = {
			type: 'info',
			message: "" + myMsg,
			buttons: ["OK"]
		};
		dialog.showMessageBoxSync(currentWindow, options);
	}

	function stopSlideTransition() {
		for (var pp=2; pp<=9; pp++) {
			clearTimeout(slideTranTimers[pp]);
		}
		mustStop = true;
	}

	function createNullSlide() {
		const now = new Date().getTime();
		const PNG = require('pngjs').PNG;
		let png;
		let buffer;
		function getMeta(url, callback) {
			var img = new Image();
			img.src = url;
			img.onload = function() { callback(this.width, this.height); }
		}
		function createSli(redVal, greenVal, blueVal, alphaVal, fileVal) {
			png = new PNG({
				width: slideWidth,
				height: slideHeight,
				filterType: -1
			});

			for (var y = 0; y < png.height; y++) {
				for (var x = 0; x < png.width; x++) {
					var idx = (png.width * y + x) << 2;
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
			func: "update"
		});
	}

	function updateScreen(noTran) {
		let curSli, nextSli;
		let nextNum;
		let re, rpc;
		if(!repo) {
			return;
		}
		rpc = tmpDir + "/Slide";
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
				for (var i=2; i<=transLvl; i++) {
					fs.unlinkSync(tmpDir + "/t" + i.toString() + ".png");
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
								tmpDir + "/Slide" + currentSlide.toString() + ".png",
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
							tmpDir + "/t" + i.toString() + ".png",
							true
						]
					});
					if ( i % 2 === 0 ) {
						let now = new Date().getTime();
						$("img.image_picker_image:first").attr("src", tmpDir + "/t" + i.toString() + ".png?" + now);
					}
				}, i * parseFloat(duration) * 50);
				if (i === transLvl) {
					setLast();
				}
			}

			function doTrans() {
				/*
				if (curSli === prevSli) {
					return;
				}
				*/

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
							fs.writeFileSync(tmpDir + "/t" + i.toString() + ".png", b64data, 'base64');
							if (i === 8) {
								for (var i2=2; i2<=transLvl; i2++) {
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
	}

	$("select").change(function() {
		if (repo == null) {
			repo = $(this);
		}
	});

	function askReloadFile(ele, msg, description) {
		const {dialog} = require('electron').remote;
		let options;
		let response;
		let defaultMsg = 'Do you want to continue?';
		let defaultDetail = 'Changes require a reload to take effect.';
		
		options = {
			type: 'question',
			buttons: ['Yes', 'No'],
			defaultId: 1,
			//title: '',
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

	$("#with_background").click(function() {
		if (maxSlideNum > 0) {
			askReloadFile(this, "", "");
		}
	});

	function fitHeight() {
		let autoHeight = belowImgWidth*slideHeight/slideWidth + "px";
		$("#below img").css({
			'background' : 'black',
			'width' : belowImgWidth + "px",
			'height' : autoHeight
		});
	}

	function initImgPicker() {
		$("#right select").imagepicker({
			hide_select: true,
			show_label: true,
			selected:function(select, picker_option, event) {
				prevSlide = currentSlide;
				currentSlide=$('.selected').text();
				updateScreen(false);
			}
		});
		if ($("#trans_checker").is(":checked")) {
			$("#right img").css('background-image', "url('trans_slide.png')");
		} else {
			$("#right img").css('background-image', "url('trans.png')");
		}
		fitHeight();
		$("img.image_picker_image:first, img.image_picker_image:eq(1)").click(function() {
			gotoNext();
		});
		$(window).trigger('resize');
	}

	function cancelLoad() {
		const kill  = require('tree-kill');
		kill(spawnpid);
		cleanupForTemp(false);
		tmpDir = preTmpDir;
		$("#fullblack, .cancelBox").hide();
	}

	function loadPPTX(file) {
		if (file === undefined) {
			$("#fullblack, .cancelBox").hide();
			return false;
		}
		$("#fullblack").show();
		isCancelTriggered = false;
		let re = new RegExp("\\.(ppt|pptx)\$", "i");
		let vbsDir, res;
		let fileArr = [];
		let options = "";
		let resX = 0;
		let resY = 0;
		if (re.exec(file)) {
			let now = new Date().getTime();
			let newVbsContent;
			const spawn = require( 'child_process' ).spawn;
			spawnpid = spawn.pid;
			preTmpDir = tmpDir;
			tmpDir = process.env.TEMP + '/ppt_ndi';
			if (!fs.existsSync(tmpDir)) {
				fs.mkdirSync(tmpDir);
			}
			tmpDir += '/' + now;
			fs.mkdirSync(tmpDir);
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
				alertMsg('Failed to access the temporary directory!');
				$("#fullblack, .cancelBox").hide();
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
			$(".cancelBox").show();
			res.stderr.on('data', (data) => {
				let myMsg = "Failed to parse the presentation!\n";
				maxSlideNum = 0;
				cleanupForTemp(false);
				tmpDir = preTmpDir;
				if (!fs.existsSync(file)) {
					alertMsg(myMsg + "The file might have been moved or deleted.");
				} else if (maxSlideNum > 0) {
					alertMsg(myMsg + "Please check the configuration.");
				} else {
					alertMsg(myMsg + "Please make sure that Microsoft PowerPoint has been installed on the system.");
				}
				$("#fullblack, .cancelBox").hide();
				return;
			});
			res.on('close', (code) => {
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
				if (isCancelTriggered) return;
				if (fileArr === undefined || fileArr.length == 0) {
					maxSlideNum = 0;
					cleanupForTemp(false);
					tmpDir = preTmpDir;
					alertMsg("Presentation file could not be loaded.\n\nPlease check whether the presentension has one or more slides.\nAlso, please remove missing fonts if applicable.");
					$("#fullblack, .cancelBox").hide();
					return;
				}
				hiddenSlides = [];
				if (fs.existsSync(tmpDir + "/hidden.dat")) {
					const hs = fs.readFileSync(tmpDir + "/hidden.dat", { encoding: 'utf8' });
					hiddenSlides = hs.split("\n");
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
				if (isCancelTriggered) return;
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
					if (!isHidden && ( slideEffects[rpc].effectName !== "0" )) {
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
				$("#fullblack, .cancelBox").hide();
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
					$("img.image_picker_image:eq(1)").attr("src", "null_slide.png");
					if (maxSlideNum === 1) {
						$("#below .thumbnail").click(function() {
							selectSlide('1');
							$(this).off('click');
						});
					}
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
			});
		} else {
			if (/\S/.test(file)) {
				alertMsg("Only allowed filename extensions are PPT and PPTX.");
			}
			$("#fullblack, .cancelBox").hide();
		}
	}

	$("#load_pptx").click(function() {
		const {dialog} = require('electron').remote;
		$("#fullblack").show();

		dialog.showOpenDialog(currentWindow,{
			properties: ['openFile'],
			filters: [
				{name: 'PowerPoint Presentations', extensions: ['pptx', 'ppt']},
				{name: 'All Files', extensions: ['*']}
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
	
	$('#prev').click(function() {
		gotoPrev();
	});

	$('#next').click(function() {
		gotoNext();
	});

	function updateBlkWhtTrn(color) {
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
			dirTo = __dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + color + "_slide.png";
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

	function makePreviewSmaller() {
		belowImgWidth -= 5;
		$("#below img").css("width", belowImgWidth + "px");
		fitHeight();
	}

	function makePreviewBigger() {
		belowImgWidth += 5;
		$("#below img").css("width", belowImgWidth + "px");
		fitHeight();
	}


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
		fitHeight();
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
		ipc.send('remote', { name: "exit" });
	}

	function registerIoHook() {
		let ioHook = ipc.sendSync("require", { lib: "iohook", on: null, args: null });
		ipc.sendSync("require", { lib: "iohook", on: "keyup", args: null });
		ipc.sendSync("require", { lib: "iohook", on: "mouseup", args: null });
		ipc.sendSync("require", { lib: "iohook", on: "mousedrag", args: null });
		ipc.sendSync("require", { lib: "iohook", func: "start", args: null });
	}

	function reflectConfig() {
		const configFile = 'config.js';
		let configPath = "";
		const { remote } = require('electron');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		if (!fs.existsSync(configPath)) {
			const appDataPath = process.env.APPDATA + "/PPT-NDI";
			configPath = appDataPath + "/" + configFile;
		}
		if (fs.existsSync(configPath)) {
			$.getJSON(configPath, function(json) {
				configData.hotKeys = json.hotKeys;
				configData.startWithTheFirstSlideSelected = json.startWithTheFirstSlideSelected;
				configData.highPerformance = json.highPerformance;
				ipc.send('remote', { name: "passConfigData", details: configData });
			});
		} else {
			// Do nothing
		}
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
			func: "transparentOn"
		});
	}

	function disableMonitorTransparent() {
		ipc.send("monitor", {
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
			text: 'None'
		}));
		for (let i=0; i<getMultipleMonitors().length; i++) {
			let monNum = i + 1;
			$('#monitorList').append($('<option>', {
				value: monNum,
				text: 'Monitor ' + monNum
			}));
		}
	}

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
						askReloadFile("", "This file has been modified. Do you want to reload it?", "");
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
			$("#right img").css('background-image', "url('trans_slide.png')");
		} else {
			$("#right img").css('background-image', "url('null_slide.png')");
		}
		$("#below img").css('background', 'black');
	});

	currentWindow.on('maximize', function (){
		$("#max_restore").attr("src", "restore.png");
    });

	currentWindow.on('unmaximize', function (){
		$("#max_restore").attr("src", "max.png");
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

	$(window).resize(function(){
		let ss = $(document).height() - $("#top").height() - $("#rest1").height() - $("#rest2 img:first").height() - 50;
		$("#below").height(ss);
	});

	$("#resWidth").val("0");
	$("#resHeight").val("0");
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
					alertMsg("The original aspect ratio does not match. The layout can be corrupted due to the different ratios.");
				}
				askReloadFile(null, "", "");
			}
		}
	});

	$("#setMonitor").click(function() {
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
	initImgPicker();
	startCurrentTime();
	registerIoHook();
	reflectConfig();
	updateMonitorList();
});
