/* vbsBg: Handles the slides with background included */
const vbsBg =`
Dim objPPT
Dim preState
Dim ap
Dim curPos
Dim newWidth
Dim newHeight

On Error Resume Next
Sub Proc()
	Dim sl
	Dim shGroup
	Dim isSaved
	Set objSlideShow = ap.SlideShowWindow.View
	If ap.Saved Then
		isSaved = True
	Else
		isSaved = False
	End If
	With ap.Slides(objSlideShow.CurrentShowPosition)
		If newWidth = 0 Then
			.Export Wscript.Arguments.Item(0) & "/Slide.png", "PNG"
		Else
			.Export Wscript.Arguments.Item(0) & "/Slide.png", "PNG", newWidth, newHeight
		End If
	End With
	If isSaved = True Then
		ap.Saved = True
	End If
	Dim entryEffect
	Dim duration
	entryEffect = ap.Slides(curPos).SlideShowTransition.EntryEffect
	duration = ap.Slides(curPos).SlideShowTransition.Duration
	Wscript.Echo "PPTNDI: Sent " & duration & " " & entryEffect & " " & objSlideShow.CurrentShowPosition
End Sub
sub Main()
	newWidth = 0
	newHeight = 0

	If Wscript.Arguments.Item(1) = 0 Then
	Else
		newWidth = Wscript.Arguments.Item(1)
		newHeight = Wscript.Arguments.Item(2)
	End If

	Do While True
		On Error Resume Next
		Err.Clear
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
			curPos = 0
			If Err.Number = 0 Then
				objPPT.DisplayAlerts = False
				Err.Clear
				curPos = ap.SlideShowWindow.View.CurrentShowPosition
				If Err.Number = 0 Then
					If ap.SlideShowWindow.View.State = -1 Then
					ElseIf ap.SlideShowWindow.View.State = 2 Then
						Wscript.Echo "PPTNDI: Paused"
					ElseIf ap.SlideShowWindow.View.State = 3 Then
						Wscript.Echo "PPTNDI: Black"
					ElseIf ap.SlideShowWindow.View.State = 4 Then
						Wscript.Echo "PPTNDI: White"
					ElseIf ap.SlideShowWindow.View.State = 5 Then
						Wscript.Echo "PPTNDI: Done"
					ElseIf ap.SlideShowWindow.View.State = 1 Or ap.SlideShowWindow.View.State = 2 Then
						Proc()
					End If
				Else
					curPos = 0
				End If
			End If
		End If
		cmd = Wscript.StdIn.ReadLine()
		If left(cmd, 6) = "setRes" Then
			Dim p1
			Dim res
			p1 = Replace(cmd, "setRes ", "")
			res = Split(p1, "x")
			newWidth = res(0)
			newHeight = res(1)
		End If
	Loop
End Sub
Main
`;

/* vbsNoBg: Handles the slides with no background included */
const vbsNoBg =`
Dim objPPT
Dim preState
Dim ap
Dim curPos
Dim newWidth
Dim newHeight

On Error Resume Next
Sub Proc()
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight
	Dim origGrpCnt
	oriGrpCnt = 0
	With ap.PageSetup
		sngWidth = .SlideWidth
		sngHeight = .SlideHeight
	End With
	Set objSlideShow = ap.SlideShowWindow.View
	With ap.Slides(objSlideShow.CurrentShowPosition)
		Dim isSaved
		If ap.Saved Then
			isSaved = True
		Else
			isSaved = False
		End If
		origGrpCnt = ap.Slides(objSlideShow.CurrentShowPosition).Shapes.Range().Count

		With .Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight)
			Set shpGroup = ap.Slides(objSlideShow.CurrentShowPosition).Shapes.Range()
			If shpGroup.Count = origGrpCnt Then
				.Delete
				origGrpCnt = 0
				If isSaved = True Then
					ap.Saved = True
				End If
				Exit Sub
			End If
			If newWidth = 0 Then
				shpGroup.Export Wscript.Arguments.Item(0) & "/Slide.png", 2, , , 1
			Else
				shpGroup.Export Wscript.Arguments.Item(0) & "/Slide.png", 2, Round(newWidth / 1.33333333, 0), Round(newHeight / 1.33333333, 0), 1
			End If
			.Delete
		End With
		If isSaved = True Then
			ap.Saved = True
		End If
		Dim entryEffect
		Dim duration
		entryEffect = ap.Slides(curPos).SlideShowTransition.EntryEffect
		duration = ap.Slides(curPos).SlideShowTransition.Duration
		Wscript.Echo "PPTNDI: Sent " & duration & " " & entryEffect & " " & objSlideShow.CurrentShowPosition
	End With
End Sub
sub Main()
	newWidth = 0
	newHeight = 0

	If Wscript.Arguments.Item(1) = 0 Then
	Else
		newWidth = Wscript.Arguments.Item(1)
		newHeight = Wscript.Arguments.Item(2)
	End If

	Do While True
		On Error Resume Next
		Err.Clear
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
			curPos = 0
			If Err.Number = 0 Then
				objPPT.DisplayAlerts = False
				Err.Clear
				curPos = ap.SlideShowWindow.View.CurrentShowPosition
				If Err.Number = 0 Then
					If ap.SlideShowWindow.View.State = -1 Then
					ElseIf ap.SlideShowWindow.View.State = 2 Then
						Wscript.Echo "PPTNDI: Paused"
					ElseIf ap.SlideShowWindow.View.State = 3 Then
						Wscript.Echo "PPTNDI: Black"
					ElseIf ap.SlideShowWindow.View.State = 4 Then
						Wscript.Echo "PPTNDI: White"
					ElseIf ap.SlideShowWindow.View.State = 5 Then
						Wscript.Echo "PPTNDI: Done"
					ElseIf ap.SlideShowWindow.View.State = 1 Or ap.SlideShowWindow.View.State = 2 Then
						Proc()
					End If
				Else
					curPos = 0
				End If
			End If
		End If
		cmd = Wscript.StdIn.ReadLine()
		If left(cmd, 6) = "setRes" Then
			Dim p1
			Dim res
			p1 = Replace(cmd, "setRes ", "")
			res = Split(p1, "x")
			newWidth = res(0)
			newHeight = res(1)
		End If
	Loop
End Sub
Main
`;

/* vbsCheckSlide: Checks the status of current slide in real time */
const vbsCheckSlide =`
Dim objPPT
Dim preSlideIdx
Dim curPos
preSlideIdx = 0

sub Main()
	Wscript.Echo "Status: 0"
	Do While True
		On Error Resume Next
		Err.Clear
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
			If Err.Number = 0 Then
				Err.Clear
				curPos = ap.SlideShowWindow.View.CurrentShowPosition
				If preSlideIdx = curPos Then
				Else
					Wscript.Echo "Status: " & curPos
					preSlideIdx = curPos
				End If
			Else
				Wscript.Echo "Status: OFF"
			End If
		Else
			preSlideIdx = 0
			Wscript.Echo "Status: 0"
		End If
		Wscript.Sleep(500)
	Loop
End sub
Main
`;

/* vbsDiretCmd: Handles hotkey */
const vbsDirectCmd = `
Dim objPPT
Dim cmd
sub Main()
	Do While True
		On Error Resume Next
		cmd = Wscript.StdIn.ReadLine()
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
			If Err.Number = 0 Then
				Err Clear
				Set objSlideShow = ap.SlideShowWindow.View
					If cmd = "prev" Then
						objSlideShow.GotoSlide objSlideShow.CurrentShowPosition - 1
					End If
					If cmd = "next" Then
						objSlideShow.GotoSlide objSlideShow.CurrentShowPosition + 1
					End If
					If cmd = "black" Then
						ap.SlideShowWindow.View.State = 3
					End If
					If cmd = "white" Then
						ap.SlideShowWindow.View.State = 4
					End If
					If cmd = "pause" Then
						ap.SlideShowWindow.View.State = 5
					End If
			End If
		End If
	Loop
End Sub
Main
`;

$(document).ready(function() {
	const spawn = require( 'child_process' ).spawn;
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const ffi = require('ffi');
	let lib;
	let ignoreIoHook = false;
	let ioHook = null;
	let tmpDir = null;
	let preFile = "";
	let slideWidth = 0;
	let slideHeight = 0;
	let customSlideX = 0;
	let customSlideY = 0;
	let lastSignalTime = 0;
	let lastSlideStat = "0";
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

	function alertMsg(myMsg) {
		const { remote } = require('electron');
		const {dialog} = require('electron').remote;
		let currentWindow = remote.getCurrentWindow();
		let options;
		options = {
			type: 'info',
			message: myMsg,
			buttons: ["OK"]
		};
		dialog.showMessageBoxSync(currentWindow, options);
	}

	function runLib() {
		try {
			const { remote } = require('electron');
			let { RTLD_NOW, RTLD_GLOBAL } = ffi.DynamicLibrary.FLAGS;
			ffi.DynamicLibrary(
				remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/Processing.NDI.Lib.x64.dll',
				RTLD_NOW | RTLD_GLOBAL
			);
			//lib = ffi.Library('./PPTNDI.dll', {
			lib = ffi.Library(remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + '/PPTNDI.dll', {
				'init': [ 'int', [] ],
				'destroy': [ 'int', [] ],
				'send': [ 'int', [ "string", "bool" ] ]
			});
		} catch(e) {
			alertMsg(e);
			ipc.send('remote', "exit");
		}

		if (lib.init() === 1) {
			alertMsg('Failed to create a listening server!');
			ipc.send('remote', "exit");
			return;
		}
	}

	function stopSlideTransition() {
		for (var pp=2; pp<=9; pp++) {
			clearTimeout(slideTranTimers[pp]);
		}
		mustStop = true;
	}

	function sendNullNDI() {
		const now = new Date().getTime();
		const file = "null_slide.png";
		$("#slidePreview").attr("src", file + "?" + now);
		lib.send(__dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + file, false);
	}

	function sendNDI(file, data) {
		const now = new Date().getTime();
		const cmd = data.toString();
		const Jimp = require('jimp');
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
		} else if(/^PPTNDI: Black/.test(cmd)) {
			file = tmpDir + "/SlideBlack.png"
			newSlideIdx = "black";
		} else if(/^PPTNDI: (Done|Paused)/.test(cmd)) {
			//file = "null_slide.png";
			$("#tip").html("Status: EXITED/PAUSED");
			return;
		} else {
			console.log(cmd);
			return;
		}

		$("#tip").html("Status: OK");
		if (/^PPTNDI: Sent /.test(cmd)) {
			let fd;
			try {
				fd = fs.openSync(file, 'r+');
			} catch (err) {
				if (err && err.code === 'EBUSY'){
					//console.log("busy");
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

			$("#slidePreview").attr("src", file + "?" + now);
			Jimp.read(file).then(image=> {
				slideWidth = image.bitmap.width;
				slideHeight = image.bitmap.height;
				$("#slideRes").html("( " + slideWidth + " x " + slideHeight + " )");
			});
		}

		if (slideIdx === newSlideIdx) {
			if (lastSignalTime >= (Date.now() - 500)) {
				return;
			}
		}
		console.log(cmd);
		slideIdx = newSlideIdx;
		lastSignalTime = Date.now();

		if (/^PPTNDI: (White|Black)/.test(cmd)) {
			function fillColor(color) {
				const Jimp = require('jimp');
				new Jimp(slideWidth, slideHeight, color, (err, image2) => {
					image2.opacity(1);
					image2.write(file, function() {
						$("#slidePreview").attr("src", file + "?" + now);
						lib.send(file, false);
					});
				});
			}
			if (newSlideIdx === "black") {
				fillColor(0x000000FF);
			} else {
				fillColor(0xFFFFFFFF);
			}
		} else {
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
			lib.send(file, false);
		}
	}

	function registerIoHook() {
		ioHook = require('iohook');
		ioHook.on('keyup', event => {
			if (event.shiftKey && event.ctrlKey) {
				let chr = String.fromCharCode( event.rawcode );
				if (chr === "") return;
				switch (chr) {
					case configData.hotKeys.prev: res2.stdin.write("prev\n"); res.stdin.write("\n"); break;
					case configData.hotKeys.next: res2.stdin.write("next\n"); res.stdin.write("\n"); break;
					case configData.hotKeys.transparent:
						setTimeout(function() {
							ignoreIoHook = true;
							sendNullNDI();
							ignoreIoHook = false;
						}, 500);
						break;
					case configData.hotKeys.black: res2.stdin.write("black\n"); res.stdin.write("\n"); break;
					case configData.hotKeys.white: res2.stdin.write("white\n"); res.stdin.write("\n"); break;
					default: break;
				}
			}
		});
		ioHook.on('keydown', event => {
			if (!ignoreIoHook) {
				res.stdin.write("\n");
			}
		});
		ioHook.on('mouseup', event => {
			if (!ignoreIoHook) {
				res.stdin.write("\n");
			}
		});
		ioHook.on('mousewheel', event => {
			if (!ignoreIoHook) {
				res.stdin.write("\n");
			}
		});
		ioHook.start();
	}

	function procTransition(file, data) {
		const transLvl=9;
		preFile = tmpDir + "/SlidePre.png";

		try {
			for (var i=2; i<=transLvl; i++) {
				fs.unlinkSync(tmpDir + "/t" + i.toString() + ".png");
			}
		} catch(e) {
		}

		function sendSlides(i) {
			console.log(i);
			if (mustStop) {
				return;
			}
			function setLast() {
				if (mustStop) {
					return;
				}
				slideTranTimers[10] = setTimeout(function() {
					lib.send(tmpDir + "/Slide.png", false);
					if (fs.existsSync(file)) {
						try {
							fs.copySync(file, preFile);
						} catch(e) {
							console.log("file could not be generated: "+ preFile);
						}
					}
				}, 10 * parseFloat(duration) * 50);
			}

			slideTranTimers[i] = setTimeout(function() {
				lib.send(tmpDir + "/t" + i.toString() + ".png", true);
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

			for (let i=2; i<=transLvl; i++) {
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
	}

	function init() {
		const { remote } = require('electron');
		let file;
		let vbsDir;
		let vbsDir2;
		let newVbsContent;
		let now = new Date().getTime();
		try {
			process.chdir(remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, ""));
		} catch(e) {
		}
		runLib();

		tmpDir = process.env.TEMP + '/ppt_ndi';
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir);
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir);
		vbsDir = tmpDir + '/wb.vbs';
		vbsDir2 = tmpDir + '/wb2.vbs';
		vbsDir3 = tmpDir + '/wb3.vbs';
		file = tmpDir + "/Slide.png";

		newVbsContent = vbsNoBg;
		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			alertMsg('Failed to access the temporary directory!');
			return;
		}
		try {
			fs.writeFileSync(vbsDir2, vbsDirectCmd, 'utf-8');
		} catch(e) {
		}
		try {
			fs.writeFileSync(vbsDir3, vbsCheckSlide, 'utf-8');
		} catch(e) {
			alertMsg('Failed to access the temporary directory!');
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
			alertMsg('Failed to parse the presentation!');
			return;
		}
		if (fs.existsSync(vbsDir2)) {
			res2 = spawn( 'cscript.exe', [ vbsDir2, "//NOLOGO", '' ] );
		}
		if (fs.existsSync(vbsDir3)) {
			res3 = spawn( 'cscript.exe', [ vbsDir3, "//NOLOGO", '' ] );
		} else {
			alertMsg('Failed to parse the presentation!');
			return;
		}

		res3.stdout.on('data', function(data) {
			let curSlideStat = data.toString().replace(/^Status: /, "");
			if (/^\s*OFF\s*$/.test(curSlideStat)) {
				// Ready
				$("#tip").html("Status: READY<br />In PowerPoint, start the Slide Show.");
			} else if (/^\s*0\s*$/.test(curSlideStat)) {
				// Not found
				$("#tip").html("Status: -");
			} else {
				// ON
				$("#tip").html("Status: OK");
			}

			if (/^\s*0\s*$/.test(lastSlideStat)) {
				if (!/^\s*0\s*$/.test(curSlideStat)) {
					if (slideIdx != curSlideStat) {
						res.stdin.write("\n");
					}
				}
			}
			lastSlideStat = curSlideStat;
		});

		// Enable Always On Top by default
		ipc.send('remote', "onTop");
		$("#pin").attr("src", "pin_green.png");
		pin = true;

		// Enable Slide Checkerboard by default
		$("#slidePreview").css('background-image', "url('trans_slide.png')");

		registerIoHook();
		reflectConfig();

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

	}

	function cleanupForTemp() {
		if (fs.existsSync(tmpDir)) {
			fs.removeSync(tmpDir);
		}
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
			});
		} else {
			// Do nothing
		}
	}

	function cleanupForExit() {
		lib.destroy();
		cleanupForTemp();
		ipc.send('remote', "exit");
	}

	ipc.on('remote' , function(event, data){
		if (data.msg == "exit") {
			cleanupForExit();
			return;
		}
		if (data.msg == "reload") {
			reflectConfig();
			return;
		}
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
				alertMsg('Failed to access the temporary directory!');
				return;
			}
		} else {
			newVbsContent = vbsNoBg;
			try {
				fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
			} catch(e) {
				alertMsg('Failed to access the temporary directory!');
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
			alertMsg('Failed to parse the presentation!');
			return;
		}
	});

	$('#trans_checker').click(function() {
		if ($("#trans_checker").is(":checked")) {
			$("#slidePreview").css('background-image', "url('trans_slide.png')");
		} else {
			$("#slidePreview").css('background-image', "url('null_slide.png')");
		}
	});
	
	$('#pin').click(function() {
		if (pin) {
			ipc.send('remote', "onTopOff");
			$("#pin").attr("src", "pin_grey.png");
			pin = false;
		} else {
			ipc.send('remote', "onTop");
			$("#pin").attr("src", "pin_green.png");
			pin = true;
		}
	});

	init();
});
