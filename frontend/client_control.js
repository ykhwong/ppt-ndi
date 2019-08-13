const vbsBg =`
Dim objPPT
Dim preState
Dim ap
Dim curPos
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
		.Export Wscript.Arguments.Item(0) & "/Slide.png", "PNG"
	End With
	If isSaved = True Then
		ap.Saved = True
	End If
	Wscript.Echo "PPTNDI: Sent"
End Sub
sub Main()
	Do While True
		On Error Resume Next
		Err.Clear
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
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
			If curPos <> 0 Then
				If ap.Slides(curPos).SlideShowTransition.AdvanceOnTime = -1 Then
					Wscript.Sleep(250)
				Else
					Wscript.StdIn.ReadLine()
				End If
			Else
				Wscript.Sleep(500)
			End If
		End If
	Loop
End Sub
Main
`;

const vbsNoBg =`
Dim objPPT
Dim preState
Dim ap
Dim curPos
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
			shpGroup.Export Wscript.Arguments.Item(0) & "/Slide.png", 2, , , 1
			.Delete
		End With
		If isSaved = True Then
			ap.Saved = True
		End If
		Wscript.Echo "PPTNDI: Sent"
	End With
End Sub
sub Main()
	Do While True
		On Error Resume Next
		Err.Clear
		Set objPPT = CreateObject("PowerPoint.Application")
		If Err.Number = 0 Then
			Err.Clear
			Set ap = objPPT.ActivePresentation
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
			If curPos <> 0 Then
				If ap.Slides(curPos).SlideShowTransition.AdvanceOnTime = -1 Then
					Wscript.Sleep(250)
				Else
					Wscript.StdIn.ReadLine()
				End If
			Else
				Wscript.Sleep(500)
			End If
		End If
	Loop
End Sub
Main
`;

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
	const binPath = './bin/PPTNDI.EXE';
	let ignoreIoHook = false;
	let ioHook = null;
	let iohook2 = null;
	let tmpDir = null;
	let slideWidth = 0;
	let slideHeight = 0;
	let configData = {};
	let pin = true;
	let child;
	let res;
	let res2;

	function runBin() {
		if (fs.existsSync(binPath)) {
			child = spawn(binPath);
		} else {
			alert('Failed to create a listening server!');
			ipc.send('remote', "exit");
			return;
		}
	}

	function sendNullNDI() {
		const now = new Date().getTime();
		const file = "null_slide.png";
		$("#slidePreview").attr("src", file + "?" + now);
		try {
			child.stdin.write(__dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + file + "\n");
		} catch(e) {
			runBin();
			child.stdin.write(__dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + file + "\n");
		}
	}

	function sendNDI(file, data) {
		const now = new Date().getTime();
		const cmd = data.toString();
		const Jimp = require('jimp');
		if (/^PPTNDI: Sent/.test(cmd)) {
			// Do nothing
		} else if(/^PPTNDI: White/.test(cmd)) {
			file = "white_slide.png";
		} else if(/^PPTNDI: Black/.test(cmd)) {
			file = "black_slide.png";
		} else if(/^PPTNDI: (Done|Paused)/.test(cmd)) {
			//file = "null_slide.png";
			return;
		} else {
			return;
		}
		$("#slidePreview").attr("src", file + "?" + now);
		if (/^PPTNDI: (White|Black)/.test(cmd)) {
			try {
				child.stdin.write(__dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + file + "\n");
			} catch(e) {
				runBin();
				child.stdin.write(__dirname.replace(/app\.asar(\\|\/)frontend/, "") + "/" + file + "\n");
			}
		} else {
			try {
				child.stdin.write(file + "\n");
			} catch(e) {
				runBin();
				child.stdin.write(file + "\n");
			}
		}
		Jimp.read(tmpDir + "/Slide.png").then(image=> {
			slideWidth = image.bitmap.width;
			slideHeight = image.bitmap.height;
			$("#slideRes").html("( " + slideWidth + " x " + slideHeight + " )");
		});
	}

	function registerIoHook() {
		ioHook = require('iohook');
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

	function registerIoHook2() {
		ioHook2 = require('iohook');
		ioHook2.on('keyup', event => {
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
				}
			}
		});
		ioHook2.start();
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
		runBin();
		child.stdin.setEncoding('utf-8');
		child.stdout.pipe(process.stdout);

		tmpDir = process.env.TEMP + '/ppt_ndi';
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir);
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir);
		vbsDir = tmpDir + '/wb.vbs';
		vbsDir2 = tmpDir + '/wb2.vbs';
		file = tmpDir + "/Slide.png";

		newVbsContent = vbsNoBg;
		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			alert('Failed to access the temporary directory!');
			return;
		}
		try {
			fs.writeFileSync(vbsDir2, vbsDirectCmd, 'utf-8');
		} catch(e) {
		}
		if (fs.existsSync(vbsDir)) {
			res = spawn( 'cscript.exe', [ vbsDir, tmpDir, '' ] );
			res.stdout.on('data', function(data) {
				sendNDI(file, data);
			});
		} else {
			alert('Failed to parse the presentation!');
			return;
		}
		if (fs.existsSync(vbsDir2)) {
			res2 = spawn( 'cscript.exe', [ vbsDir2, '' ] );
		}
		// Enable Always On Top by default
		ipc.send('remote', "onTop");
		$("#pin").attr("src", "pin_green.png");
		pin = true;

		registerIoHook();
		registerIoHook2();
		reflectConfig();
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
		try {
			child.stdin.write("destroy\n");
		} catch(e) {
		}
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
				alert('Failed to access the temporary directory!');
				return;
			}
		} else {
			newVbsContent = vbsNoBg;
			try {
				fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
			} catch(e) {
				alert('Failed to access the temporary directory!');
				return;
			}
		}
		res.stdin.pause();
		res.kill();
		res = null;
		if (fs.existsSync(vbsDir)) {
			res = spawn( 'cscript.exe', [ vbsDir, tmpDir, '' ] );
			res.stdout.on('data', function(data) {
				sendNDI(file, data);
			});
		} else {
			alert('Failed to parse the presentation!');
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
