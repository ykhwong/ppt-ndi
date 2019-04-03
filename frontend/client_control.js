const vbsBg =`
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
	Dim isSaved
	With ap.PageSetup
		sngWidth = .SlideWidth
		sngHeight = .SlideHeight
	End With
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
		With .Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight)
			Set shpGroup = ap.Slides(objSlideShow.CurrentShowPosition).Shapes.Range()
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

$(document).ready(function() {
	const spawn = require( 'child_process' ).spawn;
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const ioHook = require('iohook');
	const binPath = './bin/PPTNDI.EXE';
	let tmpDir = null;
	let pin = true;
	let child;
	let res;

	function runBin() {
		if (fs.existsSync(binPath)) {
			child = spawn(binPath);
		} else {
			alert('Failed to create a listening server!');
			ipc.send('remote', "exit");
			return;
		}
	}

	function sendNDI(file, data) {
		let now = new Date().getTime();
		let cmd = data.toString();
		if (/^PPTNDI: Sent/.test(cmd)) {
			// Do nothing
		} else if(/^PPTNDI: White/.test(cmd)) {
			file = "white_slide.png";
		} else if(/^PPTNDI: Black/.test(cmd)) {
			file = "black_slide.png";
		} else if(/^PPTNDI: Done/.test(cmd)) {
			//file = "null_slide.png";
		} else {
			return;
		}
		$("#slidePreview").attr("src", file + "?" + now);
		if (/^PPTNDI: (White|Black)/.test(cmd)) {
			try {
				child.stdin.write(__dirname + "/" + file + "\n");
			} catch(e) {
				runBin();
				child.stdin.write(__dirname + "/" +file + "\n");
			}
		} else {
			try {
				child.stdin.write(file + "\n");
			} catch(e) {
				runBin();
				child.stdin.write(file + "\n");
			}
		}
	}

	function init() {
		let file;
		let vbsDir;
		let newVbsContent;
		let now = new Date().getTime();
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
		file = tmpDir + "/Slide.png";

		newVbsContent = vbsNoBg;
		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			alert('Failed to access the temporary directory!');
			return;
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

		// Enable Always On Top by default
		ipc.send('remote', "onTop");
		$("#pin").attr("src", "pin_green.png");
		pin = true;

		ioHook.on('keydown', event => {
			res.stdin.write("\n");
		});
		ioHook.on('mouseclick', event => {
			res.stdin.write("\n");
		});
		ioHook.on('mousewheel', event => {
			res.stdin.write("\n");
		});

		ioHook.start();
	}

	function cleanupForTemp() {
		if (fs.existsSync(tmpDir)) {
			fs.removeSync(tmpDir);
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
