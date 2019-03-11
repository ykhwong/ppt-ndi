const vbsBg =`
Dim objPPT
Dim preVal
Dim preState
Dim ap
Set objPPT = CreateObject("PowerPoint.Application")
On Error Resume Next
Sub Proc(preVal)
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
		.Export Wscript.Arguments.Item(0) & "/Slide.png", "PNG"
	End With
	Wscript.Echo "PPTNDI: Sent"
End Sub
sub Main()
	objPPT.DisplayAlerts = False
	preVal = 0
	Do While True
		On Error Resume Next
		Set ap = objPPT.ActivePresentation
		If Err.Number = 0 Then
			If preVal = 0 Or ap.SlideShowWindow.View.CurrentShowPosition <> preVal Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
			End If
			If preState <> ap.SlideShowWindow.View.State Then
				preState = ap.SlideShowWindow.View.State
				If ap.SlideShowWindow.View.State = 3 Then
					Wscript.Echo "PPTNDI: Black"
				ElseIf ap.SlideShowWindow.View.State = 4 Then
					Wscript.Echo "PPTNDI: White"
				ElseIf ap.SlideShowWindow.View.State = 5 Then
					Wscript.Echo "PPTNDI: Done"
				End If
			End If
		End If
		WScript.Sleep(250)
	Loop
End Sub
Main
`;

const vbsNoBg =`
Dim objPPT
Dim preVal
Dim preState
Dim ap
Set objPPT = CreateObject("PowerPoint.Application")
On Error Resume Next
Sub Proc(preVal)
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
		With .Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight)
			.TextFrame.TextRange = ""
			Set shpGroup = ap.Slides(objSlideShow.CurrentShowPosition).Shapes.Range()
			shpGroup.Export Wscript.Arguments.Item(0) & "/Slide.png", _
								2, , , 1
			.Delete
		End With
		Wscript.Echo "PPTNDI: Sent"
	End With
End Sub
sub Main()
	objPPT.DisplayAlerts = False
	preVal = 0
	Do While True
		On Error Resume Next
		Set ap = objPPT.ActivePresentation
		If Err.Number = 0 Then
			If preVal = 0 Or ap.SlideShowWindow.View.CurrentShowPosition <> preVal Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
			End If
			If preState <> ap.SlideShowWindow.View.State Then
				preState = ap.SlideShowWindow.View.State
				If ap.SlideShowWindow.View.State = -1 Then
				ElseIf ap.SlideShowWindow.View.State = 3 Then
					Wscript.Echo "PPTNDI: Black"
				ElseIf ap.SlideShowWindow.View.State = 4 Then
					Wscript.Echo "PPTNDI: White"
				ElseIf ap.SlideShowWindow.View.State = 5 Then
					Wscript.Echo "PPTNDI: Done"
				ElseIf ap.SlideShowWindow.View.State = 1 Or ap.SlideShowWindow.View.State = 2 Then
					preVal = ap.SlideShowWindow.View.CurrentShowPosition
					Proc(preVal)
				End If
			End If
		End If
		WScript.Sleep(250)
	Loop
End Sub
Main
`;

$(document).ready(function() {
	const spawn = require( 'child_process' ).spawn;
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const binPath = './bin/PPTNDI.EXE';
	var tmpDir = null;
	var child;
	var res;

	function runBin() {
		if (fs.existsSync(binPath)) {
			child = spawn(binPath);
			//child.on('exit', function (code) {
			//	alert("EXITED " + code);
			//});
		} else {
			alert('Failed to create a listening server!');
			ipc.send('remote', "exit");
			return;
		}
	}

	function sendNDI(file, data) {
		var now = new Date().getTime();
		var cmd = data.toString();
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
		var file;
		var vbsDir;
		var newVbsContent;
		var now = new Date().getTime();
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
		var newVbsContent;
		var vbsDir = tmpDir + '/wb.vbs';
		var file = tmpDir + "/Slide.png";
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

	init();
});
