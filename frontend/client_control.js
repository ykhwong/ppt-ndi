var tmpDir = null;
var preTime = 0;

var vbsBg =`
Dim objPPT
Dim preVal
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
End Sub
sub Main()
	objPPT.DisplayAlerts = False
	preVal = 0
	Do While True
		On Error Resume Next
		Set ap = objPPT.ActivePresentation
		If Err.Number = 0 Then
			If preVal = 0 Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
			ElseIf (ap.SlideShowWindow.View.CurrentShowPosition <> preVal) Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
			End If
		End If
		WScript.Sleep(250)
	Loop
End Sub
Main
`;

var vbsNoBg =`
Dim objPPT
Dim preVal
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
		With .Shapes.AddShape( 1, 0, 0, sngWidth, sngHeight)
			.Fill.Visible = msoFalse
			.Line.Visible = msoFalse
			.SetShapesDefaultProperties
		End With
		Set shpGroup = .Shapes.Range()
		shpGroup.Export Wscript.Arguments.Item(0) & "/Slide.png", _
							2, , , 1
		With ap.Slides(objSlideShow.CurrentShowPosition).Shapes
		For intShape = .Count To 1 Step -1
			With .Item(intShape)
				.Delete
			End With
			Exit For
		Next
		End With
	End With
End Sub
sub Main()
	objPPT.DisplayAlerts = False
	preVal = 0
	Do While True
		On Error Resume Next
		Set ap = objPPT.ActivePresentation
		If Err.Number = 0 Then
			If preVal = 0 Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
			ElseIf (ap.SlideShowWindow.View.CurrentShowPosition <> preVal) Then
				preVal = ap.SlideShowWindow.View.CurrentShowPosition
				Proc(preVal)
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
	var child = spawn('./bin/PPTNDI');
	var res;

	child.on('exit', function (code) {
		//alert("EXITED " + code);
	});

	function init() {
		var vbsDir;
		var newVbsContent;
		var now = new Date().getTime();
		child.stdin.setEncoding('utf-8');
		child.stdout.pipe(process.stdout);

		tmpDir = process.env.TEMP + '/ppt_ndi';
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir);
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir);
		vbsDir = tmpDir + '/wb.vbs';

		newVbsContent = vbsNoBg;
		try {
			fs.writeFileSync(vbsDir, newVbsContent, 'utf-8');
		} catch(e) {
			alert('Failed to access the temporary directory!');
			return;
		}
		res = spawn( 'cscript.exe', [ vbsDir, tmpDir, '' ] );
		/*
		if ( res.status !== 0 ) {
			alert('Failed to parse the presentation!');
			return;
		}
		*/
	}

	function refreshSlide() {
		var stats, mtime, file;
		if (tmpDir != null) {
			file = tmpDir + "/Slide.png";
			if (fs.existsSync(file)) {
				stats = fs.statSync(file);
				mtime = stats.mtime;
				if (mtime > preTime || mtime < preTime) {
					preTime = mtime;
					file = tmpDir + "/Slide.png";
					try {
						child.stdin.write(file + "\n");
					} catch(e) {
						child = spawn('./bin/PPTNDI');
						child.stdin.write(file + "\n");
					}
				}
			}
		}
		setTimeout(refreshSlide, 100);
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
		res = spawn( 'cscript.exe', [ vbsDir, tmpDir, '' ] );
	});

	init();
	refreshSlide();
});
