const ipc = require('electron').ipcRenderer;
const fs = require("fs-extra");
const spawn = require( 'child_process' ).spawn;
const util = require('util');

var tmpDir = null;
var preTime = 0;

var vbs_bg =`
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

var vbs_no_bg =`
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
			With .Fill
			.Visible = msoFalse
			End With
			.SetShapesDefaultProperties
			With .Line
			.Visible = msoFalse
			End With
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
	var child = spawn('./bin/PPTNDI');
	var res;

	child.on('exit', function (code) {
		//alert("EXITED " + code);
	});

	function init() {
		var vbs_dir;
		var new_vbs_content;
		var now = new Date().getTime();
		child.stdin.setEncoding('utf-8');
		child.stdout.pipe(process.stdout);

		tmpDir = process.env.TEMP + '/ppt_ndi';
		if (!fs.existsSync(tmpDir)) {
			fs.mkdirSync(tmpDir);
		}
		tmpDir += '/' + now;
		fs.mkdirSync(tmpDir);
		vbs_dir = tmpDir + '/wb.vbs';

		new_vbs_content = vbs_no_bg;
		try {
			fs.writeFileSync(vbs_dir, new_vbs_content, 'utf-8');
		} catch(e) {
			alert('Failed to access the temporary directory!');
			return;
		}
		res = spawn( 'cscript.exe', [ vbs_dir, tmpDir, '' ] );
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

	function cleanup_for_temp() {
		if (fs.existsSync(tmpDir)) {
			fs.removeSync(tmpDir);
		}
	}

	function cleanup_for_exit() {
		try {
			child.stdin.write("destroy\n");
		} catch(e) {
		}
		cleanup_for_temp();
		ipc.send('remote', "exit");
	}

	ipc.on('remote' , function(event, data){
		if (data.msg == "exit") {
			cleanup_for_exit();
		}
	});

	$('#closeImg').click(function() {
		cleanup_for_exit();
	});

	$('#bk').click(function() {
		var new_vbs_content;
		var vbs_dir = tmpDir + '/wb.vbs';
		if ($("#bk").is(":checked")) {
			new_vbs_content = vbs_bg;
			try {
				fs.writeFileSync(vbs_dir, new_vbs_content, 'utf-8');
			} catch(e) {
				alert('Failed to access the temporary directory!');
				return;
			}
		} else {
			new_vbs_content = vbs_no_bg;
			try {
				fs.writeFileSync(vbs_dir, new_vbs_content, 'utf-8');
			} catch(e) {
				alert('Failed to access the temporary directory!');
				return;
			}
		}
		res.stdin.pause();
		res.kill();
		res = spawn( 'cscript.exe', [ vbs_dir, tmpDir, '' ] );
	});


	init();
	refreshSlide();
});
