const cscript = {
	slideshow: {
		vbsBg: "",
		vbsNoBg: "",
		vbsCheckSlide: "",
		vbsDirectCmd: ""
	},
	classic: {
		vbsBg: "",
		vbsNoBg: "",
		vbsQuickEdit: ""
	}
};

cscript.slideshow.vbsBg = `
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
					ElseIf ap.SlideShowWindow.View.State = 1 Then
						Proc()
					ElseIf ap.SlideShowWindow.View.State = 2 Then
						'Wscript.Echo "PPTNDI: Paused" -- breaks hotkeys
						Proc()
					ElseIf ap.SlideShowWindow.View.State = 3 Then
						Wscript.Echo "PPTNDI: Black"
					ElseIf ap.SlideShowWindow.View.State = 4 Then
						Wscript.Echo "PPTNDI: White"
					ElseIf ap.SlideShowWindow.View.State = 5 Then
						Wscript.Echo "PPTNDI: Done"
					End If
				Else
					Wscript.Echo "PPTNDI: Ready"
					curPos = 0
				End If
			End If
		Else
			Wscript.Echo "PPTNDI: NoPPT"
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

cscript.slideshow.vbsNoBg =`
Dim objPPT
Dim preState
Dim ap
Dim curPos
Dim newWidth
Dim newHeight

On Error Resume Next

' Visual bounding box (in pt) of all visible shapes, including rotation and outline.
' ShapeRange.Export renders this whole area, not just the slide.
Sub GetBounds(shpRange, ByRef bL, ByRef bT, ByRef bR, ByRef bB)
	Dim i, s, w, h, rad, rw, rh, cx, cy, lw, vis
	For i = 1 To shpRange.Count
		Set s = shpRange(i)
		Err.Clear
		vis = s.Visible
		If Err.Number <> 0 Then vis = True
		If vis Then
			w = s.Width
			h = s.Height
			rad = s.Rotation * 3.14159265358979 / 180
			rw = Abs(w * Cos(rad)) + Abs(h * Sin(rad))
			rh = Abs(w * Sin(rad)) + Abs(h * Cos(rad))
			lw = 0
			Err.Clear
			If s.Line.Visible Then
				lw = s.Line.Weight / 2
			End If
			If Err.Number <> 0 Then lw = 0
			Err.Clear
			cx = s.Left + w / 2
			cy = s.Top + h / 2
			If cx - rw / 2 - lw < bL Then bL = cx - rw / 2 - lw
			If cy - rh / 2 - lw < bT Then bT = cy - rh / 2 - lw
			If cx + rw / 2 + lw > bR Then bR = cx + rw / 2 + lw
			If cy + rh / 2 + lw > bB Then bB = cy + rh / 2 + lw
		End If
	Next
End Sub

Sub Proc()
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight
	Dim origGrpCnt
	Dim bL, bT, bR, bB, bounds
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
			bL = 0
			bT = 0
			bR = sngWidth
			bB = sngHeight
			GetBounds shpGroup, bL, bT, bR, bB
			bounds = Round(bL * 100) & " " & Round(bT * 100) & " " & Round(bR * 100) & " " & Round(bB * 100) & " " & Round(sngWidth * 100) & " " & Round(sngHeight * 100) & " " & newWidth & " " & newHeight
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
		Wscript.Echo "PPTNDI: Sent " & duration & " " & entryEffect & " " & objSlideShow.CurrentShowPosition & " " & bounds
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
					ElseIf ap.SlideShowWindow.View.State = 1 Then
						Proc()
					ElseIf ap.SlideShowWindow.View.State = 2 Then
						'Wscript.Echo "PPTNDI: Paused" -- breaks hotkeys
						Proc()
					ElseIf ap.SlideShowWindow.View.State = 3 Then
						Wscript.Echo "PPTNDI: Black"
					ElseIf ap.SlideShowWindow.View.State = 4 Then
						Wscript.Echo "PPTNDI: White"
					ElseIf ap.SlideShowWindow.View.State = 5 Then
						Wscript.Echo "PPTNDI: Done"
					End If
				Else
					Wscript.Echo "PPTNDI: Ready"
					curPos = 0
				End If
			End If
		Else
			Wscript.Echo "PPTNDI: NoPPT"
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

cscript.slideshow.vbsCheckSlide =`
Dim objPPT
'Dim preSlideIdx
Dim curPos
'preSlideIdx = 0

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
				'If preSlideIdx = curPos Then
				'Else
					Wscript.Echo "Status: " & curPos
					'preSlideIdx = curPos
				'End If
			Else
				Wscript.Echo "Status: OFF"
			End If
		Else
			'preSlideIdx = 0
			Wscript.Echo "Status: 0"
		End If
		Wscript.Sleep(500)
	Loop
End sub
Main
`;

cscript.slideshow.vbsDirectCmd = `
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

cscript.classic.vbsBg = `
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
		if (sl.SlideShowTransition.AdvanceTime > 0) {
			var objFileToWrite = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/advance.dat",8,true);
			objFileToWrite.WriteLine(sl.SlideIndex + "\t" + sl.SlideShowTransition.AdvanceTime);
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

	for (var i=0; i< objPPT.Presentations.Count; i++) {
		var opres = objPPT.Presentations.Item(i + 1);
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

cscript.classic.vbsNoBg = `
var objPPT;
var TestFile;
var opres;
objPPT = new ActiveXObject("PowerPoint.Application");

// Visual bounding box (in pt) of all visible shapes, including rotation and outline.
// ShapeRange.Export renders this whole area, not just the slide.
function getBounds(shpRange, sngWidth, sngHeight) {
	var b = { l: 0, t: 0, r: sngWidth, b: sngHeight };
	for (var i = 1; i <= shpRange.Count; i++) {
		var s = shpRange.Item(i);
		var vis = true;
		try { vis = (s.Visible !== 0); } catch (e) {}
		if (!vis) {
			continue;
		}
		var w = s.Width;
		var h = s.Height;
		var rad = s.Rotation * Math.PI / 180;
		var rw = Math.abs(w * Math.cos(rad)) + Math.abs(h * Math.sin(rad));
		var rh = Math.abs(w * Math.sin(rad)) + Math.abs(h * Math.cos(rad));
		var lw = 0;
		try { if (s.Line.Visible !== 0) { lw = s.Line.Weight / 2; } } catch (e) { lw = 0; }
		var cx = s.Left + w / 2;
		var cy = s.Top + h / 2;
		b.l = Math.min(b.l, cx - rw / 2 - lw);
		b.t = Math.min(b.t, cy - rh / 2 - lw);
		b.r = Math.max(b.r, cx + rw / 2 + lw);
		b.b = Math.max(b.b, cy + rh / 2 + lw);
	}
	return b;
}

// Exports all shapes (without background) and returns the exported area for cropping
function exportShapes(sl, fn, sngWidth, sngHeight) {
	var resX = parseInt(WScript.arguments(2), 10);
	var resY = parseInt(WScript.arguments(3), 10);
	var shp = sl.Shapes.AddTextBox( 1, 0, 0, sngWidth, sngHeight );
	var shpGroup = sl.Shapes.Range();
	var b = getBounds(shpGroup, sngWidth, sngHeight);

	if (resX === 0) {
		shpGroup.Export(fn, 2, 0, 0, 1);
	} else {
		// The scale is relative to the slide size, so the slide area gets the requested size
		shpGroup.Export(fn, 2, Math.round(resX / 1.33333333), Math.round(resY / 1.33333333), 1);
	}
	shp.Delete();

	var vals = [b.l, b.t, b.r, b.b, sngWidth, sngHeight];
	for (var i = 0; i < vals.length; i++) {
		vals[i] = Math.round(vals[i] * 100);
	}
	return vals.join(" ") + " " + resX + " " + resY;
}

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
		if (sl.SlideShowTransition.AdvanceTime > 0) {
			var objFileToWrite = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/advance.dat",8,true);
			objFileToWrite.WriteLine(sl.SlideIndex + "\t" + sl.SlideShowTransition.AdvanceTime);
			objFileToWrite.Close();
			objFileToWrite = null;
		}

		var objSlideEffect = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/slideEffect.dat",8,true);
		objSlideEffect.WriteLine(sl.SlideIndex + "," + sl.SlideShowTransition.EntryEffect + "," + sl.SlideShowTransition.Duration);
		objSlideEffect.Close();
		objSlideEffect = null;

		fn = WScript.arguments(1) + "/Slide" + sl.SlideIndex + ".png";

		var bounds = exportShapes(sl, fn, sngWidth, sngHeight);

		var fso = new ActiveXObject("Scripting.FileSystemObject");
		if (fso.FileExists(fn)) {
			var objFile = fso.GetFile(fn);
			if (objFile.size === 0) {
				for (var intShape = 1; intShape<=sl.Shapes.Count; intShape++) {
					if (sl.Shapes(intShape).Type === 7) {
						sl.Shapes(intShape).Delete();
					}
				}
				bounds = exportShapes(sl, fn, sngWidth, sngHeight);
			}
		}

		var objBounds = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(WScript.arguments(1) + "/bounds.dat",8,true);
		objBounds.WriteLine(sl.SlideIndex + " " + bounds);
		objBounds.Close();
		objBounds = null;
	}
}

function main() {
	objPPT.DisplayAlerts = false;
	ap = objPPT.Presentations.Open(WScript.arguments(0), false, false, false);
	proc(ap);

	for (var i=0; i< objPPT.Presentations.Count; i++) {
		var opres = objPPT.Presentations.Item(i + 1);
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

cscript.classic.vbsQuickEdit = `
var objPPT;
var file;
var slideNo;
var ap;
objPPT = new ActiveXObject("PowerPoint.Application");
file = WScript.arguments(0);
slideNo = parseInt(WScript.arguments(1));

function main() {
	for ( var i = 0; i < objPPT.Presentations.Count; i++ ) {
		var opres = objPPT.Presentations.Item(i + 1).FullName;
		if ( opres === file ) {
			ap = objPPT.Presentations.Item(i + 1);
			break;
		}
	}

	if ( ! ap ) {
		ap = objPPT.Presentations.Open(file, false);
	}
	ap.Windows(1).Activate();
	ap.Slides(slideNo).Select();
}
main();
`;

module.exports.script = cscript;
