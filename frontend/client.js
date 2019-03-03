var tmpDir = "";
const vbsBg = `
Dim objPPT
Dim TestFile
Dim opres
Set objPPT = CreateObject("PowerPoint.Application")

Sub Proc(ap)
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight

	For Each sl In ap.Slides
		sl.Export Wscript.Arguments.Item(1) & "/Slide" & sl.SlideIndex & ".png", "PNG"
	Next
End Sub

sub Main()
	objPPT.DisplayAlerts = False
	Set ap = objPPT.Presentations.Open(Wscript.Arguments.Item(0), , , msoFalse)
	Proc(ap)

	For each opres In objPPT.Presentations
		TestFile = opres.FullName
		Exit For
	Next

	If TestFile = "" Then objPPT.Quit
	Set objPPT = Nothing
End Sub
Main
`

const vbsNoBg = `
Dim objPPT
Dim TestFile
Dim opres
Set objPPT = CreateObject("PowerPoint.Application")

Sub Proc(ap)
	On Error Resume Next
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight

	With ap.PageSetup
		sngWidth = .SlideWidth
		sngHeight = .SlideHeight
	End With

	For Each sl In ap.Slides
		With sl.Shapes.AddShape( 1, 0, 0, sngWidth, sngHeight)
			.Fill.Visible = msoFalse
			.Line.Visible = msoFalse
			.SetShapesDefaultProperties
		End With

		For intShape = 1 To sl.Shapes.Count
			If sl.Shapes(intShape).Type = 7 Then
				sl.Shapes(intShape).Delete
			End If
		Next

		Set shpGroup = sl.Shapes.Range()
		Dim fn
		fn = Wscript.Arguments.Item(1) & "/Slide" & sl.SlideIndex & ".png"
		shpGroup.Export fn, 2, , , 1

	Next
End Sub

sub Main()
	objPPT.DisplayAlerts = False
	Set ap = objPPT.Presentations.Open(Wscript.Arguments.Item(0), , , msoFalse)
	Proc(ap)

	For each opres In objPPT.Presentations
		TestFile = opres.FullName
		Exit For
	Next

	If TestFile = "" Then objPPT.Quit
	Set objPPT = Nothing
End Sub
Main
`;

$(document).ready(function() {
	const spawn = require( 'child_process' ).spawn;
	const { remote } = require('electron');
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	var child = spawn('./bin/PPTNDI');
	var maxSlideNum = 0;
	var currentSlide = 1;
	var currentWindow = remote.getCurrentWindow();
	var repo;
	child.stdin.setEncoding('utf-8');
	child.stdout.pipe(process.stdout);

	child.on('exit', function (code) {
		//alert("EXITED " + code);
	});

	function updateScreen() {
		var curSli, nextSli;
		var nextNum;
		var re, rpc;
		if(!repo) {
			return;
		}
		rpc = tmpDir + "/Slide";
		curSli = rpc + currentSlide.toString() + '.png';
		nextNum = currentSlide;
		nextNum++;
		nextSli = rpc + nextNum.toString() + '.png';
		$("select").find('option[value="Current"]').data('img-src', curSli);
		if (!fs.existsSync(nextSli)) {
			nextSli = rpc + '1.png';
		}
		$("select").find('option[value="Next"]').data('img-src', nextSli);
		initImgPicker();
		try {
			child.stdin.write(curSli + "\n");
		} catch(e) {
		}
		$("#slide_cnt").html("SLIDE " + currentSlide + " / " + maxSlideNum);
	}

	$("select").change(function() {
		if (repo == null) {
			repo = $(this);
		}
	});

	$("#with_background").click(function() {
		$("#reloadReq").toggle();
	});

	function initImgPicker() {
		$("select").imagepicker({
			hide_select: true,
			show_label: true,
			selected:function(select, picker_option, event) {
				currentSlide=$('.selected').text();
				updateScreen();
			}
		});
		if ($("#trans_checker").is(":checked")) {
			$("#right img").css('background-image', "url('trans_slide.png')");
		} else {
			$("#right img").css('background-image', "url('null_slide.png')");
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
		}, function (file) {
			if (file !== undefined) {
				var re = new RegExp("\\.pptx*\$", "i");
				var vbsDir, res;
				var fileArr = [];
				var options = "";
				if (re.exec(file)) {
					var now = new Date().getTime();
					var newVbsContent;
					var preTmpDir;
					const spawnSync = require( 'child_process' ).spawnSync;
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
						cleanupForTemp();
						tmpDir = preTmpDir;
						alert('Failed to access the temporary directory!');
						$("#fullblack").hide();
						return;
					}
					res = spawnSync( 'cscript.exe', [ vbsDir, file, tmpDir, '' ] );
					$("#fullblack").hide();
					if ( res.status !== 0 ) {
						maxSlideNum = 0;
						cleanupForTemp();
						tmpDir = preTmpDir;
						alert('Failed to parse the presentation!');
						$("#fullblack").hide();
						return;
					}
					maxSlideNum = 0;
					fs.readdirSync(tmpDir).forEach(file2 => {
						re = new RegExp("^Slide(\\d+)\\.png\$", "i");
						if (re.exec(file2)) {
							var rpc = file2.replace(re, "\$1");
							fileArr.push(rpc);
							maxSlideNum++;
						}
					})
					if (fileArr === undefined || fileArr.length == 0) {
						maxSlideNum = 0;
						cleanupForTemp();
						tmpDir = preTmpDir;
						alert("Presentation file could not be loaded.\nPlease remove missing fonts if applicable.");
						$("#fullblack").hide();
						return;
					}

					fileArr.sort((a, b) => a - b).forEach(file2 => {
						var rpc = file2;
						options +=
						'<option data-img-label="' + rpc +
						'" data-img-src="' + tmpDir + '/Slide' + rpc
						+ '.png" value="' + rpc + '">Slide ' + rpc + "\n";
						$("#slides_grp").html(options);
						$("select").find('option[value="Current"]').prop('img-src', tmpDir + "/Slide1.png");
						if (!fs.existsSync(tmpDir + "/Slide2.png")) {
							$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide1.png");
						} else {
							$("select").find('option[value="Next"]').prop('img-src', tmpDir + "/Slide2.png");
						}
					})
					$("#fullblack").hide();
					$("#reloadReq").hide();
					selectSlide('1');
				} else {
					alert("Only allowed filename extensions are PPT and PPTX.");
					$("#fullblack").hide();
				}
			} else {
				$("#fullblack").hide();
			}
		})
	});

	function selectSlide(num) {
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').prop('selected',true);
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').change();
		currentSlide = num;

		var selected = $('.selected:eq( 0 )');
		if (selected.length) {
			$("#below").stop().animate(
			{ scrollTop: selected.position().top + $("#below").scrollTop() },
			  500, 'swing', function() {
			  });
		}

		updateScreen();
	}

	function gotoPrev() {
		var curSli;
		var re;
		if (!repo) {
			return;
		}
		curSli = currentSlide;
		curSli--;
		if (curSli == 0) {
			curSli = maxSlideNum;
		}
		selectSlide(curSli.toString());
	}

	function gotoNext() {
		var curSli;
		var re;
		if (!repo) {
			return;
		}
		curSli = currentSlide;
		curSli++;
		if (curSli > maxSlideNum) {
			curSli = 1;
		}
		selectSlide(curSli.toString());
	}
	
	$('#prev').click(function() {
		gotoPrev();
	});

	$('#next').click(function() {
		gotoNext();
	});

	$(document).keydown(function(e) {
		$("#below").trigger('click');
		if(e.which == 13 || e.which == 32 || e.which == 39 || e.which == 40) {
			// Enter, spacebar, right arrow or down
			gotoNext();
		} else if(e.which == 37 || e.which == 8 || e.which == 38) {
			// Left arrow, backspace or up
			gotoPrev();
		} else if(e.which == 36) {
			// Home
			selectSlide('1');
		} else if(e.which == 35) {
			// End
			selectSlide(maxSlideNum.toString());
		} else if (e.ctrlKey) {
			if (e.which == 87) {
				// Prevents Ctrl-W
				e.preventDefault();
				e.stopPropagation();
			}
		}
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
		var today = new Date();
		var h = today.getHours();
		var m = today.getMinutes();
		var s = today.getSeconds();
		var t;
		m = checkTime(m);
		s = checkTime(s);
		$('#current_time').html(h + ":" + m + ":" + s);
		t = setTimeout(startCurrentTime, 500);
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

	$('#trans_checker').click(function() {
		if ($("#trans_checker").is(":checked")) {
			$("#right img").css('background-image', "url('trans_slide.png')");
		} else {
			$("#right img").css('background-image', "url('null_slide.png')");
		}
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
			e.preventDefault();
		}
	}, false);

	initImgPicker();
	startCurrentTime();
});