const { remote } = require('electron');
const ipc = require('electron').ipcRenderer;
const fs = require("fs-extra");
const {dialog} = require('electron').remote;
const spawnSync = require( 'child_process' ).spawnSync;
const spawn = require( 'child_process' ).spawn;

var tmp_dir;
const vbs_bg = `
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

const vbs_no_bg = `
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
	Set fso = CreateObject("Scripting.FileSystemObject")

	With ap.PageSetup
		sngWidth = .SlideWidth
		sngHeight = .SlideHeight
	End With

	For Each sl In ap.Slides
		With sl.Shapes.AddShape( 1, 0, 0, sngWidth, sngHeight)
			With .Fill
			.Visible = msoFalse
			End With
			.SetShapesDefaultProperties
			With .Line
			.Visible = msoFalse
			End With
		End With

		Set shpGroup = sl.Shapes.Range()
		Dim fn
		fn = Wscript.Arguments.Item(1) & "/Slide" & sl.SlideIndex & ".png"
		shpGroup.Export fn, 2, , , 1
		Set objFile = fso.GetFile(fn)
		If objFile.Size = 0 Then
			sl.Export fn, "PNG"
		End If

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
	var child = spawn('./bin/PPTNDI');
	var max_slide_num = 0;
	var current_slide = 1;
	var current_window = remote.getCurrentWindow();
	var repo;
	child.stdin.setEncoding('utf-8');
	child.stdout.pipe(process.stdout);

	child.on('exit', function (code) {
		//alert("EXITED " + code);
	});

	function update_screen() {
		var cur_sli, next_sli;
		var next_num;
		var re, rpc;
		if(!repo) {
			return;
		}
		rpc = tmp_dir + "/Slide";
		cur_sli = rpc + current_slide.toString() + '.png';
		next_num = current_slide;
		next_num++;
		next_sli = rpc + next_num.toString() + '.png';
		$("select").find('option[value="Current"]').data('img-src', cur_sli);
		if (!fs.existsSync(next_sli)) {
			next_sli = rpc + '1.png';
		}
		$("select").find('option[value="Next"]').data('img-src', next_sli);
		init_imgpicker();
		try {
			child.stdin.write(cur_sli + "\n");
		} catch(e) {
		}
		$("#slide_cnt").html("SLIDE " + current_slide + " / " + max_slide_num);
	}

	$("select").change(function() {
		if (repo == null) {
			repo = $(this);
		}
	});

	function init_imgpicker() {
		$("select").imagepicker({
			hide_select: true,
			show_label: true,
			selected:function(select, picker_option, event) {
				current_slide=$('.selected').text();
				update_screen();
			}
		});
		if ($("#trans_checker").is(":checked")) {
			$("#right img").css('background-image', "url('trans_slide.png')");
		} else {
			$("#right img").css('background-image', "url('null_slide.png')");
		}
	}

	$("#load_pptx").click(function() {
		dialog.showOpenDialog(current_window,{
			properties: ['openFile'],
			filters: [
				{name: 'PowerPoint Presentations', extensions: ['pptx', 'ppt']},
				{name: 'All Files', extensions: ['*']}
			]
		}, function (file) {
			if (file !== undefined) {
				var re = new RegExp("\\.pptx*\$", "i");
				var vbs_dir, res;
				var file_arr = [];
				var options = "";
				if (re.exec(file)) {
					var now = new Date().getTime();
					var new_vbs_content;
					cleanup_for_temp();
					tmp_dir = process.env.TEMP + '/ppt_ndi';
					if (!fs.existsSync(tmp_dir)) {
						fs.mkdirSync(tmp_dir);
					}
					tmp_dir += '/' + now;
					fs.mkdirSync(tmp_dir);
					vbs_dir = tmp_dir + '/wb.vbs';

					if ($("#with_background").is(":checked")) {
						new_vbs_content = vbs_bg;
					} else {
						new_vbs_content = vbs_no_bg;
					}
					
					try {
						fs.writeFileSync(vbs_dir, new_vbs_content, 'utf-8');
					} catch(e) {
						alert('Failed to access the temporary directory!');
						return;
					}
					res = spawnSync( 'cscript.exe', [ vbs_dir, file, tmp_dir, '' ] );
					if ( res.status !== 0 ) {
						alert('Failed to parse the presentation!');
						return;
					}
					max_slide_num = 0;
					fs.readdirSync(tmp_dir).forEach(file2 => {
						re = new RegExp("^Slide(\\d+)\\.png\$", "i");
						if (re.exec(file2)) {
							var rpc = file2.replace(re, "\$1");
							file_arr.push(rpc);
							max_slide_num++;
						}
					})

					file_arr.sort((a, b) => a - b).forEach(file2 => {
						var rpc = file2;
						options +=
						'<option data-img-label="' + rpc +
						'" data-img-src="' + tmp_dir + '/Slide' + rpc
						+ '.png" value="' + rpc + '">Slide ' + rpc + "\n";
						$("#slides_grp").html(options);
						$("select").find('option[value="Current"]').prop('img-src', tmp_dir + "/Slide1.png");
						if (!fs.existsSync(tmp_dir + "/Slide2.png")) {
							$("select").find('option[value="Next"]').prop('img-src', tmp_dir + "/Slide1.png");
						} else {
							$("select").find('option[value="Next"]').prop('img-src', tmp_dir + "/Slide2.png");
						}
					})
					select_slide('1');
				} else {
					alert("Only allowed filename extensions are PPT and PPTX.");
				}
			}
		})
	});

	function select_slide(num) {
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').prop('selected',true);
		$('optgroup[label="Slides"] option[value="' + num.toString() + '"]').change();
		current_slide = num;

		var selected = $('.selected:eq( 0 )');
		if (selected.length) {
			$("#below").stop().animate(
			{ scrollTop: selected.position().top + $("#below").scrollTop() },
			  500, 'swing', function() {
			  });
		}

		update_screen();
	}

	function goto_prev() {
		var cur_sli;
		var re;
		if (!repo) {
			return;
		}
		cur_sli = current_slide;
		cur_sli--;
		if (cur_sli == 0) {
			cur_sli = max_slide_num;
		}
		select_slide(cur_sli.toString());
	}

	function goto_next() {
		var cur_sli;
		var re;
		if (!repo) {
			return;
		}
		cur_sli = current_slide;
		cur_sli++;
		if (cur_sli > max_slide_num) {
			cur_sli = 1;
		}
		select_slide(cur_sli.toString());
	}
	
	$('#prev').click(function() {
		goto_prev();
	});

	$('#next').click(function() {
		goto_next();
	});

	$(document).keydown(function(e) {
		$("#below").trigger('click');
		if(e.which == 13 || e.which == 32 || e.which == 39 || e.which == 40) {
			// Enter, spacebar, right arrow or down
			goto_next();
		} else if(e.which == 37 || e.which == 8 || e.which == 38) {
			// Left arrow, backspace or up
			goto_prev();
		} else if(e.which == 36) {
			// Home
			select_slide('1');
		} else if(e.which == 35) {
			// End
			select_slide(max_slide_num.toString());
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
			goto_next();
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

	function cleanup_for_temp() {
		if (fs.existsSync(tmp_dir)) {
			fs.removeSync(tmp_dir);
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

	$('#minimize').click(function() {
		remote.BrowserWindow.getFocusedWindow().minimize();
	});

	$('#max_restore').click(function() {
		if(current_window.isMaximized()) {
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

	current_window.on('maximize', function (){
		$("#max_restore").attr("src", "restore.png");
    });

	current_window.on('unmaximize', function (){
		$("#max_restore").attr("src", "max.png");
    });
	
	$('#exit').click(function() {
		cleanup_for_exit();
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

	init_imgpicker();
	startCurrentTime();
});