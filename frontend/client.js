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

Sub Proc()
	Dim ap
	Set ap = objPPT.ActivePresentation
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight

	For Each sl In ap.Slides
		objPPT.ActiveWindow.View.GotoSlide (sl.SlideIndex)
		sl.Export "TEMPPATH_PLACEHOLDER" & "/Slide" & sl.SlideIndex & ".png", "PNG"
	Next
End Sub

sub Main()
	objPPT.DisplayAlerts = False
	With objPPT.Presentations.Open("FILENAME_PLACEHOLDER", False)
	Proc()
	End With

	With objPPT.ActivePresentation
	.Saved = True
	.Close
	End With

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

Sub Proc()
	Dim ap
	Set ap = objPPT.ActivePresentation
	Dim sl
	Dim shGroup
	Dim sngWidth
	Dim sngHeight

	With objPPT.ActivePresentation.PageSetup
		sngWidth = .SlideWidth
		sngHeight = .SlideHeight
	End With

	objPPT.ActiveWindow.ViewType = 1
	For Each sl In ap.Slides
		objPPT.ActiveWindow.View.GotoSlide (sl.SlideIndex)
		sl.Shapes.AddShape( 1, 0, 0, sngWidth, sngHeight).Select
		With objPPT.ActiveWindow.Selection.ShapeRange
			.Fill.Visible = msoTrue
			.Fill.Solid
			.Fill.ForeColor.RGB = RGB(0, 0, 0)
			.Fill.Transparency = 1
			.Line.Visible = msoFalse
		End With

		sl.Shapes.SelectAll
		Set shGroup = objPPT.ActiveWindow.Selection.ShapeRange
		shGroup.Export "TEMPPATH_PLACEHOLDER" & "/Slide" & sl.SlideIndex & ".png", _
							2, , , 1
	Next
End Sub

sub Main()
	objPPT.DisplayAlerts = False
	With objPPT.Presentations.Open("FILENAME_PLACEHOLDER", False)
	Proc()
	End With

	With objPPT.ActivePresentation
	.Saved = True
	.Close
	End With

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
	child.stdin.setEncoding('utf-8');
	child.stdout.pipe(process.stdout);

	function init_imgpicker() {
		$("select").imagepicker({
			hide_select: true,
			show_label  : true,
			selected:function(select, picker_option, event) {
				var cur_sli = $(this).find("option[value='" + $(this).val() + "']").data('img-src');
				var re = new RegExp("^(.*)(\\d+)\\.png\$", "i");
				var rpc = cur_sli.replace(re, "\$1");
				var next_sli = rpc;
				var next_num = parseInt(cur_sli.replace(re, "\$2"), 10);
				next_num++;
				next_sli += next_num.toString() + '.png';
				$("select").find('option[value="Current"]').data('img-src', cur_sli);
				if (!fs.existsSync(next_sli)) {
					next_sli = rpc + '1.png';
				}
				$("select").find('option[value="Next"]').data('img-src', next_sli);
				init_imgpicker();
				child.stdin.write(cur_sli + "\n");
				//child.stdin.end();
			}
		});
	}

	$("#load_pptx").click(function() {
		dialog.showOpenDialog({
			properties: ['openFile'],
			filters: [
			  {name: 'PowerPoint Presentations', extensions: ['pptx']},
			  {name: 'All Files', extensions: ['*']}
			]
		}, function (file) {
			if (file !== undefined) {
				var re = new RegExp("\\.pptx\$", "i");
				var vbs_dir;
				if (re.exec(file)) {
					var now = new Date().getTime();
					var new_vbs_content;
					tmp_dir = process.env.TEMP + '/ppt_ndi';
					if (!fs.existsSync(tmp_dir)) {
						fs.mkdirSync(tmp_dir);
					}
					tmp_dir += '/' + now;
					fs.mkdirSync(tmp_dir);
					vbs_dir = tmp_dir + '/wb.vbs';
					re = new RegExp("FILENAME_PLACEHOLDER", "");
					if ($("#with_background").is(":checked")) {
						new_vbs_content = vbs_bg.replace(re, file);
					} else {
						new_vbs_content = vbs_no_bg.replace(re, file);
					}
					re = new RegExp("TEMPPATH_PLACEHOLDER", "");
					new_vbs_content = new_vbs_content.replace(re, tmp_dir);
					
					try {
						fs.writeFileSync(vbs_dir, new_vbs_content, 'utf-8');
					} catch(e) {
						alert('Failed to access the temporary directory!');
						return;
					}
					var res = spawnSync( 'cscript.exe', [ vbs_dir, '' ] );
					if ( res.status !== 0 ) {
						alert('Failed to parse the presentation!');
						return;
					}
					var options = "";
					fs.readdirSync(tmp_dir).forEach(file2 => {
						re = new RegExp("^Slide(\\d+)\\.png\$", "i");
						if (re.exec(file2)) {
							var rpc = file2.replace(re, "\$1");
							options += '<option data-img-label="' + rpc + '" data-img-src="' + tmp_dir + '/Slide' + rpc + '.png" value="' + rpc + '">Slide ' + rpc + "\n";
							$("#slides_grp").html(options);
							$("select").find('option[value="Current"]').attr('img-src', tmp_dir + "/Slide1.png");
							if (!fs.existsSync(tmp_dir + "/Slide2.png")) {
								$("select").find('option[value="Next"]').attr('img-src', tmp_dir + "/Slide1.png");
							} else {
								$("select").find('option[value="Next"]').attr('img-src', tmp_dir + "/Slide2.png");
							}
							init_imgpicker();
						}
					})

				} else {
					alert("PPTX file is only allowed.");
				}
			}
		})
	});

	$('#exit').click(function() {
	   fs.removeSync(tmp_dir);
       window.close();
	});

	init_imgpicker();
});
