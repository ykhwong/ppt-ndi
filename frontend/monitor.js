$(document).ready(function() {
	const { remote } = require('electron');
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	var isTransSet = false;

	$("html").css({
		"background-color": "transparent",
		"width": "100%",
		"height": "100%"
	});

	$("#monitorDisp").css({
		"position": "fixed",
		"top": "0",
		"left": "0",
		"width": "100%",
		"height": "100%",
		"background-color": "transparent",
		"background-repeat": "no-repeat",
		"background-size": "100% 100%"
	});

	$("html").keydown(function(e) {
		if(e.keyCode === 27) {
			let currentWindow = remote.getCurrentWindow();
			currentWindow.hide();
		}
	});

	ipc.on('monitor', (event, data) => {

		function updateMon() {
			var mode = data.mode;
			var workerinit = data.workerinit;
			var usebg = data.modeusebg;
			var dFile = data.file.replace(/\\/g, "/");
			var dFile2 = dFile;
			if (/\/mode2\/Slide\d+\.png/i.test(dFile2)) {
				dFile2 = dFile2.replace(/\/mode2(\/Slide\d+\.png)/i, "$1");
			} else {
				dFile2 = dFile2.replace(/(\/Slide\d+\.png)/i, "/mode2/$1");
			}

			$("#html, #monitorDisp").css("background-color", isTransSet ? "transparent" : "black");
			if (!workerinit || /Slide0\.png$/i.test(dFile)) {
				$("#monitorDisp").css({
					"background-image": 'url(file:///' + dFile + ')'
				});
			} else {
				if (isTransSet) {
					if (usebg) {
						$("#monitorDisp").css({
							"background-image": 'url(file:///' + dFile + ')'
						});
					} else {
						$("#monitorDisp").css({
							"background-image": 'url(file:///' + dFile2 + ')'
						});
					}
				} else {
					if (usebg) {
						$("#monitorDisp").css({
							"background-image": 'url(file:///' + dFile2 + ')'
						});
					} else {
						$("#monitorDisp").css({
							"background-image": 'url(file:///' + dFile + ')'
						});
					}
				}
			}
		}

		switch (data.func) {
			case "update":
				updateMon();
				break;
			case "transparentOn":
				$("#html, #monitorDisp").css("background-color", "transparent");
				isTransSet = true;
				updateMon();
				break;
			case "transparentOff":
				$("#html, #monitorDisp").css("background-color", "black");
				isTransSet = false;
				updateMon();
				break;
			case "monitorBlack":
				$("#html, #monitorDisp").css("background-color", "black");
				$("#monitorDisp").css({
					"background-image": 'none'
				});
				break;
			case "monitorWhite":
				$("#html, #monitorDisp").css("background-color", "white");
				$("#monitorDisp").css({
					"background-image": 'none'
				});
				break;
			case "monitorTrans":
				$("#html, #monitorDisp").css("background-color", "transparent");
				$("#monitorDisp").css({
					"background-image": 'none'
				});
				break;
			default:
				break;
		}
	});
});
