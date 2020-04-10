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

	ipc.on('monitor', (event, data) => {
		switch (data.func) {
			case "update":
				var dFile = data.file.replace(/\\/g, "/");
				$("#html, #monitorDisp").css("background-color", isTransSet ? "transparent" : "black");
				$("#monitorDisp").css({
					"background-image": 'url(file:///' + dFile + ')'
				});
				break;
			case "transparentOn":
				$("#html, #monitorDisp").css("background-color", "transparent");
				isTransSet = true;
				break;
			case "transparentOff":
				$("#html, #monitorDisp").css("background-color", "black");
				isTransSet = false;
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
