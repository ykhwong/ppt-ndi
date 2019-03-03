const ipc = require('electron').ipcRenderer;

$(document).ready(function() {
	$("#select1img").click(function() {
		ipc.send('remote', "select1");
	});
	$("#select2img").click(function() {
		ipc.send('remote', "select2");
	});
	$("#closeImg").click(function() {
		ipc.send('remote', "exit");
	});
});