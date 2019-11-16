const ipc = require('electron').ipcRenderer;

$(document).ready(function() {
	$("#select1img").click(function() {
		ipc.send('remote', { name: "select1" });
	});
	$("#select2img").click(function() {
		ipc.send('remote', { name: "select2" });
	});
	$("#closeImg").click(function() {
		ipc.send('remote', { name: "exit" });
	});
});