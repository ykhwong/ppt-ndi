$(document).ready(function() {
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const configFile = 'config.js';
	const appDataPath = process.env.APPDATA + "/PPT-NDI";
	const version = "20190702a";
	const keyCombi = "Ctrl-Shift-";
	const defaultData = {
		"version" : version,
		"startAsTray" : false,
		"hotKeys" : {
			"prev" : "",
			"next" : "",
			"transparent" : "",
			"black" : "",
			"white" : ""
		}
	};
	let configData = defaultData;
	let configPath = "";

	function filterHotKey(key) {
		return $(key).val().replace(/^.*-/, "");
	}

	function getHotKey(key) {
		return ( key == "" ? "" : keyCombi + key );
	}

	function loadConfig() {
		$.getJSON(configPath, function(json) {
			configData.startAsTray = json.startAsTray;
			configData.hotKeys = json.hotKeys;
			$("#systray").prop('checked', configData.startAsTray);
			$("#prevTxtBox").val(getHotKey(configData.hotKeys.prev));
			$("#nextTxtBox").val(getHotKey(configData.hotKeys.next));
			$("#transTxtBox").val(getHotKey(configData.hotKeys.transparent));
			$("#blackTxtBox").val(getHotKey(configData.hotKeys.black));
			$("#whiteTxtBox").val(getHotKey(configData.hotKeys.white));
		});
	}

	function setConfig() {
		// Save the config file

		// TO-DO: reflect the config to mainwindow2
		let hotKeys = {
			"prev" : filterHotKey($("#prevTxtBox")),
			"next" : filterHotKey($("#nextTxtBox")),
			"transparent" : filterHotKey($("#transTxtBox")),
			"black" : filterHotKey($("#blackTxtBox")),
			"white" : filterHotKey($("#whiteTxtBox"))
		};

		configData.startAsTray = $("#systray").prop("checked");
		configData.hotKeys = hotKeys;
		fs.writeFile(configPath, JSON.stringify(configData, null, "\t"), (err) => {
			if (err) {
				alert("Could not save the configuration.");
			} else {
				ipc.send('remote', "reflectConfig");
			}
		});
	}

	function init() {
		const { remote } = require('electron');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		if (fs.existsSync(configPath)) {
			loadConfig();
		} else {
			configPath = appDataPath + "/" + configFile;
			if (fs.existsSync(configPath)) {
				loadConfig();
			} else {
				// Do nothing
			}
		}
	}

	$('#closeImg').click(function() {
		ipc.send('remote', "hideConfig");
	});
	$('#saveConfig').click(function() {
		setConfig();
	});

	$(".txtBox").on("click",function(){
		if ($(".txtBox").is(':focus')) {
			let myVal = $(this).val();
			$(this).focus().val("").val(myVal);
		}
	});
	$(".txtBox").keyup(function(e) {
		let myVal = $(this).val();
		let chr = String.fromCharCode( e.keyCode );
		if (e.keyCode >= 48 && e.keyCode <= 57) {
			// 0 - 9
			$(this).focus().val("").val(keyCombi + chr);
		} else if (e.keyCode >= 96 && e.keyCode <= 105) {
			// 0 - 9 (numpad)
			$(this).focus().val(keyCombi + (e.keyCode-96).toString());
		} else if (e.keyCode >= 65 && e.keyCode <= 90) {
			// a - z
			$(this).focus().val(keyCombi + chr);
		} else if (e.keyCode == 8 || e.keyCode == 46) {
			$(this).focus().val("");
		} else {
			$(this).focus().val("").val(myVal);
		}
	});
	init();
});
