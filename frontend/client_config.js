$(document).ready(function() {
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const configFile = 'config.js';
	const appDataPath = process.env.APPDATA + "/PPT-NDI";
	const version = "v" + require('electron').remote.app.getVersion();
	const keyCombi = "Ctrl-Shift-";
	const defaultData = {
		"version" : version,
		"startAsTray" : false,
		"startWithTheFirstSlideSelected": false,
		"highPerformance": false,
		"hotKeys" : {
			"prev" : "",
			"next" : "",
			"transparent" : "",
			"black" : "",
			"white" : ""
		},
		"lang": "en"
	};
	let configData = defaultData;
	let configPath = "";

	function alertMsg(myMsg) {
		const { remote } = require('electron');
		const {dialog} = require('electron').remote;
		let currentWindow = remote.getCurrentWindow();
		let options;
		options = {
			type: 'info',
			message: myMsg,
			buttons: ["OK"]
		};
		dialog.showMessageBoxSync(currentWindow, options);
	}

	function filterHotKey(key) {
		return $(key).val().replace(/^.*-/, "");
	}

	function getHotKey(key) {
		return ( key == "" ? "" : keyCombi + key );
	}

	function loadConfig() {
		$.getJSON(configPath, function(json) {
			configData.startAsTray = json.startAsTray;
			configData.startWithTheFirstSlideSelected = json.startWithTheFirstSlideSelected;
			configData.highPerformance = json.highPerformance;
			configData.hotKeys = json.hotKeys;
			configData.lang = json.lang;
			$("#systray").prop('checked', configData.startAsTray);
			$("#startWithFirstSlide").prop('checked', configData.startWithTheFirstSlideSelected);
			$("#highPerformance").prop('checked', configData.highPerformance);
			$("#prevTxtBox").val(getHotKey(configData.hotKeys.prev));
			$("#nextTxtBox").val(getHotKey(configData.hotKeys.next));
			$("#transTxtBox").val(getHotKey(configData.hotKeys.transparent));
			$("#blackTxtBox").val(getHotKey(configData.hotKeys.black));
			$("#whiteTxtBox").val(getHotKey(configData.hotKeys.white));
			if (typeof(configData.lang) === "undefined") {
				$("#langList").val("lang_en");
			} else {
				$("#langList").val("lang_" + configData.lang);
			}
			setLangRsc();
		});
	}

	function setLangRsc() {
		setLangRscDiv("#minimize-systray", "ui_config/minimize-systray", true, configData.lang);
		setLangRscDiv("#start-with-first-slide-selected", "ui_config/start-with-first-slide-selected", true, configData.lang);
		setLangRscDiv("#enable-high-performance-mode", "ui_config/enable-high-performance-mode", true, configData.lang);
		setLangRscDiv("#hotkeys", "ui_config/hotkeys", true, configData.lang);
		setLangRscDiv("#prev", "ui_config/prev", true, configData.lang);
		setLangRscDiv("#next", "ui_config/next", true, configData.lang);
		setLangRscDiv("#black", "ui_config/black", true, configData.lang);
		setLangRscDiv("#white", "ui_config/white", true, configData.lang);
		setLangRscDiv("#trans", "ui_config/transparent", true, configData.lang);
		setLangRscDiv("#localization", "ui_config/localization", true, configData.lang);
		setLangRscDiv("#saveConfig", "ui_config/save", true, configData.lang);
	}

	function setConfig() {
		// Save the config file

		let hotKeys = {
			"prev" : filterHotKey($("#prevTxtBox")),
			"next" : filterHotKey($("#nextTxtBox")),
			"transparent" : filterHotKey($("#transTxtBox")),
			"black" : filterHotKey($("#blackTxtBox")),
			"white" : filterHotKey($("#whiteTxtBox"))
		};

		configData.startAsTray = $("#systray").prop("checked");
		configData.startWithTheFirstSlideSelected = $("#startWithFirstSlide").prop("checked");
		configData.highPerformance = $("#highPerformance").prop("checked");
		configData.hotKeys = hotKeys;
		configData.lang = $("#langList").val().replace(/^lang_/i, "");
		fs.writeFile(configPath, JSON.stringify(configData, null, "\t"), (err) => {
			if (err) {
				alertMsg(getLangRsc("ui_config/could-not-save-config", configData.lang));
			} else {
				ipc.send('remote', { name: "reflectConfig" });
				setLangRsc();
				setTimeout(function() {
					alertMsg(getLangRsc("ui_config/saved", configData.lang));
				}, 100);
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
		ipc.send('remote', { name: "hideConfig" });
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
	$(".txtBox").keydown(function(e) {
		if(e.keyCode === 8) {
			e.preventDefault();
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
			//$(this).focus().val("").val(myVal);
		}
	});
	$("#version").append(version);

	init();
	
	$.each(getLangList(), function (i, item) {
		$("#langList").append($('<option>', { 
			value: "lang_" + item.langCode,
			text : item.details
		}));
	});
});
