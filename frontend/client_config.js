$(document).ready(function() {
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	const configFile = 'config.js';
	const appDataPath = (process.env.APPDATA || (process.platform == 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + "/.local/share")) + "/PPT-NDI";
	const version = "v" + require('electron').remote.app.getVersion();
	const keyCombi = "Ctrl-Shift-";
	let rendererList;
	if (process.platform === 'darwin') {
		rendererList = ["Internal"];
	} else { // win32
		rendererList = ["Microsoft PowerPoint", "Internal"];
	}
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
		"renderer": rendererList[0],
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
			if (json) {
				configData.startAsTray = json.startAsTray;
				configData.startWithTheFirstSlideSelected = json.startWithTheFirstSlideSelected;
				configData.highPerformance = json.highPerformance;
				configData.hotKeys = json.hotKeys;
				configData.lang = json.lang;
				configData.renderer = json.renderer;
				$("#systray").prop('checked', configData.startAsTray);
				$("#startWithFirstSlide").prop('checked', configData.startWithTheFirstSlideSelected);
				$("#highPerformance").prop('checked', configData.highPerformance);
				$("#prevTxtBox").val(getHotKey(configData.hotKeys.prev));
				$("#nextTxtBox").val(getHotKey(configData.hotKeys.next));
				$("#transTxtBox").val(getHotKey(configData.hotKeys.transparent));
				$("#blackTxtBox").val(getHotKey(configData.hotKeys.black));
				$("#whiteTxtBox").val(getHotKey(configData.hotKeys.white));
				if (typeof(configData.lang) === "undefined" || !/\S/.test(configData.lang)) {
					configData.lang = "en";
				}
				$("#rendererList").val(configData.renderer);
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
		setLangRscDiv("#renderer", "ui_config/renderer", true, configData.lang);
		setLangRscDiv("#saveConfig", "ui_config/save", true, configData.lang);
	}

	function setConfig(showInfo, useDefaultData = false) {
		// Save the config file

		if (!useDefaultData) {

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
			configData.renderer = $("#rendererList").val();
		}
		
		if (!fs.existsSync(appDataPath)) {
			fs.mkdirSync(appDataPath, {
				recursive: false // ~/Library/Preferences should already exist
			});
		}
		fs.writeFile(configPath, JSON.stringify(configData, null, "\t"), { flag: 'w' }, (err) => {
			if (err) {
				alertMsg(getLangRsc("ui_config/could-not-save-config", configData.lang));
			} else {
				ipc.send('remote', { name: "reflectConfig" });
				setLangRsc();
				if (showInfo) {
					setTimeout(function() {
						alertMsg(getLangRsc("ui_config/saved", configData.lang));
					}, 100);
				}
			}
		});
	}

	function init() {
		const { remote } = require('electron');
		$.ajaxSetup({
			async: false
		});
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		if (fs.existsSync(configPath)) {
			loadConfig();
		} else {
			configPath = appDataPath + "/" + configFile;
			if (fs.existsSync(configPath)) {
				loadConfig();
				setConfig(false);
			} else {
				setConfig(false, true);
			}
		}
	}

	$('#closeImg').click(function() {
		ipc.send('remote', { name: "hideConfig" });
	});
	$('#saveConfig').click(function() {
		setConfig(true);
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

	$.each(getLangList(), function (i, item) {
		$("#langList").append($('<option>', { 
			value: "lang_" + item.langCode,
			text : item.details
		}));
	});

	for (let i=0; i<rendererList.length; i++) {
		$("#rendererList").append($('<option>', { 
			value: rendererList[i],
			text : rendererList[i]
		}));	
	}

	init();
});
