$(document).ready(function() {
	const ipc = require('electron').ipcRenderer;
	const fs = require("fs-extra");
	let configData = {
		"lang": "en"
	};
	$("#select1img").click(function() {
		ipc.send('remote', { name: "select1" });
	});
	$("#select2img").click(function() {
		ipc.send('remote', { name: "select2" });
	});
	$("#closeImg").click(function() {
		ipc.send('remote', { name: "exit" });
	});

	function reflectConfig() {
		const configFile = 'config.js';
		let configPath = "";
		const remote = require('@electron/remote');
		configPath = remote.app.getAppPath().replace(/(\\|\/)resources(\\|\/)app\.asar/, "") + "/" + configFile;
		if (!fs.existsSync(configPath)) {
			const appDataPath = (process.env.APPDATA || (process.platform === 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + "/.local/share")) + "/PPT-NDI";
			configPath = appDataPath + "/" + configFile;
		}
		if (fs.existsSync(configPath)) {
			$.getJSON(configPath, function(json) {
				configData.lang = json.lang;
				setLangRsc();
				//ipc.send('remote', { name: "passConfigData", details: configData });
			});
		}
	}

	function setLangRsc() {
		setLangRscDiv("#select1", "ui_main/select1", true, configData.lang);
		setLangRscDiv("#select2", "ui_main/select2", true, configData.lang);
	}

	ipc.on('remote' , function(event, data){
		switch (data.msg) {
			case "reload":
				reflectConfig();
				break;
		}
	});
	
	reflectConfig();
});
