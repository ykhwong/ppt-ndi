const _WORKDIR = process.cwd();
const path = require('path');
const _TMPDIR = path.join(_WORKDIR, "/tmp");
const fs = require("fs-extra");
const sleep = require('system-sleep');
const execSync = require("child_process").execSync;
const spawnSync = require("child_process").spawnSync;
const rimraf = require("rimraf");

const _url = {
	"ndi_sdk": {
		"win32": "https://downloads.ndi.tv/SDK/NDI_SDK/NDI%204%20SDK.exe",
		"linux": "https://downloads.ndi.tv/SDK/NDI_SDK_Linux/InstallNDISDK_v4_Linux.tar.gz",
		"mac": "https://downloads.ndi.tv/SDK/NDI_SDK_Mac/InstallNDISDK_v4_Apple.pkg"
	},
	"innoextract": {
		"win32": "https://constexpr.org/innoextract/files/innoextract-1.9-windows.zip"
	}
};

const _filepath = {
	"ndi_sdk": {
		"win32": [
			"app/Bin/x64/Processing.NDI.Lib.x64.dll",
			"app/Lib/x64/Processing.NDI.Lib.x64.lib",
			"app/Include"
		],
		"linux": [],
		"mac": []
	}
}

function _prepare() {
	if ( ! /^(win32|linux|darwin)$/.test(process.platform) ) {
		console.error("Unknown or unsupported OS type: " + process.platform);
		_exit(1);
	} else {
		// TO-DO: Support linux and darwin
		if ( ! /^win32/.test(process.platform) ) {
			console.error("Unsupported OS type: " + process.platform);
			_exit(1);
		}
	}

	if ( ! fs.existsSync(path.join(_WORKDIR, "src")) ) {
		console.error("Failed to find " + path.join(_WORKDIR, "src"));
		_exit(1);
	}

	if ( ! fs.existsSync(path.join(_WORKDIR, "node_modules")) ) {
		console.error("Failed to locate " + path.join(_WORKDIR, "node_modules"));
		_exit(1);
	}

	try {
		rimraf.sync( _TMPDIR );
		if ( fs.existsSync(_TMPDIR) ) {
			console.error("Failed to remove " + _TMPDIR);
			_exit(1);
		}
		fs.mkdirSync( _TMPDIR, { recursive: true } );
	} catch (e) {
		console.error(e);
		console.error("Failed to remove " + _TMPDIR);
		_exit(1);
	}

	if ( ! fs.existsSync(_TMPDIR) ) {
		console.error("Failed to create " + _TMPDIR);
		_exit(1);
	}
}

function _init() {
	const wget = require('wget-improved');

	process.chdir(_TMPDIR);
	console.log("Downloading NDI SDK...");
	if ( process.platform === "win32" ) {
		let dl1 = wget.download(_url.ndi_sdk.win32, 'ndi_sdk_win32.exe', {});
		let dl2 = wget.download(_url.innoextract.win32, 'innoextract.zip', {});
		let dl1_done = false;
		let dl2_done = false;
		dl1.on('error', function(err) {
			console.error(err);
			_exit(1);
		});

		dl2.on('error', function(err) {
			console.error(err);
			_exit(1);
		});

		dl1.on('end', function(output) {
			dl1_done = true;
		});

		dl2.on('end', function(output) {
			dl2_done = true;
		});

		// 60 sec timeout
		for ( let i = 0; i < 60; i++ ) {
			sleep(1000);
			if (dl1_done && dl2_done) {
				//console.log("done");
				break;
			}
		}
		if ( !dl1_done || !dl2_done ) {
			console.error("Failed to retrieve NDI SDK files");
			_exit(1);
		}
	}
}


function _build() {
	if ( process.platform === "win32" ) {
		const DecompressZip = require("decompress-zip");
		let unzipper = new DecompressZip('innoextract.zip');
		let unzip1_done = false;
		unzipper.on('error', function(err) {
			console.error(err);
			_exit(1);
		});

		unzipper.on('extract', function(err) {
			unzip1_done = true;
		});

		// build PPTNDI lib
		try {
			fs.copySync( path.join(_WORKDIR, "src"), "src" );
		} catch(err) {
			console.error(err);
			_exit(1);
		}

		unzipper.extract({
			filter: function (file) {
				return file.type !== "SymbolicLink";
			}
		});

		// 60 sec timeout
		for ( let i = 0; i < 60; i++ ) {
			sleep(1000);
			if (unzip1_done) {
				//console.log("done");
				break;
			}
		}
		if ( !unzip1_done ) {
			console.error("Failed to unzip innoextract.zip");
			_exit(1);
		}
		
		try {
			execSync("innoextract.exe ndi_sdk_win32.exe");
		} catch(e) {
			console.error(e);
			_exit(1);
		}
		
		try {
			let data = fs.readFileSync("./src/PPTNDI/PPTNDI.cpp", 'utf8');
			let PF86 = process.env["ProgramFiles(x86)"];
			let PF = process.env["ProgramFiles"];
			let realPF;
			let vswhere;
			let out;
			let arr;
			let cmd;
			data = data.replace(/C:\/Program Files\/NewTek\/NDI 4 SDK/g, "./app");
			
			if ( typeof(PF86) === 'undefined' ) {
				realPF = PF86;
			} else {
				realPF = PF;
			}
			vswhere = path.join(PF86, "Microsoft Visual Studio/Installer/vswhere.exe");
			if ( ! fs.existsSync( vswhere ) ) {
				console.error("Visual Studio 15.2 (26418.1 Preview) or higher must be installed");
				_exit(1);
			}
			out = execSync('"' + vswhere.replace(/\//g, "\\") + '"' + " -latest -property installationPath");
			out = path.join(out.toString().replace(/\r|\n/g, ""), "MSBuild", "Current", "Bin", "amd64", "MSBuild.exe");
			
			console.log("Building PPTNDI...");
			cmd = '"' + out + '"' + " ./src/PPTNDI.sln /property:Configuration=Release;Platform=x64 /clp:NoSummary;NoItemAndPropertyList;ErrorsOnly /verbosity:quiet /nologo";
			console.log(cmd);
			out = execSync(cmd);
			console.log(out.toString());
			console.log("Build completed: PPTNDI.dll");
			// final output to ./src/x64/Release/PPTNDI.dll
		} catch(e) {
			console.error(e.stack);
			console.error(e.stderr.toString());
			console.error(e.stdout.toString());
			_exit(1);
		}
		
		// copy the resulting file to deploy dir
		fs.mkdirSync( "deploy/frontend", { recursive: true } );
		fs.mkdirSync( "dev/node_modules", { recursive: true } );
		fs.copySync( path.join(_WORKDIR, "backend.js"), "deploy/backend.js" );
		fs.copySync( path.join(_WORKDIR, "package.json"), "deploy/package.json" );
		fs.copySync( path.join(_WORKDIR, "big_icon.png"), "deploy/big_icon.png" );
		fs.copySync( path.join(_WORKDIR, "icon.ico"), "deploy/icon.ico" );
		fs.copySync( path.join(_WORKDIR, "icon.png"), "deploy/icon.png" );
		fs.copySync( path.join(_WORKDIR, "iconOrig.png"), "deploy/iconOrig.png" );
		fs.copySync( path.join(_WORKDIR, "null_slide.png"), "deploy/null_slide.png" );
		fs.copySync( path.join(_WORKDIR, "frontend"), "deploy/frontend" );

		process.chdir("./deploy");
		fs.copySync( path.join(_WORKDIR, "node_modules"), "node_modules" );		
		fs.copySync("./node_modules/electron", "../dev/node_modules/electron");
		fs.copySync("./node_modules/electron-packager", "../dev/node_modules/electron-packager");
	}
}

function _pack() {
	if ( process.platform === "win32" ) {
		const opt='--icon=./deploy/icon.ico --platform=win32 --overwrite --asar --app-copyright="MIT License (github.com/ykhwong/ppt-ndi)"';
		let ver;
		let out;
		let abi;
		process.chdir( path.join(_WORKDIR, "tmp") );
		try {
			ver = execSync(".\\dev\\node_modules\\electron\\dist\\electron.exe --version").toString().replace(/\r|\n/g, "");
			abi = execSync(".\\dev\\node_modules\\electron\\dist\\electron.exe --abi").toString().replace(/\r|\n/g, "");
			out = execSync("node .\\dev\\node_modules\\electron-packager\\bin\\electron-packager.js ./deploy ppt-ndi --electron-version=" + ver + " " + opt);

			fs.copySync( path.join(".", "src", "x64", "Release", "PPTNDI.dll"), "ppt-ndi-win32-x64/PPTNDI.dll" );
			fs.copySync( path.join(".", "app", "Bin", "x64", "Processing.NDI.Lib.x64.dll"), "ppt-ndi-win32-x64/Processing.NDI.Lib.x64.dll" );
			rimraf.sync( "ppt-ndi-win32-x64/locales" );
			fs.copySync( path.join( _TMPDIR, "deploy", "frontend", "i18n" ), "ppt-ndi-win32-x64/locales" );
			fs.copySync( path.join( _TMPDIR, "deploy", "node_modules", "iohook", "builds", "electron-v" + abi + "-win32-x64",
			"build", "Release", "uiohook.dll"), "ppt-ndi-win32-x64/uiohook.dll" );
		} catch(e) {
			console.error(e.stack);
			console.error(e.stderr.toString());
			console.error(e.stdout.toString());
			_exit(1);
		}
		console.log(out.toString());
	}
}

function _exit(code) {
	process.chdir(_WORKDIR);
	process.exit(code);
}

function _main() {
	_prepare();
	_init();
	_build();
	_pack();
	_exit(0);
}

_main();
