const path = require('path');
const fs = require("fs-extra");
const sleep = require('system-sleep');
const execSync = require("child_process").execSync;
const rimraf = require("rimraf");
const _WORKDIR = process.cwd();
const _TMPDIR = path.join(_WORKDIR, "/tmp");
const license = '"MIT License (github.com/ykhwong/ppt-ndi)"';

const _url = {
	"ndi_sdk": {
		"win32": "https://downloads.ndi.tv/SDK/NDI_SDK/NDI%205%20SDK.exe",
		"linux": "https://downloads.ndi.tv/SDK/NDI_SDK_Linux/Install_NDI_SDK_v5_Linux.tar.gz",
		"darwin": "https://downloads.ndi.tv/SDK/NDI_SDK_Mac/Install_NDI_SDK_v5_macOS.pkg"
	},
	"innoextract": {
		"win32": "https://constexpr.org/innoextract/files/innoextract-1.9-windows.zip"
	},
	"dummyDLL": {
		"win32": "https://github.com/ykhwong/dummy-dll-generator/releases/download/0.2/dummyDLL.exe"
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
		"darwin": [
			"/Library/NDI SDK for macOS/lib/macOS/libndi.dylib"
		]
	}
}

function _prepare() {
	switch (process.platform) {
		case "win32":
		case "darwin":
		case "linux":
			break;
		default:
			console.error("Unknown or unsupported platform: " + process.platform);
			_exit(1);
	}

	switch (process.arch) {
		case "x64":
			break;
		case "x86":
		case "arm64":
			console.error("Unsupported arch type: " + process.arch);
			_exit(1);
		default:
			console.error("Unknown or unsupported arch type: " + process.arch);
			_exit(1);
	}

	if ( ! fs.existsSync(path.join(_WORKDIR, "backend", "src")) ) {
		console.error("Failed to find " + path.join(_WORKDIR, "backend", "src"));
		_exit(1);
	}

	if ( ! fs.existsSync(path.join(_WORKDIR, "node_modules")) ) {
		console.error("Failed to locate " + path.join(_WORKDIR, "node_modules"));
		_exit(1);
	}

	try {
		let pptNdiExecPath = path.join(_TMPDIR, "ppt-ndi-" + process.platform + "-" + process.arch);
		if ( fs.existsSync( _TMPDIR ) ) {
			rimraf.sync(pptNdiExecPath);
			if ( fs.existsSync (pptNdiExecPath) ) {
				console.error("Failed to remove " + pptNdiExecPath);
				_exit(1);
			}
		} else {
			fs.mkdirSync( _TMPDIR, { recursive: true } );
		}
	} catch (e) {
		console.error(e);
		console.error("Failed to locate " + _TMPDIR);
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

	let dl1_done = false;
	let dl2_done = false;
	let dl3_done = false;
	let dl1;
	let dl2;
	let dl3;
	switch (process.platform) {
		case "win32":
			console.log("Downloading NDI SDK...");

			dl1_done = fs.existsSync(path.join(_TMPDIR, 'ndi_sdk_win32.exe'));
			dl2_done = fs.existsSync(path.join(_TMPDIR, 'innoextract.zip'));
			dl3_done = fs.existsSync(path.join(_TMPDIR, 'dummyDLL.exe'));
			if ( ! dl1_done ) {
				dl1 = wget.download(_url.ndi_sdk.win32, 'ndi_sdk_win32.exe', {});
			}
			if ( ! dl2_done ) {
				dl2 = wget.download(_url.innoextract.win32, 'innoextract.zip', {});
			}
			if ( ! dl3_done ) {
				dl3 = wget.download(_url.dummyDLL.win32, 'dummyDLL.exe', {});
			}
			break;
		case "linux":
			console.log("Downloading NDI SDK...");
			dl1_done = fs.existsSync(path.join(_TMPDIR, 'Install_NDI_SDK_v5_Linux.tar.gz'));
			dl2_done = true;
			dl3_done = true;
			if ( ! dl1_done ) {
				dl1 = wget.download(_url.ndi_sdk.linux, 'Install_NDI_SDK_v5_Linux.tar.gz', {});
			}
			break;
		case "darwin":
			// we assume NDI SDK v5 has been installed already on macOS
			dl1_done = true;
			dl2_done = true;
			dl3_done = true;
		default:
			return;
	}

	if (!dl1_done) {
		dl1.on('error', function(err) {
			console.error(err);
			_exit(1);
		});

		dl1.on('end', function(output) {
			dl1_done = true;
		});
	}

	if (!dl2_done && dl2) {
		dl2.on('end', function(output) {
			dl2_done = true;
		});

		dl2.on('error', function(err) {
			console.error(err);
			_exit(1);
		});
	}

	if (!dl3_done && dl3) {
		dl3.on('end', function(output) {
			dl3_done = true;
		});

		dl3.on('error', function(err) {
			console.error(err);
			_exit(1);
		});
	}

	// 1800 sec timeout
	for ( let i = 0; i < 1800; i++ ) {
		if (dl1_done && dl2_done && dl3_done) {
			console.log("Done");
			break;
		}
		sleep(1000);
	}
	if ( !dl1_done || !dl2_done || !dl3_done ) {
		console.error("Failed to retrieve NDI SDK files");
		_exit(1);
	}
}

function _build() {
	switch ( process.platform ) {
		case 'win32':
			_buildWin32();
			break;
		case 'darwin':
			_buildDarwin();
			break;
		case 'linux':
			_buildLinux();
			break;
		default:
			break;
	}
}

function _buildWin32() {
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
		fs.copySync( path.join(_WORKDIR, "backend", "src"), "src" );
	} catch(err) {
		console.error(err);
		_exit(1);
	}

	unzipper.extract({
		filter: function (file) {
			return file.type !== "SymbolicLink";
		}
	});

	// 1800 sec timeout
	for ( let i = 0; i < 1800; i++ ) {
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
		let latestMscVersion;
		let platformToolset;
		let searchDir;
		let searchDirs;
		data = data.replace(/C:\/Program Files\/NewTek\/NDI 4 SDK/g, _TMPDIR.replace(/\\/g, "/") + "/app");
		fs.writeFileSync("./src/PPTNDI/PPTNDI.cpp", '\ufeff' + data, { encoding: 'utf8' });
		
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
		vswhere = vswhere.replace(/\//g, "\\")
		out = execSync('"' + vswhere + '"' + " -latest -property installationPath");
		if ( ! /\S/.test( out ) ) {
			out = execSync('"' + vswhere + '"' + " -prerelease -property installationPath");
		}
		if ( ! /\S/.test( out ) ) {
			console.error("Could not locate the Visual Studio");
			_exit(1);
		}
		out = out.toString().replace(/\r|\n/g, "");

		searchDir = path.join(out, "MSBuild", "Microsoft", "VC");
		searchDirs = getDirs(searchDir.replace(/\//g, "\\"));
		latestMscVersion = searchDirs[searchDirs.length -1];
		if ( !/\S/.test(latestMscVersion) ) {
			console.error("Could not locate the latest MSC version");
			_exit(1);
		}
		searchDir = path.join(out, "MSBuild", "Microsoft", "VC", latestMscVersion, "Platforms", "x64", "PlatformToolsets");
		searchDirs = getDirs(searchDir.replace(/\//g, "\\"));
		platformToolset = searchDirs[searchDirs.length -1];
		if ( !/\S/.test(platformToolset) ) {
			console.error("Could not locate the latest toolset");
			_exit(1);
		}

		console.log("Building PPTNDI...");
		out = path.join(out, "MSBuild", "Current", "Bin", "amd64", "MSBuild.exe");
		cmd = '"' + out + '"' + " ./src/PPTNDI.sln /property:Configuration=Release;Platform=x64 /clp:NoSummary;NoItemAndPropertyList;ErrorsOnly /verbosity:quiet /nologo /p:PlatformToolset=" + platformToolset;
		console.log(cmd);
		out = execSync(cmd);
		console.log(out.toString());
		console.log("Build completed: PPTNDI.dll");
		// final output to ./src/x64/Release/PPTNDI.dll
	} catch(e) {
		console.error(e.stack);
		if (e.stderr) console.error(e.stderr.toString());
		if (e.stdout) console.error(e.stdout.toString());
		_exit(1);
	}
		
	// copy the resulting file to deploy dir
	if ( fs.existsSync("deploy") ) {
		fs.rmdirSync("deploy", { recursive: true });
	}
	if ( fs.existsSync("dev") ) {
		fs.rmdirSync("dev", { recursive: true });
	}
	fs.mkdirSync( "deploy/frontend", { recursive: true } );
	fs.mkdirSync( "deploy/backend/img", { recursive: true } );
	fs.mkdirSync( "dev/node_modules", { recursive: true } );
	fs.copySync( path.join(_WORKDIR, "backend", "backend.js"), "deploy/backend/backend.js" );
	fs.copySync( path.join(_WORKDIR, "package.json"), "deploy/package.json" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.ico"), "deploy/backend/img/icon.ico" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.png"), "deploy/backend/img/icon.png" );
	fs.copySync( path.join(_WORKDIR, "frontend"), "deploy/frontend" );

	process.chdir("./deploy");
	fs.copySync( path.join(_WORKDIR, "node_modules"), "node_modules" );		
	fs.copySync("./node_modules/electron", "../dev/node_modules/electron");
	fs.copySync("./node_modules/electron-packager", "../dev/node_modules/electron-packager");
}

function _buildDarwin() {
	// build PPTNDI lib
	try {
		fs.copySync( path.join(_WORKDIR, "backend", "src"), "src" );
	} catch(err) {
		console.error(err);
		_exit(1);
	}
	try {
		console.log("Building PPTNDI...");
		cmd = 'g++ -dynamiclib -o ./src/PPTNDI.dylib ./src/PPTNDI/PPTNDI.cpp "/Library/NDI SDK for macOS/lib/macOS/libndi.dylib" -rpath .'
		console.log(cmd);
		out = execSync(cmd);
		console.log(out.toString());
		console.log("Build completed: PPTNDI.dylib");
		// final output to ./src/PPTNDI.dylib
	} catch(e) {
		console.error(e.stack);
		if (e.stderr) console.error(e.stderr.toString());
		if (e.stdout) console.error(e.stdout.toString());
		_exit(1);
	}
		
	// copy the resulting file to deploy dir
	if ( fs.existsSync("deploy") ) {
		fs.rmdirSync("deploy", { recursive: true });
	}
	if ( fs.existsSync("dev") ) {
		fs.rmdirSync("dev", { recursive: true });
	}
	fs.mkdirSync( "deploy/frontend", { recursive: true } );
	fs.mkdirSync( "deploy/backend/img", { recursive: true } );
	fs.mkdirSync( "dev/node_modules", { recursive: true } );
	fs.copySync( path.join(_WORKDIR, "backend", "backend.js"), "deploy/backend/backend.js" );
	fs.copySync( path.join(_WORKDIR, "package.json"), "deploy/package.json" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.icns"), "deploy/backend/img/icon.icns" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.png"), "deploy/backend/img/icon.png" );
	fs.copySync( path.join(_WORKDIR, "frontend"), "deploy/frontend" );

	process.chdir("./deploy");
	fs.copySync( path.join(_WORKDIR, "node_modules"), "node_modules" );		
	fs.copySync("./node_modules/electron", "../dev/node_modules/electron");
	fs.copySync("./node_modules/electron-packager", "../dev/node_modules/electron-packager");
}

function _buildLinux() {
	// build PPTNDI lib
	try {
		fs.copySync( path.join(_WORKDIR, "backend", "src"), "src" );
		if ( ! fs.existsSync("NDI-SDK") ) {
			execSync('tar -xzf Install_NDI_SDK_v5_Linux.tar.gz');
			execSync('echo y | sh Install_NDI_SDK_v5_Linux.sh 1>/dev/null 2>/dev/null');
			fs.renameSync( 'NDI SDK for Linux', 'NDI-SDK' );
		}
	} catch(err) {
		console.error(err);
		_exit(1);
	}
	try {
		console.log("Building PPTNDI...");
		cmd = 'g++ -shared -s -fPIC -o ./src/libpptndi.so ./src/PPTNDI/PPTNDI.cpp -L "./NDI-SDK/lib/x86_64-linux-gnu" -l:libndi.so.5.1.1'
		console.log(cmd);
		out = execSync(cmd);
		console.log(out.toString());
		console.log("Build completed: libpptndi.so");
		// final output to ./src/libpptndi.so
	} catch(e) {
		console.error(e.stack);
		if (e.stderr) console.error(e.stderr.toString());
		if (e.stdout) console.error(e.stdout.toString());
		_exit(1);
	}

	// copy the resulting file to deploy dir
	if ( fs.existsSync("deploy") ) {
		fs.rmdirSync("deploy", { recursive: true });
	}
	if ( fs.existsSync("dev") ) {
		fs.rmdirSync("dev", { recursive: true });
	}
	fs.mkdirSync( "deploy/frontend", { recursive: true } );
	fs.mkdirSync( "deploy/backend/img", { recursive: true } );
	fs.mkdirSync( "dev/node_modules", { recursive: true } );
	fs.copySync( path.join(_WORKDIR, "backend", "backend.js"), "deploy/backend/backend.js" );
	fs.copySync( path.join(_WORKDIR, "package.json"), "deploy/package.json" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.icns"), "deploy/backend/img/icon.icns" );
	fs.copySync( path.join(_WORKDIR, "backend", "img", "icon.png"), "deploy/backend/img/icon.png" );
	fs.copySync( path.join(_WORKDIR, "frontend"), "deploy/frontend" );

	process.chdir("./deploy");
	fs.copySync( path.join(_WORKDIR, "node_modules"), "node_modules" );
	fs.copySync("./node_modules/electron", "../dev/node_modules/electron");
	fs.copySync("./node_modules/electron-packager", "../dev/node_modules/electron-packager");
}

function _pack() {
	process.chdir( path.join(_WORKDIR, "tmp") );
	let ver;
	let out;
	let abi;

	if ( process.platform === "win32" ) {
		const opt='--icon=./deploy/backend/img/icon.ico --platform=win32 --overwrite --asar --app-copyright=' + license;
		try {
			ver = execSync(".\\dev\\node_modules\\electron\\dist\\electron.exe --version").toString().replace(/\r|\n/g, "");
			abi = execSync(".\\dev\\node_modules\\electron\\dist\\electron.exe --abi").toString().replace(/\r|\n/g, "");
			out = execSync("node .\\dev\\node_modules\\electron-packager\\bin\\electron-packager.js ./deploy ppt-ndi --electron-version=" + ver + " " + opt);

			fs.copySync( path.join(".", "src", "x64", "Release", "PPTNDI.dll"), "ppt-ndi-win32-x64/PPTNDI.dll" );
			fs.copySync( path.join(".", "app", "Bin", "x64", "Processing.NDI.Lib.x64.dll"), "ppt-ndi-win32-x64/Processing.NDI.Lib.x64.dll" );
			rimraf.sync( "ppt-ndi-win32-x64/locales" );
			fs.copySync( path.join( _TMPDIR, "deploy", "frontend", "i18n" ), "ppt-ndi-win32-x64/locales" );
		} catch(e) {
			console.error(e.stack);
			if (e.stderr) console.error(e.stderr.toString());
			if (e.stdout) console.error(e.stdout.toString());
			_exit(1);
		}
		try {
			execSync("dummyDLL.exe ./ppt-ndi-win32-x64/d3dcompiler_47.dll");
			fs.copySync( "out.dll", "ppt-ndi-win32-x64/d3dcompiler_47.dll" );
			execSync("dummyDLL.exe ./ppt-ndi-win32-x64/ffmpeg.dll");
			fs.copySync( "out.dll", "ppt-ndi-win32-x64/ffmpeg.dll" );
			execSync("dummyDLL.exe ./ppt-ndi-win32-x64/vk_swiftshader.dll");
			fs.copySync( "out.dll", "ppt-ndi-win32-x64/vk_swiftshader.dll" );
			execSync("dummyDLL.exe ./ppt-ndi-win32-x64/vulkan-1.dll");
			fs.copySync( "out.dll", "ppt-ndi-win32-x64/vulkan-1.dll" );
		} catch(e) {
			console.error(e);
			_exit(1);
		}
		console.log(out.toString());
	} else if ( process.platform === "darwin" ) {
		const opt='--icon=./deploy/backend/img/icon.icns --platform=darwin --overwrite --app-copyright=' + license;
		try {
			ver = execSync("./dev/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron --version").toString().replace(/\r|\n/g, "");
			abi = execSync("./dev/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron --abi").toString().replace(/\r|\n/g, "");
			
			fs.copySync( path.join(".", "src", "PPTNDI.dylib"), "deploy/PPTNDI.dylib" );
			fs.copySync( "/Library/NDI SDK for macOS/lib/macOS/libndi.dylib", "deploy/libndi.dylib" );
			rimraf.sync( "deploy/locales" );
			fs.copySync( path.join( _TMPDIR, "deploy", "frontend", "i18n" ), "deploy/locales" );

			out = execSync("node dev/node_modules/electron-packager/bin/electron-packager.js ./deploy ppt-ndi --electron-version=" + ver + " " + opt);
			process.chdir( path.join(_WORKDIR, "tmp", "ppt-ndi-darwin-x64", "ppt-ndi.app", "Contents", "Frameworks" ) );
			fs.symlinkSync("../../Contents/Resources/app/libndi.dylib", "libndi.dylib");
			process.chdir( path.join(_WORKDIR, "tmp") );
		} catch(e) {
			console.error(e.stack);
			if (e.stderr) console.error(e.stderr.toString());
			if (e.stdout) console.error(e.stdout.toString());
			_exit(1);
		}
		console.log(out.toString());
	} else if ( process.platform === "linux" ) {
		const opt='--platform=linux --overwrite --app-copyright=' + license;
		const data = '#!/bin/sh' + "\n" + './ppt-ndi-core --no-sandbox $@' + "\n";
		try {
			ver = execSync("./dev/node_modules/electron/dist/electron --no-sandbox --version").toString().replace(/\r|\n/g, "");
			abi = execSync("./dev/node_modules/electron/dist/electron --no-sandbox --abi").toString().replace(/\r|\n/g, "");
			
			fs.copySync( path.join(".", "src", "libpptndi.so"), "deploy/libpptndi.so" );
			rimraf.sync( "deploy/locales" );
			fs.copySync( path.join( _TMPDIR, "deploy", "frontend", "i18n" ), "deploy/locales" );

			out = execSync("node dev/node_modules/electron-packager/bin/electron-packager.js ./deploy ppt-ndi --electron-version=" + ver + " " + opt);
			fs.copySync( "./NDI-SDK/lib/x86_64-linux-gnu/libndi.so.5.1.1", "ppt-ndi-linux-x64/libndi.so.5" );
			fs.renameSync( './ppt-ndi-linux-x64/ppt-ndi', './ppt-ndi-linux-x64/ppt-ndi-core' );
			fs.writeFileSync( './ppt-ndi-linux-x64/ppt-ndi', data, { encoding: 'utf8' });
			fs.chmodSync('./ppt-ndi-linux-x64/ppt-ndi', "755");
		} catch(e) {
			console.error(e.stack);
			if (e.stderr) console.error(e.stderr.toString());
			if (e.stdout) console.error(e.stdout.toString());
			_exit(1);
		}
		console.log(out.toString());

	}
}

function getDirs(path) {
  return fs.readdirSync(path).filter(function (file) {
    return fs.statSync(path+'/'+file).isDirectory();
  });
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
