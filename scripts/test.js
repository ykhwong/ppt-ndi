const _WORKDIR = process.cwd();
const path = require('path');
const _TMPDIR = path.join(_WORKDIR, "/tmp");
const fs = require("fs-extra");
const execSync = require("child_process").execSync;

function _exit(code) {
	process.chdir(_WORKDIR);
	process.exit(code);
}

function _main() {
	if ( ! /^(win32|linux|darwin)$/.test(process.platform) ) {
		console.error("Unknown or unsupported OS type: " + process.platform);
		_exit(1);
	} else {
		// TO-DO: Support linux
		if ( ! /^(win32|darwin)/.test(process.platform) ) {
			console.error("Unsupported OS type: " + process.platform);
			_exit(1);
		}
	}

	let dirPath = {
		binPath: "",
		execBin: ""
	};

	switch ( process.platform ) {
		case "win32":
			dirPath.binPath = path.join(_TMPDIR, "ppt-ndi-win32-x64");
			dirPath.execBin = "ppt-ndi.exe";
			break;
		case "darwin":
			dirPath.binPath = path.join(_TMPDIR, "ppt-ndi-darwin-x64");
			dirPath.execBin = "ppt-ndi";
			break;
		default:
			break;
	}

	let fullBinPath = path.join(dirPath.binPath, dirPath.execBin);
	if ( ! fs.existsSync(dirPath.binPath) ) {
		console.error("Failed to find " + fullBinPath);
		_exit(1);
	}
	process.chdir(dirPath.binPath);
	execSync(fullBinPath);
}

_main();
