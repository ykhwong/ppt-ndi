$(document).ready(function() {
	const PPTXCompose = require("pptx-compose").default;
	const iconv = require('iconv-lite');
	const composer = new PPTXCompose();
	const baseDiv = 9525;
	const htmlToImage = require('html-to-image');
	const fs = require("fs-extra");
	const ipc = require('electron').ipcRenderer;
	let sp = { };
	let nvGrpSpPr = { };
	let grpSpPr = { };
	let resSize = {
		resX : 0,
		resY : 0
	};
	let isLoaded = false;
	let hasError = false;
	let outPath = "";
	let pptFile = "";
	let isCancelTriggered = false;

	// x, y = location
	// cx, cy = width, height

	function loadPPT(filename, outDir) {
		cleanUpJSON();
		pptFile = filename;
		outPath = outDir;
		composer.toJSON(filename).then((output) => {
			output = JSON.parse(iconv.decode(JSON.stringify(output), 'utf-8'));
			let slideCnt = 0;
			let sldSz;
			let xmlFound = false;
			try {
				sldSz = output["ppt/presentation.xml"]["p:presentation"]["p:sldSz"][0]["$"];
			} catch (e) {
				notifyError("Unknown p:sldSz");
				return;
			}
			resSize.resX = sldSz.cx / baseDiv;
			resSize.resY = sldSz.cy / baseDiv;

			for (let key in output) {
				if (/^ppt\/slides\/slide\d+\.xml$/.test(key)) {
					xmlFound = true;
					let json = new Array;
					let num;
					let zData;
					let zSp;
					let zNvGrpSpPr;
					let zGrpSpPr;
					try {
						num = parseInt(key.replace(/\.xml$/, "").replace(/^.*(\d+)/, "$1"), 10);
						zData = output[key]["p:sld"]["p:cSld"][0]["p:spTree"][0];
						zSp = zData["p:sp"];
						zNvGrpSpPr = zData["p:nvGrpSpPr"];
						zGrpSpPr = zData["p:grpSpPr"];
					} catch (e) {
						notifyError("1");
						return;
					}

					for (let i=0; i<zSp.length; i++) {
							json.push(zSp[i]);
					}
					sp[num] = json; json = [];

					for (let i=0; i<zNvGrpSpPr.length; i++) {
							json.push(zNvGrpSpPr[i]);
					}
					nvGrpSpPr[num] = json; json = [];

					for (let i=0; i<zGrpSpPr.length; i++) {
							json.push(zGrpSpPr[i]);
					}
					grpSpPr[num] = json; json = [];
				}
			}
			if (!xmlFound) {
				notifyError("2");
				return;
			}

			slideCnt = Object.keys(sp).length;
			drawSlide(1, slideCnt);
		});
	}
	
	function cleanUpJSON() {
		sp = { };
		nvGrpSpPr = { };
		grpSpPr = { };
		isLoaded = false;
		hasError = false;
		outPath = "";
		resSize = {
			resX : 0,
			resY : 0
		};
	}

	function drawSlide(selectedNo, slideCnt) {
		$("#renderer").html("");
		$("#renderer").css({
			"position": "fixed",
			"width": resSize.resX,
			"height": resSize.resY
		});
		if (isCancelTriggered) {
			notifyCanceled();
			return;
		}

		if (sp[selectedNo] === undefined) {
			notifyError("undefined sp[selectedNo]: " + selectedNo);
			return;
		}
		for (let i=0; i < sp[selectedNo].length; i++) {
			let element = sp[selectedNo][i];
			let elementType = element["p:spPr"][0]["a:prstGeom"][0]["$"]["prst"];
			let elementOffX = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["x"] / baseDiv).toFixed(3);
			let elementOffY = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["y"] / baseDiv).toFixed(3);
			let elementExtCX = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cx"] / baseDiv).toFixed(3);
			let elementExtCY = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cy"] / baseDiv).toFixed(3);
			if (elementType === "rect") {
				let txBody = element["p:txBody"];
				if (txBody !== null) {
					let xText = "";
					let xFont = "";
					let xFontLatin = "";
					let xFontEa = "";
					let xFontAlgn = "left";
					let xFontSize = "";

					try { xText = txBody[0]["a:p"][0]["a:r"][0]["a:t"][0]; } catch(e) {}
					try { xFont = txBody[0]["a:p"][0]["a:r"][0]["a:rPr"][0]; } catch(e) {}
					try { xFontLatin = xFont["a:latin"][0]["$"]["typeface"]; } catch(e) {}
					try { xFontEa = xFont["a:ea"][0]["$"]["typeface"]; } catch(e) {}
					try { xFontAlgn = txBody[0]["a:p"][0]["a:pPr"][0]["$"]["algn"]; } catch(e) {}
					try { xFontSize = xFont["$"]["sz"] / 100; } catch(e) {}

					let fontFamily = '';
					if (/\S/.test(xFontLatin)) {
						fontFamily = "'" + xFontLatin + "'";
					}
					if (/\S/.test(xFontAlgn)) {
						if (xFontAlgn === 'ctr') {
							xFontAlgn = "center";
						} else if (xFontAlgn === 'r') {
							xFontAlgn = "right";
						} else if (xFontAlgn === 'just') {
							xFontAlgn = "justify";
						} else {
							xFontAlgn = "left";
						}
					}
					if (/\S/.test(xFontEa)) {
						if (/\S/.test(fontFamily)) {
							fontFamily += ",";
						}
						fontFamily += "'" + xFontEa + "'";
					}
					let rendererConf = 
						'<div style="color: white; position: fixed; ' +
						'left: ' + elementOffX + 'px;' +
						'top: ' + elementOffY + 'px;' + 
						(/\S/.test(fontFamily)?'font-family: ' + fontFamily + ";" : '') +
						'text-align: ' + xFontAlgn + ';' +
						'width: ' + elementExtCX +'px;' + 
						'height: ' + elementExtCY +'px;' + 
						'font-size: ' + xFontSize + 'px;' +
						'white-space: nowrap;' +
						'">' + xText + '</div>';
					//console.log(rendererConf);
					if (isCancelTriggered) {
						notifyCanceled();
						return;
					}
					$("#renderer").append(rendererConf);
				}
			}
		}
		if (isCancelTriggered) {
			notifyCanceled();
			return;
		}
		htmlToImage.toPng(document.getElementById('renderer'))
		.then(function (png) {
			if (isCancelTriggered) {
				notifyCanceled();
				return;
			}
			png = png.replace(/^data:image\/png;base64,/, "");
			fs.writeFileSync(outPath + "/Slide" + selectedNo + ".png", png, 'base64');
			console.log(">>" + selectedNo);
			if (isCancelTriggered) {
				notifyCanceled();
				return;
			}
			if (selectedNo === slideCnt) {
				if (isCancelTriggered) {
					notifyCanceled();
				} else {
					notifyLoaded();
				}
			} else {
				if (isCancelTriggered) {
					notifyCanceled();
					return;
				} else {
					selectedNo++;
					drawSlide(selectedNo, slideCnt);	
				}
			}
		});
	}

	function notifyError(msg) {
		hasError = true;
		isLoaded = false;
		isCancelTriggered = false;
		console.log("renderer : error");
		ipc.send("renderer", { name: "notifyError", message: msg });
	}

	function notifyLoaded() {
		hasError = false;
		isLoaded = true;
		isCancelTriggered = false;
		console.log("renderer : loaded");
		ipc.send("renderer", {
			name: "notifyLoaded",
			outDir: outPath,
			pptFile: pptFile
		});
	}

	function notifyCanceled() {
		hasError = false;
		isLoaded = true;
		if (isCancelTriggered) {
			console.log("renderer : cancelled");
			ipc.send("renderer", {
				name: "notifyCanceled"
			});
		}
		isCancelTriggered = false;
	}

	function getSlideSize() {
		return resSize;
	}

	function _isLoaded() {
		return isLoaded;
	}

	function _hasError() {
		return hasError;
	}

	ipc.on('renderer' , function(event, data){
		switch (data.func) {
			case "load":
				loadPPT(data.options.file, data.options.outDir);
				break;
			case "cancel":
				isCancelTriggered = true;
				break;
		}
	});
});
