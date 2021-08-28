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
	let pic = { }; // e.g. { 1: [{ "p:blipFill": ..., "p:nvPicPr": ..., "p:spPr": ... }] }, for zPic items on slide 1
	let relationships = { }; // e.g. { 1: { rId1: "../media/image1.png" } }, for an image on slide 1
	let imageDict = { }; // e.g. { "image1.png": { type: "Buffer", data: [...] } }, strips out the ppt/media/ prefix
	let resSize = {
		resX : 0,
		resY : 0
	};
	let customSize = resSize;
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
					let zPic;
					try {
						num = parseInt(key.replace(/\.xml$/, "").replace(/^.*?(\d+)/, "$1"), 10);
						zData = output[key]["p:sld"]["p:cSld"][0]["p:spTree"][0]; // slide, common slide data, shape tree
						zSp = zData["p:sp"]; // shape (content)
						zNvGrpSpPr = zData["p:nvGrpSpPr"]; // non-visual properties for a group shape (e.g. whether it can be moved, resized, rotated)
						zGrpSpPr = zData["p:grpSpPr"]; // properties for a group shape (e.g. size, location)
						zPic = zData["p:pic"]; // pictures
					} catch (e) {
						notifyError("1");
						return;
					}

					if (zSp) {
						for (let i=0; i<zSp.length; i++) {
							json.push(zSp[i]);
						}
					}
					sp[num] = json; json = [];

					if (zNvGrpSpPr) {
						for (let i=0; i<zNvGrpSpPr.length; i++) {
							json.push(zNvGrpSpPr[i]);
						}
					}
					nvGrpSpPr[num] = json; json = [];

					if (zGrpSpPr) {
						for (let i=0; i<zGrpSpPr.length; i++) {
							json.push(zGrpSpPr[i]);
						}
					}
					grpSpPr[num] = json; json = [];

					if (zPic) {
						for (let i=0; i<zPic.length; i++) {
							json.push(zPic[i]);
						}
					}
					pic[num] = json; json = [];
				}

				// add objects containing Buffer for media to dict
				if (/^ppt\/media\/.*$/.test(key)) {
					try {
						let fileName = key.replace("ppt/media/", "");
						imageDict[fileName] = output[key];
					} catch (e) {
						notifyError("1");
						return;
					}
				}

				// store the image id <-> file name relationship to display images
				if (/^ppt\/slides\/_rels\/slide\d+\.xml.rels$/.test(key)) {
					let json = new Array;
					try {
						num = parseInt(key.replace(/\.xml.rels$/, "").replace(/^.*?(\d+)/, "$1"), 10);
						for (let i=0; i<output[key].Relationships.Relationship.length; i++) {
							let relationship = output[key].Relationships.Relationship[i]["$"];
							json[relationship.Id] = relationship.Target;
						}
					} catch (e) {
						notifyError("1");
						return;
					}

					relationships[num] = json; json = [];
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
		pic = { };
		relationships = { };
		imageDict = { };
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
		if (customSize.resX !== 0 || customSize.resY !== 0) {
			$("#renderer").css({
				"position": "fixed",
				"width": customSize.resX,
				"height": customSize.resY
			});
		} else {
			$("#renderer").css({
				"position": "fixed",
				"width": resSize.resX,
				"height": resSize.resY
			});
		}
		if (isCancelTriggered) {
			notifyCanceled();
			return;
		}

		if (sp[selectedNo] === undefined) {
			notifyError("undefined sp[selectedNo]: " + selectedNo);
			return;
		}

		let blobs = []; // maintain a list to revoke after the slide is captured
		
		// create HTML elements for text
		for (let i=0; i < sp[selectedNo].length; i++) {
			let element = sp[selectedNo][i];
			let elementType = element["p:spPr"][0]["a:prstGeom"][0]["$"]["prst"];
			let elementOffX = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["x"] / baseDiv).toFixed(3);
			let elementOffY = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["y"] / baseDiv).toFixed(3);
			let elementExtCX = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cx"] / baseDiv).toFixed(3);
			let elementExtCY = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cy"] / baseDiv).toFixed(3);
			if (customSize.resX !== 0 || customSize.resY !== 0) {
				// resSizeX : eX = customX : ?
				elementOffX = elementOffX * customSize.resX / resSize.resX;
				elementOffY = elementOffY * customSize.resY / resSize.resY;
				elementExtCX = elementExtCX * customSize.resX / resSize.resX;
				elementExtCY = elementExtCX * customSize.resY / resSize.resY;
			}

			if (elementType === "rect") {
				let txBody = element["p:txBody"];
				let xText = "";
				if (txBody !== null) {
					for (let i2=0; i2<txBody[0]["a:p"].length; i2++) {
						let xFont = "";
						let xFontLatin = "";
						let xFontEa = "";
						let xFontAlgn = "left";
						let xFontSize = "";
						let fontFamily = '';
						let xTextA = '';

						try { xTextA = txBody[0]["a:p"][i2]["a:r"][0]["a:t"][0]; } catch(e) {}
						try { xFont = txBody[0]["a:p"][i2]["a:r"][0]["a:rPr"][0]; } catch(e) {}
						try { xFontLatin = xFont["a:latin"][0]["$"]["typeface"]; } catch(e) {}
						try { xFontEa = xFont["a:ea"][0]["$"]["typeface"]; } catch(e) {}
						try { xFontAlgn = txBody[0]["a:p"][i2]["a:pPr"][0]["$"]["algn"]; } catch(e) {}
						try {
							xFontSize = xFont["$"]["sz"] / 100;
							if (customSize.resX !== 0 || customSize.resY !== 0) {
								xFontSize = xFontSize * customSize.resX / resSize.resX;
							}
						} catch(e) {}
	
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
						xText +=
						'<div style="' +
						(/\S/.test(fontFamily)?'font-family: ' + fontFamily + ";" : '') +
						'text-align: ' + xFontAlgn + ';' +
						'width: ' + elementExtCX +'px;' + 
						'height: ' + elementExtCY +'px;' + 
						'font-size: ' + xFontSize + 'px;' +
						'display: inline;' +
						'white-space: nowrap;' +
						'line-height: 80%;' +
						'"><p>' + xTextA + '</div></p>';
					}

					let rendererConf = 
					'<div style="color: white; position: fixed; ' +
					'left: ' + elementOffX + 'px;' +
					'top: ' + elementOffY + 'px;' + 
					'">' + xText + '</div>';
					if (isCancelTriggered) {
						notifyCanceled();
						return;
					}
					$("#renderer").append(rendererConf);
					}
			}
		}

		// create HTML elements for images
		for (let i=0; i < pic[selectedNo].length; i++) {
			const element = pic[selectedNo][i];
			const elementType = element["p:spPr"][0]["a:prstGeom"][0]["$"]["prst"];
			let elementOffX = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["x"] / baseDiv).toFixed(3);
			let elementOffY = (element["p:spPr"][0]["a:xfrm"][0]["a:off"][0]["$"]["y"] / baseDiv).toFixed(3);
			let elementExtCX = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cx"] / baseDiv).toFixed(3);
			let elementExtCY = (element["p:spPr"][0]["a:xfrm"][0]["a:ext"][0]["$"]["cy"] / baseDiv).toFixed(3);
			if (customSize.resX !== 0 || customSize.resY !== 0) {
				elementOffX = elementOffX * customSize.resX / resSize.resX;
				elementOffY = elementOffY * customSize.resY / resSize.resY;
				elementExtCX = elementExtCX * customSize.resX / resSize.resX;
				elementExtCY = elementExtCX * customSize.resY / resSize.resY;
			}

			if (elementType === "rect") {
				const blipFill = element["p:blipFill"];
				if (!blipFill || blipFill.length === 0) continue;
				const blip = blipFill[0]["a:blip"];
				if (!blip || blip.length === 0) continue;
				const embed = blip[0]["$"]["r:embed"];

				let xText = "";
				if (embed !== null) {
					const imgFilePath = relationships[selectedNo][embed]; // e.g. "../media/image1.jpg"
					const imgFileName = /[^/]*$/.exec(imgFilePath)[0];
					const mimeType = getMimeType(imgFileName);
					if (!mimeType) {
						continue;; // give up on unsupported file types
					}
					const buffer = imageDict[imgFileName].data;

					const img = new Uint8Array(buffer.length);
					for (let i = 0; i < buffer.length; ++i) {
						img[i] = buffer[i];
					}
					const imgUrl = URL.createObjectURL(new Blob([img], { type: mimeType }));
					blobs.push(imgUrl);

					xText +=
					'<img src="' + imgUrl + '" style="' +
					'width: ' + elementExtCX +'px;' + 
					'height: ' + elementExtCY +'px;' + 
					'display: inline;' +
					'white-space: nowrap;' +
					'"/>';
				}

				let rendererConf = 
				'<div style="color: white; position: fixed; ' +
				'left: ' + elementOffX + 'px;' +
				'top: ' + elementOffY + 'px;' + 
				'">' + xText + '</div>';
				if (isCancelTriggered) {
					notifyCanceled();
					return;
				}
				$("#renderer").append(rendererConf);
			}
		}
		
		if (isCancelTriggered) {
			for (let i = 0; i < blobs.length; i += 1) {
				URL.revokeObjectURL(blobs[i]);
			}
			blobs = [];
			notifyCanceled();
			return;
		}
		let options = {};
		if (customSize.resX !== 0 || customSize.resY !== 0) {
			options = {
				width: customSize.resX,
				height: customSize.resY
			}
		}

		htmlToImage.toPng(document.getElementById('renderer'), options)
		.then(function (png) {
			for (let i = 0; i < blobs.length; i += 1) {
				URL.revokeObjectURL(blobs[i]);
			}
			blobs = [];

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

	function getMimeType(fileName) {
		if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) {
			return "image/jpeg";
		} else if (fileName.endsWith(".png")) {
			return "image/png";
		} else {
			return undefined;
		}
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
				customSize = {
					resX : parseInt(data.options.resX),
					resY : parseInt(data.options.resY)
				}
				break;
			case "cancel":
				isCancelTriggered = true;
				break;
		}
	});
});
