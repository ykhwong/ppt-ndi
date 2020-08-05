$(document).ready(function() {
	const PPTXCompose = require("pptx-compose").default;
	const iconv = require('iconv-lite');
	const composer = new PPTXCompose();
	const baseDiv = 9525;
	const htmlToImage = require('html-to-image');
	const fs = require("fs-extra");
	let sp = { };
	let nvGrpSpPr = { };
	let grpSpPr = { };
	let resSize = {
		resX : 0,
		resY : 0
	};
	let isLoaded = false;

	// x, y = location
	// cx, cy = width, height

	function loadPPT(filename) {
		cleanUpJSON();
		composer.toJSON(filename).then((output) => {
			output = JSON.parse(iconv.decode(JSON.stringify(output), 'utf-8'));
			let sldSz = output["ppt/presentation.xml"]["p:presentation"]["p:sldSz"][0]["$"];
			let slideCnt = 0;
			resSize.resX = sldSz.cx / baseDiv;
			resSize.resY = sldSz.cy / baseDiv;

			for (let key in output) {
				if (/^ppt\/slides\/slide\d+\.xml$/.test(key)) {
					let num = parseInt(key.replace(/\.xml$/, "").replace(/^.*(\d+)/, "$1"), 10);
					let zData = output[key]["p:sld"]["p:cSld"][0]["p:spTree"][0];
					let zSp = zData["p:sp"];
					let zNvGrpSpPr = zData["p:nvGrpSpPr"];
					let zGrpSpPr = zData["p:grpSpPr"];
					let json = new Array;

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
			//console.log(output);

			slideCnt = Object.keys(sp).length;
			drawSlide(1, slideCnt);
		});
	}
	
	function cleanUpJSON() {
		sp = { };
		nvGrpSpPr = { };
		grpSpPr = { };
		isLoaded = false;
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
					$("#renderer").append(rendererConf);
					htmlToImage.toPng(document.getElementById('renderer'))
					.then(function (png) {
						png = png.replace(/^data:image\/png;base64,/, "");
						//fs.writeFileSync("Slide" + selectedNo + ".png", png, 'base64');
						if (selectedNo === slideCnt) {
							isLoaded = true;
							return;
						}
						drawSlide(selectedNo + 1, slideCnt);
					});
				}
			}
		}
	}

	function getSlideSize() {
		return resSize;
	}

	//loadPPT("d:/sandbox/simple.pptx");

});
