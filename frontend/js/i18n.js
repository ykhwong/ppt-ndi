$(document).ready(function() {
	const fs = require("fs-extra");
	const path = require('path');

	const getPath = (namespace) => {
		let tmpPath = path.join(__dirname, 'i18n', namespace + ".json");
		if (!fs.existsSync(tmpPath)) {
			tmpPath = path.join(__dirname, '..', 'locales', namespace + ".json");
		}
		return tmpPath;
	}

	getLangRsc=function(title, curLang) {
		const fullTitle = title.split("/");
		const namespace = fullTitle[0];
		const sTitle = fullTitle[1];
		const jsonPath = getPath(namespace);
		let result = null;
		if (!/\S/.test(namespace) || !/\S/.test(sTitle)) {
			return null;
		}
		if (fs.existsSync(jsonPath)) {
			const rsc = fs.readFileSync(jsonPath, { encoding: 'utf8' });
			const fullObj = $.parseJSON(rsc);
			for ( let i = 0; i < fullObj.length; i++ ) {
				let obj = fullObj[i];
				if (obj["item"] === sTitle) {
					result = obj.lang[curLang].msg;
					break;
				}
			}
		}
		if (!result) {
			return title;
		}
		return result;
	}
	
	getLangList=function() {
		const jsonPath = getPath("info");
		const rsc = fs.readFileSync(jsonPath, { encoding: 'utf8' });
		const fullObj = $.parseJSON(rsc);
		const langList = fullObj.enabledLangs;
		let result = [];
		for ( let i = 0; i < langList.length; i++ ) {
			const obj = langList[i];
			const details = fullObj.langDetails[obj];
			const val = details.display_name + " - " + details.native_name;
			const newObj = {
				"langCode" : obj,
				"details" : val
			};
			result.push(newObj);
		}
		return result;
	}
	
	setLangRscDiv=function(div, rscName, nbsp, curLang) {
		$(div).html((nbsp?"&nbsp;":"") + getLangRsc(rscName, curLang));
	}	
});
