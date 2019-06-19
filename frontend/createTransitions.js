
const Jimp = require('jimp');

function createTransition(buffer, buffer2, tmpDir, i) {
	Jimp.read(buffer).then(image => {
		Jimp.read(buffer2).then(image2 => {
			//console.log("processing:" + i.toString());
			image.composite(image2, 0, 0, {
				opacitySource: 1 - (0.1 * i),
				opacityDest: 0.1 * i
			});
			image.deflateLevel(0).filterType(0);
			image.write(tmpDir + "/t" + i.toString() + ".png", function() {
				//console.log("sent");
				self.postMessage({ "slide" : i });
			});
		});
	});
}

onmessage = function(e) {
	createTransition(
		Buffer.from(JSON.parse(e.data.buffer).data),
		Buffer.from(JSON.parse(e.data.buffer2).data),
		e.data.tmpDir,
		e.data.i
	);
};
