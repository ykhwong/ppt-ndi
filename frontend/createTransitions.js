
const JimpRead = require('jimp').read;
let mustStop = false;

function createTransition(buffer, buffer2, tmpDir, i) {
	JimpRead(buffer).then(image => {
		JimpRead(buffer2).then(image2 => {
			//console.log("processing:" + i.toString());
			image.composite(image2, 0, 0, {
				opacitySource: 1 - (0.1 * i),
				opacityDest: 0.1 * i
			}).deflateLevel(0).filterType(0).write(tmpDir + "/t" + i.toString() + ".png", function() {
				//console.log("sent");
				if (!mustStop) {
					self.postMessage("");
				}
			});
		});
	});
}

onmessage = function(e) {
	mustStop = e.data.mustStop;
	if (mustStop) {
		//console.log("mustStop");
		return;
	}
	createTransition(
		Buffer.from(JSON.parse(e.data.buffer).data),
		Buffer.from(JSON.parse(e.data.buffer2).data),
		e.data.tmpDir,
		e.data.i
	);
};
