// Crops an image exported via ShapeRange.Export back to the slide area.
// ShapeRange.Export renders the bounding box of all shapes, so shapes that
// stick out of the slide enlarge the image and break its aspect ratio.
//
// bounds: [left, top, right, bottom, slideWidth, slideHeight, targetX, targetY]
//   - the first six values are in 1/100 pt (integers, to avoid locale issues)
//   - targetX/targetY are the requested output resolution in px (0 = auto)
const fs = require('fs');
const PNG = require('pngjs').PNG;

function parseBounds(arr) {
	if (!arr || arr.length < 8) {
		return null;
	}
	const v = arr.map((x) => parseFloat(x));
	if (v.some((x) => isNaN(x))) {
		return null;
	}
	return {
		left: v[0] / 100, top: v[1] / 100, right: v[2] / 100, bottom: v[3] / 100,
		slideW: v[4] / 100, slideH: v[5] / 100,
		targetX: v[6], targetY: v[7]
	};
}

function hasOverflow(b) {
	const eps = 0.5;
	return b.left < -eps || b.top < -eps ||
		b.right > b.slideW + eps || b.bottom > b.slideH + eps;
}

function resize(src, sw, sh, dw, dh) {
	// bilinear on premultiplied alpha to avoid dark fringes on transparent edges
	const dst = Buffer.alloc(dw * dh * 4);
	const rx = sw / dw;
	const ry = sh / dh;
	for (let y = 0; y < dh; y++) {
		let fy = (y + 0.5) * ry - 0.5;
		if (fy < 0) fy = 0;
		const y0 = Math.min(Math.floor(fy), sh - 1);
		const y1 = Math.min(y0 + 1, sh - 1);
		const wy = fy - y0;
		for (let x = 0; x < dw; x++) {
			let fx = (x + 0.5) * rx - 0.5;
			if (fx < 0) fx = 0;
			const x0 = Math.min(Math.floor(fx), sw - 1);
			const x1 = Math.min(x0 + 1, sw - 1);
			const wx = fx - x0;
			const idx = [(y0 * sw + x0) << 2, (y0 * sw + x1) << 2, (y1 * sw + x0) << 2, (y1 * sw + x1) << 2];
			const w = [(1 - wx) * (1 - wy), wx * (1 - wy), (1 - wx) * wy, wx * wy];
			let a = 0, r = 0, g = 0, b = 0;
			for (let k = 0; k < 4; k++) {
				const pa = src[idx[k] + 3] * w[k];
				a += pa;
				r += src[idx[k]] * pa;
				g += src[idx[k] + 1] * pa;
				b += src[idx[k] + 2] * pa;
			}
			const o = (y * dw + x) << 2;
			if (a > 0) {
				dst[o] = Math.round(r / a);
				dst[o + 1] = Math.round(g / a);
				dst[o + 2] = Math.round(b / a);
			}
			dst[o + 3] = Math.round(a);
		}
	}
	return dst;
}

function cropToSlide(file, boundsArr) {
	const b = parseBounds(boundsArr);
	if (!b || !hasOverflow(b) || !fs.existsSync(file)) {
		return false;
	}
	let png;
	try {
		png = PNG.sync.read(fs.readFileSync(file));
	} catch (e) {
		return false;
	}
	const bw = b.right - b.left;
	const bh = b.bottom - b.top;
	if (bw <= 0 || bh <= 0) {
		return false;
	}
	// px per pt, estimated from the image. The bounds may be off by a few px
	// (antialiasing, effects), so the output size is snapped below.
	const sx = png.width / bw;
	const sy = png.height / bh;
	let outW, outH;
	if (b.targetX > 0 && b.targetY > 0) {
		outW = b.targetX;
		outH = b.targetY;
	} else {
		// Default export resolution is a whole DPI (usually 96 dpi = 4/3 px per pt)
		const scale = Math.round((sx + sy) / 2 * 72) / 72;
		outW = Math.round(b.slideW * scale);
		outH = Math.round(b.slideH * scale);
	}
	let cw = Math.round(b.slideW * sx);
	let ch = Math.round(b.slideH * sy);
	if (Math.abs(cw - outW) <= 2 && Math.abs(ch - outH) <= 2) {
		// close enough: plain crop without resampling keeps the image sharp
		cw = outW;
		ch = outH;
	}
	const x0 = Math.round(-b.left * sx);
	const y0 = Math.round(-b.top * sy);
	if (cw <= 0 || ch <= 0) {
		return false;
	}

	// copy the slide area; anything outside the exported image stays transparent
	let data = Buffer.alloc(cw * ch * 4);
	const sxFrom = Math.max(0, x0);
	const sxTo = Math.min(png.width, x0 + cw);
	if (sxTo > sxFrom) {
		for (let y = 0; y < ch; y++) {
			const yy = y0 + y;
			if (yy < 0 || yy >= png.height) {
				continue;
			}
			const s = (yy * png.width + sxFrom) << 2;
			png.data.copy(data, (y * cw + (sxFrom - x0)) << 2, s, s + ((sxTo - sxFrom) << 2));
		}
	}

	if (cw !== outW || ch !== outH) {
		data = resize(data, cw, ch, outW, outH);
	}

	const out = new PNG({ width: outW, height: outH });
	data.copy(out.data);
	fs.writeFileSync(file, PNG.sync.write(out));
	return true;
}

// Reads <dir>/bounds.dat ("<slideIndex> <8 bound values>" per line) written by
// the classic export script and crops each <dir>/Slide<n>.png accordingly.
function cropSlidesInDir(dir) {
	const boundsFile = dir + "/bounds.dat";
	if (!fs.existsSync(boundsFile)) {
		return;
	}
	const lines = fs.readFileSync(boundsFile, { encoding: 'utf8' }).split(/\r\n|\n/);
	for (let i = 0; i < lines.length; i++) {
		const v = lines[i].trim().split(/\s+/);
		if (v.length < 9 || !/^\d+$/.test(v[0])) {
			continue;
		}
		cropToSlide(dir + "/Slide" + v[0] + ".png", v.slice(1, 9));
	}
}

module.exports.cropToSlide = cropToSlide;
module.exports.cropSlidesInDir = cropSlidesInDir;
