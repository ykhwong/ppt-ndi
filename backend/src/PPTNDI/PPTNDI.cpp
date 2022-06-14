#include <stdlib.h>
#include <string>

using namespace std;

#ifdef _WIN32
#define EXPORT comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)
#include <windows.h>

#ifdef _WIN64
	#pragma comment(lib, "C:/Program Files/NewTek/NDI 4 SDK/Lib/x64/Processing.NDI.Lib.x64.lib")
#else // _WIN64
	#pragma comment(lib, "C:/Program Files/NewTek/NDI 4 SDK/Lib/x86/Processing.NDI.Lib.x86.lib")
#endif // _WIN64

#include "C:/Program Files/NewTek/NDI 4 SDK/Examples/C++/NDIlib_Send_PNG/picopng.hpp"
#include <C:/Program Files/NewTek/NDI 4 SDK/Include/Processing.NDI.Lib.h>
#elif __APPLE__
#include <stdio.h>
#include <stdlib.h>
#include <dlfcn.h>

#include "/Library/NDI SDK for macOS/examples/C++/NDIlib_Send_PNG/picopng.hpp"
#include </Library/NDI SDK for macOS/include/Processing.NDI.Lib.h>
#else
#include <stdio.h>
#include <stdlib.h>
#include "../../NDI-SDK/examples/C++/NDIlib_Send_PNG/picopng.hpp"
#include "../../NDI-SDK/include/Processing.NDI.Lib.h"
#endif

bool initSucceeded = false;
NDIlib_send_create_t NDI_send_create_desc;
NDIlib_send_instance_t pNDI_send;
const int maxInstance = 5;
const char *ndiName = "PPTNDI";

#if __APPLE__
extern "C"
#endif
int init(void) {
	#pragma EXPORT
	NDIlib_initialize();
	
	NDI_send_create_desc.p_ndi_name = ndiName;
	pNDI_send = NDIlib_send_create(&NDI_send_create_desc);
	if (!pNDI_send) {
		for (int i = 2; i <= maxInstance; i++) {
			char buffer[15];
			#ifdef _WIN32
				sprintf_s(buffer, "%s (%d)", ndiName, i);
			#elif __APPLE__
				snprintf(buffer, 15, "%s (%d)", ndiName, i);
			#endif
			NDI_send_create_desc.p_ndi_name = buffer;
			pNDI_send = NDIlib_send_create(&NDI_send_create_desc);

			if (!pNDI_send) {
				if (i == maxInstance) {
					return 1;
				}
			} else {
				break;
			}
		}
	}
	initSucceeded = true;

	return 0;
}

#if __APPLE__
extern "C"
#endif
int destroy(void) {
	#pragma EXPORT
	if (!initSucceeded) {
		return 1;
	}
	NDIlib_send_destroy(pNDI_send);
	NDIlib_destroy();
	return 0;
}

#if __APPLE__
extern "C"
#endif
int send(const char *path, bool trans) {
	#pragma EXPORT
	NDIlib_video_frame_v2_t NDI_video_frame;
	vector<unsigned char> png_data;
	vector<unsigned char> image_data;
	string str(path);

	if (!initSucceeded) {
		return 1;
	}
	loadFile(png_data, str);
	if (png_data.empty()) {
		return 1;
	}

	unsigned long xres = 0, yres = 0;
	if (decodePNG(image_data, xres, yres, &png_data[0], png_data.size(), true)) {
		return 1;
	}

	NDI_video_frame.xres = xres;
	NDI_video_frame.yres = yres;
	NDI_video_frame.FourCC = NDIlib_FourCC_type_RGBA;
	NDI_video_frame.p_data = &image_data[0];
	NDI_video_frame.line_stride_in_bytes = xres * 4;

	for (int i = 1; i <= (trans?1:2); ++i) {
		NDIlib_send_send_video_v2(pNDI_send, &NDI_video_frame);
	}
	return 0;
}
