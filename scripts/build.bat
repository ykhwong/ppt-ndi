@echo off
set VERINFO=5.0.8
set OPT=--icon=.\icon.ico --platform=win32 --overwrite --asar --app-copyright="MIT License (github.com/ykhwong/ppt-ndi)"

call electron-packager ./ppt-ndi ppt-ndi --electron-version=%VERINFO% %OPT%
IF %ERRORLEVEL% == 1 goto done
cd ppt-ndi-win32-x64
del /q d3dcompiler_47.dll
rmdir /s /q .\swiftshader
rmdir /s /q .\locales
rmdir /s /q .\resources\inspector
del /q LICENSE
del /q LICENSES.chromium.html
del /q osmesa.dll 2>nul
del /q libEGL.dll
del /q libGLESv2.dll
del /q chrome_*.pak
del /q snapshot_blob.bin
del /q version
mkdir bin
copy ..\PPTNDI.EXE .\bin\.
copy ..\Processing.NDI.Lib.x64.dll .\.
copy ..\uiohook.dll .\.
copy ..\*_slide.png .\resources\.
copy ..\sample.pptx .\resources\.
:done
pause >nul
