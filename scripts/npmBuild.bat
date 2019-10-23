@echo off
if not exist package.json goto err
echo [Installation]
CMD /C npm install --force --save 2>nul
rmdir /s /q .\node_modules\ffi\node_modules\ref 2>nul
rmdir /s /q .\node_modules\ffi\node_modules\ref-struct 2>nul
rmdir /s /q .\node_modules\ref-struct\node_modules\ref 2>nul
mklink /J .\node_modules\ffi\node_modules\ref .\node_modules\ref
mklink /J .\node_modules\ffi\node_modules\ref-struct .\node_modules\ref-struct
mklink /J .\node_modules\ref-struct\node_modules\ref .\node_modules\ref
echo [Rebuild]
electron-rebuild
goto done

:err

:done
