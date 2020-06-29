@echo off
if not exist package.json goto err
echo [Installation]
CMD /C npm install --force --save 2>nul
echo [Rebuild]
electron-rebuild
goto done

:err

:done
