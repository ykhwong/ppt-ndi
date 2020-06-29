@echo off
if not exist package.json goto err
echo [Installation]
CMD /C npm install --force --save 2>nul
echo [Rebuild]
call electron-rebuild
modclean --patterns="default:*"
goto done

:err

:done
