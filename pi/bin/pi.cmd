@ECHO off
REM Pi ¥ patch wrapper — placed early in PATH (~\.pi\agent\bin\pi.cmd)
REM Patches dist files then delegates to the real pi.

REM Run patch hook
powershell -NoProfile -ExecutionPolicy Bypass -File "D:\Program Files\nodejs\node_global\pi-yen-patch.ps1"

REM Delegate to real pi
node "D:\Program Files\nodejs\node_global\node_modules\@earendil-works\pi-coding-agent\dist\cli.js" %*
