@echo off
cd /d %~dp0 
regsvr32 /S dll\file_controlv2.dll
regsvr32 /S dll\mathv3.dll
regsvr32 /S dll\Windows_FormApi.dll
regsvr32 /S dll\RC6.dll
start /wait dll\runtime_install.exe 
pause
exit