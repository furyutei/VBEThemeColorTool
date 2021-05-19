@echo off
pushd "%~dp0"

setlocal
set VBE7DLL="C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
set ThemeXml=".\Themes\VS2017 Dark.xml"
set ForeColors="14 9 12 9 5 4 2 5 5 5"
set BackColors="1 6 1 4 10 1 1 1 1 6"

vbetctool -l %VBE7DLL% -t %ThemeXml% -f %ForeColors% -b %BackColors%

popd
