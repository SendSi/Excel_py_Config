@echo off

set cur=%cd%
set folder=%cd%\wind2_config
cd ..\LuaScripts\lua_source\config\configlogic
set exportDir=%cd%
cd %cur%

cls
%cd%\wind2_config\python\python.exe .\py_export_excel\export_to_lua.py %folder% %exportDir%
@pause