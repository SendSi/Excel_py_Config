@echo off

set host=192.168.0.239
set user=plan
set psw=123456
set db=wind_basetw_config
set dir=%cd%\wind2_config

cls
%cd%\wind2_config\python\python.exe .\py_export_excel\export_to_db.py %host% %user% %psw% %db% %dir%
@pause