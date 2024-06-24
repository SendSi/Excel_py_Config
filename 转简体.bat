@echo off

set dir=%cd%\wind2_config

cls
%cd%\wind2_config\python\python.exe .\py_export_excel\to_t_chinese_simple.py %dir%
@pause