del TCM.ncb
attrib -h TCM.suo
del TCM.suo
del /q TCM\TCM.vcproj.*.user
del /q Debug\TCM.*
del /q TCM\Debug\*
rmdir TCM\Debug
del /q Release\*
rmdir Release
del /q TCM\Release\*
rmdir TCM\Release
del /q /s x64\*
rmdir x64\Release
rmdir x64
del /q /s TCM\x64\*
rmdir TCM\x64\Release
rmdir TCM\x64
rmdir Debug
pause
