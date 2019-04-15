@ECHO OFF
TITLE Pyxcel
SET PYXCEL_HOST=localhost
SET PYXCEL_PORT=5555
start "" "%programfiles(x86)%\Microsoft Office\Office14\EXCEL.EXE" /x ../vba/pyxcel.xla /p %userprofile%
python ../server.py --host %PYXCEL_HOST% --port %PYXCEL_PORT% -m "..\examples\example.py"
