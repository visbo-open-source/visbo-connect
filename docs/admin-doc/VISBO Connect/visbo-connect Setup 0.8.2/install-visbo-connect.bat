cls
@pushd %~dp0
dir \excel.exe /s /b > befehl.txt
set /P locEXCEL=<befehl.txt
set locEXCEL=%locEXCEL:\=\\%
set speSHEET=c:\\VISBO\\VISBO SPE\\VISBO Project Edit.xlsx
set a= {"excelExe"
set b= %a%: "%locEXCEL%","speSheet": "%speSheet%"}
echo %b% > %AppData%\VISBO\visbo-connect\vcn_config.json
if errorlevel 1 goto fehler3
cscript "CreateIconVISBOConnect.vbs"
goto ende
:fehler3
echo Probleme beim Installieren von visbo-connect
:ende
@popd