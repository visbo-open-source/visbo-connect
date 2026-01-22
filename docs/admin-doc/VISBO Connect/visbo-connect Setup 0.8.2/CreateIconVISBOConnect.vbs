set wshshell = CreateObject("WScript.Shell") 
' Installation von VISBO Connect
Call visboConnectInstall()
' Erstellen des ICON auf Desktop
'Call CreateIcon()
wscript.Quit


' Installation von VISBO Connect
'
Sub visboConnectInstall()
'
WScript.echo "VISBO Connect wird installiert"
currentDir = wshshell.CurrentDirectory
'WScript.echo  currentDir
wshshell.exec currentDir & "\visbo-connect Setup 0.8.2.exe"
wscript.Sleep 4000
End Sub
