'Created by IT Department - Muhammad Majid
'muhammad.majid@esi-steel.com

Dim FSO, OSHELL
Set FSO = CreateObject("Scripting.FileSystemObject")
set OSHELL = CreateObject("WScript.Shell") 

'Const TemplatesSourceFolder = "\\eisf.co.ae\SysVol\eisf.co.ae\scripts\Templates\Emirates Steel"
'TemplatesDestinationFolder = OSHELL.ExpandEnvironmentStrings("%appdata%") & "\Microsoft\Templates\"

'If not FSO.FolderExists(TemplatesDestinationFolder) Then
'FSO.CreateFolder(TemplatesDestinationFolder)
'End If

'wscript.echo "this is " & TemplatesDestinationFolder
'wscript.echo "Searching for " & TemplatesSourceFolder

'If FSO.FolderExists(TemplatesSourceFolder) Then
'  wscript.echo "Found " & TemplatesSourceFolder
'  FSO.CopyFolder TemplatesSourceFolder, TemplatesDestinationFolder, True
'  wscript.echo "Successfully Installed " & TemplatesSourceFolder & " to " & TemplatesDestinationFolder
'  wscript.echo "Successfully Installed Template to:" & vbCRLF & TemplatesDestinationFolder
'Else
'  wscript.echo "Could not find " & TemplatesSourceFolder & vbCRLF & "Please call IT on Ext:2444 for help.."
'End If
  wscript.echo "This is an outdated version of ES templates." & vbCRLF & "Please install the revised templates using IT Public Advertised Email on Oct 9, 2012" & vbCRLF & " Or call IT on Ext:2444 for help.."