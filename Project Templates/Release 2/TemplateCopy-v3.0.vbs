'Created by IT Department - Muhammad Majid
'muhammad.majid@emiratessteel.com

Dim FSO, OSHELL
Set FSO = CreateObject("Scripting.FileSystemObject")
set OSHELL = CreateObject("WScript.Shell") 

Const TemplatesSourceFolder = "\\eisf.co.ae\SysVol\eisf.co.ae\scripts\Templates\Emirates Steel"
TemplatesDestinationFolder = OSHELL.ExpandEnvironmentStrings("%appdata%") & "\Microsoft\Templates\"

If not FSO.FolderExists(TemplatesDestinationFolder) Then
  'wscript.echo TemplatesDestinationFolder & vbCRLF & "does not exist"
  FSO.CreateFolder(TemplatesDestinationFolder)
  'wscript.echo TemplatesDestinationFolder & vbCRLF & "created!"
End If

If FSO.FolderExists(TemplatesSourceFolder) Then
  'wscript.echo "Found " & TemplatesSourceFolder
  FSO.CopyFolder TemplatesSourceFolder, TemplatesDestinationFolder, True
  wscript.echo "Successfully Installed ES Templates Revision 2.0" & vbCRLF & " to " & TemplatesDestinationFolder
Else
  wscript.echo "Could not find " & TemplatesSourceFolder & vbCRLF & "Please call IT on Ext:2444 for help.."
End If