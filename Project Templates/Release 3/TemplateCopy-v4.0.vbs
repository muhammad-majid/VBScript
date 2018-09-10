'Created by IT Department - Muhammad Majid
'muhammad.majid@emiratessteel.com

Dim FSO, OSHELL
Set FSO = CreateObject("Scripting.FileSystemObject")
set OSHELL = CreateObject("WScript.Shell") 

Const TemplatesSourceFolder = "\\eisf.co.ae\SysVol\eisf.co.ae\scripts\Templates\Emirates Steel"
TemplatesDestinationFolder = OSHELL.ExpandEnvironmentStrings("%appdata%") & "\Microsoft\Templates\"

If not FSO.FolderExists(TemplatesDestinationFolder) Then                    'if templates folder is not there, create it
  'wscript.echo TemplatesDestinationFolder & vbCRLF & "does not exist"
  FSO.CreateFolder(TemplatesDestinationFolder)
  'wscript.echo TemplatesDestinationFolder & vbCRLF & "created!"
Else                                                                        'if its there, check for emirates steel and delete it
  If FSO.FolderExists(TemplatesDestinationFolder+"Emirates Steel\") Then
  'wscript.echo TemplatesDestinationFolder+"Emirates Steel\" & vbCRLF & "already exists"
  FSO.DeleteFolder(TemplatesDestinationFolder+"Emirates Steel"), True       'remember NO backslash in the end!!
  'wscript.echo TemplatesDestinationFolder+"Emirates Steel\" & vbCRLF & "deleted"
  End If
End If

If FSO.FolderExists(TemplatesSourceFolder) Then
  'wscript.echo "Found " & TemplatesSourceFolder
  FSO.CopyFolder TemplatesSourceFolder, TemplatesDestinationFolder, True    'copy the emirates steel folder
  wscript.echo "Successfully Installed ES Templates Revision 3.0" & vbCRLF & " to " & TemplatesDestinationFolder
Else
  wscript.echo "Could not find " & TemplatesSourceFolder & vbCRLF & "Please take a screenshot of this error message & call IT on Ext:2444 for help.."
End If