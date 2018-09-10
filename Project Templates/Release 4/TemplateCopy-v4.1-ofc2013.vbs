'Created by IT Department - Muhammad Majid
'muhammad.majid@emiratessteel.com
'TemplateCopy v4.1 for Office 2013

Dim FSO, OSHELL
Set FSO = CreateObject("Scripting.FileSystemObject")
set OSHELL = CreateObject("WScript.Shell") 

Const TemplatesSourceFolder = "\\eisf\SYSVOL\eisf.co.ae\scripts\ES_Templates\Emirates Steel r4"
TemplatesDestinationFolder_Ofc2013 = OSHELL.ExpandEnvironmentStrings("%homepath%") & "\Documents\Custom Office Templates\"

If not FSO.FolderExists(TemplatesDestinationFolder_Ofc2013) Then                    'if templates folder is not there, create it
  FSO.CreateFolder(TemplatesDestinationFolder_Ofc2013)
Else                                                                        'if its there, check for emirates steel and delete it
  If FSO.FolderExists(TemplatesDestinationFolder_Ofc2013+"Emirates Steel\") Then
	FSO.DeleteFolder(TemplatesDestinationFolder_Ofc2013+"Emirates Steel"), True       'remember NO backslash in the end!!
  End If
End If

If FSO.FolderExists(TemplatesSourceFolder) Then
  FSO.CopyFolder TemplatesSourceFolder, TemplatesDestinationFolder_Ofc2013, True    'copy the emirates steel folder
  wscript.echo "Successfully Installed ES Templates Release 4" & vbCRLF & " to " & TemplatesDestinationFolder_Ofc2013
Else
  wscript.echo "Could not find " & TemplatesSourceFolder & vbCRLF & "Please take a screenshot of this error message & call IT on Ext:2444 for help.."
End If