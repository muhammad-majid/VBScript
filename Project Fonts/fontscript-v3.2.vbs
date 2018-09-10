Option Explicit
' Installing multiple Fonts in Windows 7
' http://www.cloudtec.ch 2011
'muhammad.majid@emiratessteel.com

Dim objShell, objFSO, wshShell
Dim objFontInstallationFileLog, objFontInstallationFileLog2, objFontInstallationSuccessFlag, objFontsSourceFolder, objFontsDestinationFolder, objFont, objNameSpace, objFile

Set objShell = CreateObject("Shell.Application")
Set wshShell = CreateObject("WScript.Shell")
Set objFSO = createobject("Scripting.Filesystemobject")

Const strScriptFlagsPath = "C:\ScriptFlags\"
Const strFontSourcePath = "\\eisf.co.ae\SysVol\eisf.co.ae\Policies\{2E463CB6-7E7A-4714-8AB5-6A9231082E5D}\Machine\Scripts\Shutdown\Fonts\"
Const strFontInstallationLogFileName = "FontInstallation.log"
Const strFontInstallationSuccessFlagName = "FontInstallationSuccess.flag"
Const ForAppending = 8
Const FONTS = &H14&
Const strFontGpoResultsPath = "\\eisf.co.ae\SYSVOL\eisf.co.ae\scripts\FontGPO_Results\"
Const strFontGpoResultsLogFileName = "Fonts-GPO-Script-Results.log"

Set objFontsDestinationFolder = objShell.Namespace(FONTS)

objFontInstallationFileLog = strScriptFlagsPath + strFontInstallationLogFileName
objFontInstallationSuccessFlag = strScriptFlagsPath + strFontInstallationSuccessFlagName

If not objFSO.FolderExists(strScriptFlagsPath) Then
objFSO.CreateFolder(strScriptFlagsPath)
End If

If not objFSO.FileExists(objFontInstallationFileLog) Then
objFSO.CreateTextFile(objFontInstallationFileLog)
End If

Set objFontInstallationFileLog2 = objFSO.OpenTextFile(objFontInstallationFileLog, ForAppending, True)
objFontInstallationFileLog2.WriteLine("----------------------------------------------------------------------------------------------------")

objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Starting Log Output at " & strScriptFlagsPath & "...")
objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Installing fonts script version 3.2")
objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Created by ESI IT Department - Muhammad Majid")


If not objFSO.FileExists(objFontInstallationSuccessFlag) Then
    objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Previous font installation success Flag not found on this machine..")
    objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Looking for fonts at " & strFontSourcePath & "...")
    
    If objFSO.FolderExists(strFontSourcePath) Then
        objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Font Source Path found...")
        Set objNameSpace = objShell.Namespace(strFontSourcePath)
        Set objFontsSourceFolder = objFSO.getFolder(strFontSourcePath)
      
        For Each objFile In objFontsSourceFolder.files
          If LCase(right(objFile,4)) = ".ttf" OR LCase(right(objFile,4)) = ".otf" Then
            
            Set objFont = objNameSpace.ParseName(objFile.Name)
            objFontsDestinationFolder.CopyHere objFontsSourceFolder & "\" & objFSO.GetFileName(objFile)
            objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Installed Font: " & objFont)
            Set objFont = Nothing
          End If
        Next
        
        objFSO.CreateTextFile(objFontInstallationSuccessFlag)
        
 '      
        If not objFSO.FolderExists(strFontGpoResultsPath) Then
        objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & strFontGpoResultsPath &" does not exist...")
        End If
        
        If not objFSO.FileExists(strFontGpoResultsPath & strFontGpoResultsLogFileName) Then
        objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & strFontGpoResultsPath & strFontGpoResultsLogFileName & " does not exist...")
        objFSO.CreateTextFile(strFontGpoResultsPath & strFontGpoResultsLogFileName)
        objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & strFontGpoResultsPath & strFontGpoResultsLogFileName & " created...")
        End If
        
        Dim objOutput
        Set objOutput = objFSO.OpenTextFile(strFontGpoResultsPath & strFontGpoResultsLogFileName, ForAppending, True)
        objOutput.WriteLine(Date & "-" & Time & vbTab & "Fonts installation completed successfully on " & wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" ))
        objOutput.Close
        
 '       
        Else
        objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Font Source Path does not exists !!")
    End If
 Else
 objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Previous font installation success Flag detected on this machine..")
 objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Fonts are already installed on this machine ! ")
 End If
 objFontInstallationFileLog2.WriteLine(Date & "-" & Time & vbTab & "Ending Log Output..")
 objFontInstallationFileLog2.WriteLine("----------------------------------------------------------------------------------------------------")
 objFontInstallationFileLog2.WriteLine("")
 objFontInstallationFileLog2.WriteLine("")
 objFontInstallationFileLog2.WriteLine("")
 objFontInstallationFileLog2.Close