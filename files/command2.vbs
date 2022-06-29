Option Explicit

Dim strPath, oShell

strPath = SelectFolder( "" )
If strPath = vbNull Then
    WScript.Echo "Nie wybrano katalogu. By kontynuowa" & ChrW(&H107) & ", wska" & ChrW(&H17C) & " dowolny folder na wymiennym no"&ChrW(&H15B)&"niku, by skopiowa"&ChrW(&H107)&" pliki..."
Else
    Set oShell = WScript.CreateObject("WScript.Shell")
    oShell.Run "cmd.exe /c echo " & strPath & "| clip", 0, True
    Dim x : x = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
    x = Replace(Replace(x, Chr(10), ""), Chr(13), "")
    Dim datavar,cmd,scriptdir
    datavar = x
    scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    cmd = "cmd.exe /c cd """ & datavar & """ && copy """ & scriptdir & "\disable_tpm.reg"" " & """" & datavar &  "\disable_tpm.reg"""
    oShell.Exec(cmd)
    WScript.Echo "Gotowe :)"
End If

Function SelectFolder( myStartFolder )
    Dim objFolder, objItem, objShell
    On Error Resume Next
    SelectFolder = vbNull

    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Wybierz folder do kt"&ChrW(&HF3)&"rego chcesz skopiowa"&ChrW(&H107)&" pliki", 0, myStartFolder )

    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function

