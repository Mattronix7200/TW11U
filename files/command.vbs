Option Explicit

Dim strPath, oShell

strPath = SelectFolder( "" )
If strPath = vbNull Then
    WScript.Echo "Nie wybrano katalogu. By kontynuowa" & ChrW(&H107) & ", wska" & ChrW(&H17C) & " folder z rozpakowanymi plikami instalacyjnymi Windows 11..."
Else
    Set oShell = WScript.CreateObject("WScript.Shell")
    oShell.Run "cmd.exe /c echo " & strPath & "| clip", 0, True
    Dim x : x = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
    x = Replace(Replace(x, Chr(10), ""), Chr(13), "")
    Dim datavar,datavar2,datavar3,cmd,cmd2,scriptdir
    datavar = x
    datavar = datavar & "\sources\"
    datavar2= datavar
    datavar3= x
    scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    cmd = "cmd.exe /k cd " & datavar2 & " && del /f /q appraiserres.dll"
    cmd2 = "cmd.exe /k cd " & datavar3 & " && start setup.exe"
    oShell.Exec(cmd)
    oShell.Run "cmd.exe /c """ & scriptdir & "\fixer.bat""", 1, True
    oShell.Exec(cmd2)
    WScript.Echo "Gotowe :)"
End If

Function SelectFolder( myStartFolder )
    Dim objFolder, objItem, objShell
    On Error Resume Next
    SelectFolder = vbNull

    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Wybierz folder z rozpakowanymi plikami instalacyjnymi systemu Windows 11", 0, myStartFolder )

    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function

