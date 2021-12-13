Option Explicit

'--- configs
Const enableUacCtl    = True
Const steamRegEntry32 = "HKLM\SOFTWARE\Valve\Steam\InstallPath"
Const steamRegEntry64 = "HKLM\SOFTWARE\Wow6432Node\Valve\Steam\InstallPath"
Const auRegEntry      = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 945360\InstallLocation"


'--- codes
Dim oShell : Set oShell = WScript.CreateObject("Wscript.Shell")
Dim oFso : Set oFso = WScript.CreateObject("Scripting.FileSystemObject")

Sub removeDirectory(path)
  Call oFso.DeleteFolder(path, True)
End Sub
Sub removeFile(path)
  Call oFso.DeleteFile(path, True)
End Sub

Sub remove(path)
  Dim removeSub : removeSub = Null
  If oFso.FolderExists(path) Then
    Set removeSub = GetRef("removeDirectory")
  ElseIf oFso.FileExists(path) Then
    Set removeSub = GetRef("removeFile")
  End If

  If Not IsNull(removeSub) Then
    WScript.Echo """" & path & """ を削除しています"
    Call removeSub(path)
  End If
End Sub

Sub cleanModEnvironments(auDir)
  Call remove(auDir & "\BepInEx")
  Call remove(auDir & "\mono")
  Call remove(auDir & "\winhttp.dll")
  Call remove(auDir & "\doorstop_config.ini")
  Call remove(auDir & "\steam_appid.txt")
End Sub

Function getAuDir()
  Dim auDir
  getAuDir = ""
  On Error Resume Next
    auDir = oShell.RegRead(auRegEntry)
    If Err.Number = 0 Then
      getAuDir = auDir
    End If
  Err.Clear
  On Error Goto 0
End Function

Function isSteamInstalled()
  isSteamInstalled = True
  On Error Resume Next
    oShell.RegRead steamRegEntry64
    If Err.Number <> 0 Then
      Err.Clear : oShell.RegRead steamRegEntry32
      If Err.Number <> 0 Then
        WScript.Echo "Steam がインストールされていません"
        isSteamInstalled = False
      End If
    End If
  Err.Clear
  On Error Goto 0
End Function

Function main()
  Dim auDir
  Dim retMsg
  
  WScript.Echo "Among Us Mod Cleaner (Steam)"
  WScript.Echo ""

  If Not isSteamInstalled Then
    main = 1 : Exit Function
  End If

  auDir = getAuDir
  If auDir = "" Then
    main = 1 : Exit Function
  End If
  WScript.Echo "Among Us Dir: """ & auDir & """"

  retMsg = MsgBox("Among Us のMod環境をおそうじしてもよいですか？", vbYesNo Or vbQuestion Or vbDefaultButton2)
  If retMsg = vbNo Then
    WScript.Echo "Mod環境のおそうじはキャンセルされました"
    main = 0 : Exit Function
  End If

  WScript.Echo ""
  Call cleanModEnvironments(auDir)
  WScript.Echo "おそうじ完了"
  main = 0
End Function

'--- entry point

' UAC (https://www.server-world.info/query?os=Other&p=vbs&f=1)
If enableUacCtl Then
  Do While WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    Dim Wmi : Set Wmi = GetObject("winmgmts:\\.\root\CIMV2")
    Dim OS, Value
    '##### Check if it is WScript 5.7 or Vista or later
    Set OS = Wmi.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value in OS
    If left(Value.Version, 3) < 6.0 Then Exit Do
    Next

    '##### Run as administrator.
    WScript.Quit WScript.CreateObject("Shell.Application").ShellExecute("cmd.exe", " /k cscript.exe /nologo """ & WScript.ScriptFullName & """ uac", "", "runas")
  Loop
End If

If LCase(Right(WScript.FullName, 11)) = "wscript.exe" Then
  Dim args : args = Array("cmd.exe /k cscript.exe /nologo",""""&WScript.ScriptFullName&"""")
  Dim arg
  For Each arg In WScript.Arguments
     ReDim Preserve args(UBound(args)+1)
     args(UBound(args)) = """" & arg & """"
  Next
  WScript.Quit CreateObject("WScript.Shell").Run(Join(args), 1, True)
End If
WScript.Quit main
