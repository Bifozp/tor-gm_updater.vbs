Option Explicit

'--- configs
Const enableUacCtl    = True
Const steamRegEntry32 = "HKLM\SOFTWARE\Valve\Steam\InstallPath"
Const steamRegEntry64 = "HKLM\SOFTWARE\Wow6432Node\Valve\Steam\InstallPath"
Const auRegEntry      = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 945360\InstallLocation"
Const githubApiEntry  = "https://api.github.com/repos"
Const githubWeb       = "https://github.com"
Const modDir          = "\BepInEx\plugins"
Const githubModRepo   = "/yukinogatari/TheOtherRoles-GM"  ' TheOtherRolesGM
Const modFileName     = "TheOtherRolesGM.dll"


'--- codes
Dim oShell : Set oShell = WScript.CreateObject("Wscript.Shell")
Dim oXmlHttp : Set oXmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")
Dim oFso : Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
Dim oHtmlfile : Set oHtmlfile = Wscript.CreateObject("htmlfile")
Dim oStream : Set oStream = WScript.CreateObject("ADODB.Stream")

Function downloadFile(url, path, size)
  downloadFile = False
  On Error Resume Next
    With oXmlHttp
      .Open "GET", url, False
      .send
    End With
    If Err.Number = 0 Then
      If oXmlHttp.status <> 200 Then
        WScript.Echo "�_�E�����[�h�Ɏ��s���܂���"
        WScript.Echo "( " & oXmlHttp.statusText & ")"
        Exit Function
      End If
      Err.Clear
      With oStream
        .type = 1
        .open
        .write oXmlHttp.responseBody
        If size <> 0 And .Size <> size Then
          WScript.Echo "�_�E�����[�h�����t�@�C���̃T�C�Y����v���܂���"
          WScript.Echo "( info:" & .Size & " / size:" & size & ")"
          .close
          Exit Function
        End If
        WScript.Echo ">> " & path
        .SaveToFile path, 2
        .close
      End With
      If Err.Number = 0 Then
        downloadFile = True
      Else
        WScript.Echo "�t�@�C���̏������݂Ɏ��s���܂��� (" & Err.Number & " : " & Err.Description & ")"
      End If
    Else
      WScript.Echo "�_�E�����[�h�Ɏ��s���܂��� (" & Err.Number & " : " & Err.Description & ")"
    End If
  Err.Clear
  On Error Goto 0
End Function

Function updateMod(auDir)
  Dim curVersion
  Dim latestVersion
  Dim json
  Dim modPath
  updateMod = -1

  WScript.Echo ""

  '--- get latest release info
  On Error Resume Next
    With oXmlHttp
      .Open "GET", githubApiEntry & githubModRepo & "/releases/latest", False
      .setOption 2, .getOption(2)
      .setRequestHeader "Content-Type", "application/json"
      .send
    End With
    If Err.Number <> 0 Or oXmlHttp.status <> 200 Then
      WScript.Echo "Mod�̍ŐV�o�[�W�����̊m�F�Ɏ��s���܂���"
      Exit Function
    End If
    Set json = oHtmlfile.JsonParse(oXmlHttp.responseText)
    latestVersion = Mid(json.tag_name, 2)
    If IsEmpty(latestVersion) Then
      WScript.Echo "Mod�̍ŐV�o�[�W�����̏�񂪂���܂���"
      Exit Function
    End If
  On Error Goto 0
  WScript.Echo "�ŐV�o�[�W����: " & latestVersion
  
  '--- get current mod version
  modPath = auDir & modDir & "\" & modFileName
  On Error Resume Next
    curVersion = oFso.GetFileVersion(modPath)
  On Error Goto 0
  If IsEmpty(curVersion) Then
    WScript.Echo "Mod�̌��s�o�[�W�����̏�񂪂���܂���"
  Else
    WScript.Echo "���s�o�[�W����: " & curVersion
  End If

  If StrComp(latestVersion, Left(curVersion, Len(latestVersion))) = 0 Then
    WScript.Echo "Mod�͊��ɍŐV�łł��A�X�V�̕K�v�͂���܂���"
    updateMod = 0
    Exit Function
  End If
  
  WScript.Echo ""
  WScript.Echo "�A�b�v�f�[�g���J�n���܂�"
  WScript.Echo "----------"
  WScript.Echo json.tag_name
  WScript.Echo "[�X�V���e]"
  WScript.Echo json.body
  WScript.Echo "----------"
  '--- Download mod
  Dim i
  For i=0 To json.assets.Length-1
    Dim assetUrl  : assetUrl  = Eval("json.assets.[" & i & "].browser_download_url")
    Dim assetSize : assetSize = Eval("json.assets.[" & i & "].size")
    If StrComp(Right(assetUrl, 4), ".dll") = 0 Then
      WScript.Echo "Mod���_�E�����[�h���Ă��܂�"
      If downloadFile(assetUrl, modPath, assetSize) Then
        WScript.Echo "Mod�̍X�V���������܂���"
        updateMod = 0
      Else
        WScript.Echo "Mod�̍X�V�Ɏ��s���܂���"
      End If
      Exit For
    End If
  Next
End Function

Function extractZip(zipPath, auDir)
  extractZip = False
  If Not oFso.FolderExists(auDir) Then
    WScript.Echo "�𓀐�̃t�H���_�����݂��܂���"
    Exit Function
  End If
  With CreateObject("Shell.Application")
    .NameSpace(auDir).CopyHere .NameSpace(zipPath).Items
  End With
  extractZip = True
End Function

Function installNew(auDir)
  Dim latestVersion
  Dim jsons
  Dim modPath
  installNew = -1

  WScript.Echo ""

  '--- get latest release info
  On Error Resume Next
    With oXmlHttp
      .Open "GET", githubApiEntry & githubModRepo & "/releases?per_page=5", False
      .setOption 2, .getOption(2)
      .setRequestHeader "Content-Type", "application/json"
      .send
    End With
    If Err.Number <> 0 Or oXmlHttp.status <> 200 Then
      WScript.Echo "Mod�̍ŐV�o�[�W�����̊m�F�Ɏ��s���܂���"
      Exit Function
    End If
    Set jsons = oHtmlfile.JsonParse(oXmlHttp.responseText)
    latestVersion = Mid(jsons.[0].tag_name, 2)
  On Error Goto 0
  WScript.Echo "�ŐV�o�[�W����: " & latestVersion
  modPath = auDir & modDir & "\" & modFileName
  WScript.Echo ""
  WScript.Echo "Mod�̃C���X�g�[�����J�n���܂�"

  '--- Download mod
  Dim i : Dim j
  Dim assetVersion
  Dim triedInstall : triedInstall = False
  j = 0
  Do While j < jsons.Length-1
    Dim json : Set json = Eval("jsons.[" & j & "]")
    For i=0 To json.assets.Length-1
      Dim assetUrl  : assetUrl  = Eval("json.assets.[" & i & "].browser_download_url")
      Dim assetSize : assetSize = Eval("json.assets.[" & i & "].size")
      If StrComp(Right(assetUrl, 4), ".zip") = 0 Then
        Dim zipPath : zipPath = oFso.getParentFolderName(WScript.ScriptFullName) & "\" & oFso.GetFileName(assetUrl)
        triedInstall = True
        On Error Resume Next
          WScript.Echo "----------"
          WScript.Echo json.tag_name
          WScript.Echo "[�X�V���e]"
          WScript.Echo json.body
          WScript.Echo "----------"
          WScript.Echo "Mod���_�E�����[�h���Ă��܂�"
          If downloadFile(assetUrl, zipPath, assetSize) Then
            WScript.Echo "zip�t�@�C�����𓀂��Ă��܂�"
            If extractZip(zipPath, auDir) Then
              If j = 0 Then
                WScript.Echo "Mod�̃C���X�g�[�����������܂���"
                installNew = 0
              Else
                WScript.Echo "�Â��o�[�W�������珉���C���X�g�[�������{�������߁AMod�̃A�b�v�f�[�g���s���܂�"
                installNew = updateMod(auDir)
              End If
            Else
              WScript.Echo "Mod�̃C���X�g�[���Ɏ��s���܂���"
            End If
          End If
        On Error Goto 0
        Exit Do
      End If
    Next
    j = j + 1
  Loop

  If Not triedInstall Then
    WScript.Echo "�ŋ߂̃o�[�W��������A�Ή�����zip�A�Z�b�g�������邱�Ƃ��ł��܂���ł���"
    WScript.Echo "�ȉ��̃����[�X�ꗗ�́A���߂�Release��Assets����zip�t�@�C����T���āA�蓮�ŃC���X�g�[�����Ă��������B"
    WScript.Echo githubWeb & githubModRepo & "/releases"
  End If
End Function

Function isModEnvInstalled(auDir)
  isModEnvInstalled = False
  If Not oFso.FolderExists(auDir & "\BepInEx") Then
    Exit Function
  ElseIf Not oFso.FolderExists(auDir & "\mono") Then
    Exit Function
  ElseIf Not oFso.FileExists(auDir & "\winhttp.dll") Then
    Exit Function
  ElseIf Not oFso.FileExists(auDir & "\doorstop_config.ini") Then
    Exit Function
  End If
  isModEnvInstalled = True
End Function

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
        WScript.Echo "Steam ���C���X�g�[������Ă��܂���"
        isSteamInstalled = False
      End If
    End If
  Err.Clear
  On Error Goto 0
End Function

Function main()
  Dim auDir
  
  WScript.Echo "Among Us Mod Updater for TheOtherRoles-GM (Steam)"
  WScript.Echo ""

  If Not isSteamInstalled Then
    main = 1 : Exit Function
  End If

  auDir = getAuDir
  If auDir = "" Then
    main = 1 : Exit Function
  End If
  WScript.Echo "Among Us Dir: """ & auDir & """"

  ' --- Init htmlfile for parse json (http://bougyuusonnin.seesaa.net/article/446183415.html)
  oHtmlfile.write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"
  oHtmlfile.write "<script>document.JsonParse=function (s) {return eval('(' + s + ')');}</script>"
  oHtmlfile.write "<script>document.JsonStringify=JSON.stringify;</script>"

  ' --- Install New
  If Not isModEnvInstalled(auDir) Then
    WScript.Echo "Mod�����C���X�g�[������Ă��Ȃ����A�s���S�ȈׁA�����C���X�g�[�������{���܂�"
    main = installNew(auDir)
    Exit Function
  End If
  
  ' --- Update
  WScript.Echo "Mod�������o���܂����AMod�݂̂̃A�b�v�f�[�g�����{���܂�"
  main = updateMod(auDir)
  Exit Function
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
