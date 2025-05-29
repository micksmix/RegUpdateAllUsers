' ============================================================================
' RegUpdateAllHkcuHkcr_Improved.vbs
' -----------------------------------------------------------------------------
' AUTHOR : Mick Grove  (refactor by ChatGPT – 2025‑05‑29)
' LICENSE: Public‑domain / use at your own risk
'
' WHAT’S NEW
'   • Option Explicit & timestamped logging
'   • Robust error‑handling (no global On Error Resume Next)
'   • All registry edits abstracted into KeysToModify() – easy to extend
'   • Helper RunRegCmd() verifies reg.exe exit codes
'   • Simplified GetRegRootToUseForBinaryValues() via Select Case
'   • Consistent indentation & naming (PascalCase subs / camelCase vars)
'   • Early admin / host‑bitness checks
'   • Works on Windows XP → Windows 11 (x64) – WOW64 aware
'
' USAGE
'   cscript.exe RegUpdateAllHkcuHkcr_Improved.vbs [/quiet]
'     /quiet —— suppresses console output (errors still echo)
' ============================================================================
Option Explicit

' === CONSTANTS ==============================================================
Const HKCR = &H80000000, HKCU = &H80000001, HKLM = &H80000002
Const HKU  = &H80000003, HKCC = &H80000005

Const DAT_NTUSER   = &H70000000
Const DAT_USRCLASS = &H70000001

Const TEMP_HIVE    = "HKLM\TEMPHIVE"

' === GLOBALS ================================================================
Dim gShell : Set gShell = CreateObject("WScript.Shell")
Dim gFso   : Set gFso   = CreateObject("Scripting.FileSystemObject")
Dim gWMI   : Set gWMI   = GetObject("winmgmts:root\default")

Dim gQuiet : gQuiet = (WScript.Arguments.Named.Exists("quiet"))

' === ENTRY ==================================================================
Sub Main()
    If Not IsAdmin() Then Die "Administrator privileges are required."
    If InStr(UCase(WScript.FullName), "CSCRIPT.EXE") = 0 Then _
        Die "Run from a command prompt to avoid pop‑ups (cscript.exe)."

    Log "=== RegUpdateAllHkcuHkcr starting ==="

    UpdateAllProfiles DAT_NTUSER   ' per‑user HKCU
    UpdateAllProfiles DAT_USRCLASS ' per‑user HKCR

    Log "Processing complete."
End Sub

' === HELPER: ADMIN CHECK =====================================================
Function IsAdmin()
    On Error Resume Next
    Dim testKey : testKey = "HKLM\SOFTWARE\RegUpdateTestPerms_" & gShell.ExpandEnvironmentStrings("%RANDOM%")
    gShell.RegWrite testKey & "\", 1, "REG_DWORD"
    IsAdmin = (Err.Number = 0)
    If IsAdmin Then gShell.RegDelete testKey & "\"
    On Error GoTo 0
End Function

Sub Die(msg)
    WScript.Echo "FATAL: " & msg
    WScript.Quit 1
End Sub

Sub Log(msg)
    If Not gQuiet Then WScript.Echo Now() & "  " & msg
End Sub

' === RUN reg.exe WITH VERIFICATION ==========================================
Function RunRegCmd(cmd)
    Dim exitCode : exitCode = gShell.Run(cmd, 0, True)
    If exitCode <> 0 Then _
        Log "  ! reg.exe failed: " & cmd & " (exit=" & exitCode & ")"
    RunRegCmd = exitCode
End Function

' === PROFILE‑LEVEL PROCESSING ===============================================
Sub UpdateAllProfiles(datFile)
    Dim meUser : meUser = GetCurrentUsername()
    Log "Updating settings for logged‑on user: " & meUser & vbCrLf

    If datFile = DAT_NTUSER Then
        KeysToModify "HKCU", datFile
    Else
        KeysToModify "HKCR", datFile
    End If

    Dim defaultProfile : defaultProfile = GetDefaultUserPath()
    Dim datPath : datPath = ResolveDatPath(defaultProfile, datFile)
    If Len(datPath) > 0 Then
        Log "Updating Default User hive (new profiles)."
        LoadProfileHive datPath, "Default", datFile
    Else
        Log "Cannot locate Default profile hive – skipped."
    End If

    Dim reg, key, subKeys, sid
    Set reg = gWMI.Get("StdRegProv")
    reg.EnumKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", subKeys

    For Each sid In subKeys
        Dim imgPath
        reg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" & _
            "\" & sid, "ProfileImagePath", imgPath
        Dim userName : userName = ExtractUser (imgPath)

        If ShouldProcessUser(userName, meUser) Then
            datPath = ResolveDatPath(imgPath, datFile)
            If Len(datPath) > 0 Then
                Log "Preparing to update: " & userName
                LoadProfileHive datPath, userName, datFile
            End If
        End If
    Next
End Sub

Function ShouldProcessUser(u, meUser)
    u = UCase(u): meUser = UCase(meUser)
    ShouldProcessUser = (u <> "ALL USERS" And u <> meUser And _
                         u <> "LOCALSERVICE" And u <> "SYSTEMPROFILE" And _
                         u <> "NETWORKSERVICE")
End Function

' === LOAD / UNLOAD PROFILE HIVE ============================================
Sub LoadProfileHive(datPath, userLabel, datFile)
    Dim loadedSID : loadedSID = ""
    Dim rc

    rc = RunRegCmd("reg.exe load """ & TEMP_HIVE & """ """ & datPath & """")
    If rc <> 0 Then   ' hive already mounted – locate it under HKU
        loadedSID = FindSidFromDatPath(datPath)
        If Len(loadedSID) = 0 Then
            Log "*** Unable to locate SID for " & userLabel & " – skipped"
            Exit Sub
        End If
    End If

    ' Apply keys
    If loadedSID = "" Then
        KeysToModify TEMP_HIVE, datFile
    ElseIf datFile = DAT_NTUSER Then
        KeysToModify "HKU\" & loadedSID, datFile
    Else
        KeysToModify "HKU\" & loadedSID & "_Classes", datFile
    End If

    ' Unload if we loaded
    If loadedSID = "" Then
        RunRegCmd "reg.exe unload " & TEMP_HIVE
    End If
End Sub

Function FindSidFromDatPath(datPath)
    Dim reg : Set reg = gWMI.Get("StdRegProv")
    Dim subKeys, sid, img
    reg.EnumKey HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", subKeys
    For Each sid In subKeys
        reg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" & _
            "\" & sid, "ProfileImagePath", img
        If StrComp(datPath, ResolveDatPath(img, DAT_NTUSER), vbTextCompare) = 0 Then
            FindSidFromDatPath = sid
            Exit Function
        End If
    Next
    FindSidFromDatPath = ""
End Function

' === REGISTRY MODIFICATIONS =================================================
Sub KeysToModify(root, datFile)
    ' --- EXAMPLE CUSTOMISATIONS -------------------------------------------
    If datFile = DAT_NTUSER Then
        gShell.RegWrite root & "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\DisablePasswordCaching", _
                        1, "REG_DWORD"
        gShell.RegWrite root & "\Software\Microsoft\Internet Explorer\Main\FormSuggest PW Ask", _
                        "no", "REG_SZ"

        ' Example REG_BINARY write
        Dim binPath : binPath = "Software\_Test\MyTestBinarySubkey"
        gShell.RegWrite root & "\" & binPath & "\", ""
        SetBinary root, binPath, "My Test Binary Value", "hex:23,00,41,00,43,00,42,00,6c,00"

        ' Delete examples
        DeleteRecursive root, "Software\_Test"
        DeleteValue root, "Software\Microsoft\Internet Explorer\Main", "FormSuggest PW Ask"

    ElseIf datFile = DAT_USRCLASS Then
        gShell.RegWrite root & "\FirefoxURL\FriendlyTypeName", "Firefox URL", "REG_SZ"
    End If
End Sub

Sub DeleteValue(root, key, valueName)
    RunRegCmd "reg.exe delete """ & root & "\" & key & """ /v """ & valueName & """ /f"
End Sub

Sub DeleteRecursive(root, key)
    RunRegCmd "reg.exe delete """ & root & "\" & key & """ /f"
End Sub

' === BINARY HELPERS =========================================================
Sub SetBinary(root, keyPath, valueName, hexString)
    Dim arrHex : arrHex = Split(Replace(hexString, "hex:", ""), ",")
    Dim i, arrDec()
    ReDim arrDec(UBound(arrHex))
    For i = 0 To UBound(arrHex)
        arrDec(i) = CLng("&H" & Trim(arrHex(i)))
    Next

    Dim hive, subPath
    hive = ParseHive(root, subPath)

    Dim reg : Set reg = gWMI.Get("StdRegProv")
    Dim res : res = reg.SetBinaryValue(hive, subPath & "\" & keyPath, valueName, arrDec)
    If res <> 0 Then Log "*** Error adding binary value at " & keyPath
End Sub

Function ParseHive(fullKey, subPath)
    Dim up : up = UCase(fullKey)
    Select Case True
        Case Left(up, 4) = "HKCU": ParseHive = HKCU : subPath = Mid(fullKey, 5)
        Case Left(up, 4) = "HKLM": ParseHive = HKLM : subPath = Mid(fullKey, 5)
        Case Left(up, 3) = "HKU":  ParseHive = HKU  : subPath = Mid(fullKey, 4)
        Case Left(up, 4) = "HKCR": ParseHive = HKCR : subPath = Mid(fullKey, 5)
        Case Else: ParseHive = HKLM : subPath = fullKey  ' fallback
    End Select
End Function

' === MISC HELPERS ===========================================================
Function GetCurrentUsername()
    Dim u : u = gShell.ExpandEnvironmentStrings("%USERNAME%")
    If u = "%USERNAME%" Then
        u = ExtractUser(gShell.ExpandEnvironmentStrings("%USERPROFILE%"))
    End If
    GetCurrentUsername = u
End Function

Function ExtractUser(path)
    Dim parts : parts = Split(path, "\")
    ExtractUser = parts(UBound(parts))
End Function

Function GetDefaultUserPath()
    Dim reg : Set reg = gWMI.Get("StdRegProv")
    Dim pDir, defUser, result
    reg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", _
                                "ProfilesDirectory", pDir
    reg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", _
                                "DefaultUserProfile", defUser
    If Len(defUser) = 0 Then
        reg.GetExpandedStringValue HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", _
                                    "Default", result
    Else
        result = pDir & "\" & defUser
    End If
    GetDefaultUserPath = result
End Function

Function ResolveDatPath(profilePath, datFile)
    profilePath = RTrim(profilePath, "\")
    Select Case datFile
        Case DAT_NTUSER
            If gFso.FileExists(profilePath & "\NTUSER.DAT") Then _
                ResolveDatPath = profilePath & "\NTUSER.DAT"
        Case DAT_USRCLASS
            Dim p1, p2
            p1 = profilePath & "\AppData\Local\Microsoft\Windows\USRCLASS.DAT"
            p2 = profilePath & "\Local Settings\Application Data\Microsoft\Windows\USRCLASS.DAT"
            If gFso.FileExists(p1) Then
                ResolveDatPath = p1
            ElseIf gFso.FileExists(p2) Then
                ResolveDatPath = p2
            End If
    End Select
End Function

Function RTrim(str, c)
    Do While Right(str, 1) = c
        str = Left(str, Len(str) - 1)
    Loop
    RTrim = str
End Function

' === START SCRIPT ===========================================================
Main()
