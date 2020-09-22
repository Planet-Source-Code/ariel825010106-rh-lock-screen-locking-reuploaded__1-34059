Attribute VB_Name = "RHLockModule01"
'=========================================================================================
'  [RH] Lock Module
'  [RH] Lock Module holds most of [RH] Lock calls
'  if you understand what you're doing, you can modify these codes
'  and make 'em perfect. Be sure to give me your copy of modified [RH] Lock
'=========================================================================================
'  Created By: Ariel
'  Published Date: 20/04/2002
'  E-Mail: ariel825010106@Yahoo.com
'  Legal Copyright: Ariel © 20/04/2002
'=========================================================================================

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const SND_SYNC = &H0
Public Const SND_FILENAME = &H20000
Public Const SRCCOPY = &HCC0020
Public Const EWX_LOGOFF As Long = 0
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_LWIN = &H5B
Public Sub ActivateLock()
    Dim PasswordCount
    On Error GoTo ActivateLock_Error

    With Form2
        PasswordCount = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
        If PasswordCount = "NOT EXISTS" Then GoTo CancelLock
        PasswordCount = DecryptCount(PasswordCount) - 1980
        If CInt(PasswordCount) = 0 Then
CancelLock:
            MsgBox "[RH] Lock cannot lock this computer, because no username or password detected! Please create at least one username or password. To create username and password go to [RH] Lock settings.", vbExclamation, "[RH] Lock"
            Exit Sub
        End If
        'MinimizeAllWindows
        DisableCAD
        HideTaskBar
        .MousePointer = 11
        If Form1.cmbBackground.ListIndex = 0 Then
            Form1.SetBG0
        ElseIf Form1.cmbBackground.ListIndex = 1 Then
            Form1.SetBG1
        ElseIf Form1.cmbBackground.ListIndex = 2 Then
            Form1.SetBG2
        ElseIf Form1.cmbBackground.ListIndex = 3 Then
            Form1.SetBG3
        End If
        .Show
        .LockTheComputer .hwnd
        .MousePointer = 0
    End With

    On Error GoTo 0
    Exit Sub

ActivateLock_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ActivateLock of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Function FileExists(CheckFilename As String) As Boolean
    Const SD = "\"

    Dim TFileLoader As Integer

    On Error Resume Next

    If Right$(CheckPathName, 1) = SD Then
        CheckFilename = Left$(CheckFilename, Len(CheckFilename) - 1)
    End If

    TFileLoader = FreeFile
    Open CheckFilename For Input As TFileLoader
    FileExists = IIf(Err = 0, True, False)
    Close TFileLoader
    Err = 0
End Function

Public Sub CenterInputPassBox()
    Dim TempVarA As Long
    Dim TempVarB As Long
    On Error GoTo CenterInputPassBox_Error

    With Form2
        .picInputPassword.Visible = False
        .picInputPassword.Top = (.Height / 2) - (.picInputPassword.Height / 2)
        .picInputPassword.Left = (.Width / 2) - (.picInputPassword.Width / 2)
        If Form1.cmbBackground.ListIndex <> 2 Then
            .picInputPassword.ScaleMode = 3
            .ScaleMode = 3
            TempVarA = GetDesktopWindow
            TempVarB = GetDC(TempVarA)
            Call BitBlt(.picInputPassword.hdc, 0, 0, .picInputPassword.ScaleWidth, .picInputPassword.ScaleHeight, TempVarB, .picInputPassword.Top, .picInputPassword.Left, SRCCOPY)
            Call ReleaseDC(TempVarA, TempVarB)
            .picInputPassword.DrawMode = 3
            .picInputPassword.ForeColor = RGB(255, 255, 0)
            .picInputPassword.Line (0, 0)-(.picInputPassword.Width, .picInputPassword.Height), , BF
            .picInputPassword.Refresh
            .ScaleMode = 1
            .picInputPassword.ScaleMode = 1
        End If
        .txtPassword.Text = ""
        .txtUsername.Text = ""
        .picInputPassword.Visible = True
    End With

    On Error GoTo 0
    Exit Sub

CenterInputPassBox_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure CenterInputPassBox of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Function DecryptPassword(PasswordString As String) As String
    On Error GoTo DecryptPassword_Error

    PasswordString = Right(PasswordString, Len(PasswordString) - 3)
    For i = 1 To Len(PasswordString)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(PasswordString, i, 1)) - 80 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(PasswordString, i, 1)) - 80 Mod i / 18)
        End If
    Next
    DecryptPassword = TempVar

    On Error GoTo 0
    Exit Function

DecryptPassword_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DecryptPassword of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Sub EnableCAD()
    Dim f
    On Error GoTo EnableCAD_Error

    f = SystemParametersInfo(97, False, CStr(1), 0)

    On Error GoTo 0
    Exit Sub

EnableCAD_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure EnableCAD of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub DisableCAD()
    Dim f
    On Error GoTo DisableCAD_Error

    f = SystemParametersInfo(97, True, CStr(1), 0)

    On Error GoTo 0
    Exit Sub

DisableCAD_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DisableCAD of Module RHLockModule01. Please report this to the author of this program."
End Sub
Public Sub GoToTray()
    On Error GoTo GoToTray_Error

    Screen.MousePointer = 11
    If SettingChanged = True Then
        SaveRHSettings
        SettingChanged = False
    End If
    With Form1
        .Icon1.CreateIcon .picSystray.Picture, "[RH] Lock"
        .Hide
    End With
    Screen.MousePointer = 0

    On Error GoTo 0
    Exit Sub

GoToTray_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure GoToTray of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub InstallDefaultSettings()
    On Error GoTo InstallDefaultSettings_Error

    With Form1
        .chkHotkey.Value = 1
        .chkMaskPassword.Value = 1
        .chkPassSensitive.Value = 1
        .chkShowLockText.Value = 0
        .chkStartup.Value = 1
        .chkUseMaxTry.Value = 0
        .chkUseMultipleUser.Value = 1
        .chkUseSplash = 1
        .cmbAction.ListIndex = 1
        .cmbBackground.ListIndex = 3
        .cmbNumberOfTry.ListIndex = 4
        .txtBGLocation.Text = ""
        SaveRHSettings
    End With

    On Error GoTo 0
    Exit Sub

InstallDefaultSettings_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure InstallDefaultSettings of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub MinimizeAllWindows()
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub
Public Function EncryptPassword(PasswordString As String) As String
    On Error GoTo EncryptPassword_Error

    For i = 1 To Len(PasswordString)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(PasswordString, i, 1)) + 80 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(PasswordString, i, 1)) + 80 Mod i / 18)
        End If
    Next
    TempVar = "GAS" & TempVar
    EncryptPassword = TempVar

    On Error GoTo 0
    Exit Function

EncryptPassword_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure EncryptPassword of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function EncryptCount(CountInteger) As String
    On Error GoTo EncryptCount_Error

    For i = 1 To Len(CountInteger)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(CountInteger, i, 1)) + 32 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(CountInteger, i, 1)) + 32 Mod i / 6)
        End If
    Next
    TempVar = "G1N" & TempVar
    EncryptCount = TempVar

    On Error GoTo 0
    Exit Function

EncryptCount_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure EncryptCount of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function EncryptUsername(UsernameString As String) As String
    On Error GoTo EncryptUsername_Error

    For i = 1 To Len(UsernameString)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(UsernameString, i, 1)) + 28 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(UsernameString, i, 1)) + 28 Mod i / 4)
        End If
    Next
    TempVar = "RH" & TempVar
    EncryptUsername = TempVar

    On Error GoTo 0
    Exit Function

EncryptUsername_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure EncryptUsername of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function DeleteKey(ByVal HKEY As Long, ByVal STRKEY As String)
    Dim TempVar As Long
    On Error GoTo DeleteKey_Error

    TempVar = RegDeleteKey(HKEY, STRKEY)
    If TempVar = 0 Then
        DeleteKey = "Success"
    Else
        DeleteKey = "Not Success"
    End If

    On Error GoTo 0
    Exit Function

DeleteKey_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DeleteKey of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function DeleteValue(ByVal HKEY As Long, ByVal StrPath As String, ByVal StrValue As String)
    Dim KeyHandle As Long
    Dim TempVar As Long
    On Error GoTo DeleteValue_Error

    TempVar = RegOpenKey(HKEY, StrPath, KeyHandle)
    TempVar = RegDeleteValue(KeyHandle, StrValue)
    TempVar = RegCloseKey(KeyHandle)

    If TempVar = 0 Then
        DeleteValue = "Success"
    Else
        DeleteValue = "Not Success"
    End If

    On Error GoTo 0
    Exit Function

DeleteValue_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DeleteValue of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function DecryptUsername(UsernameString As String) As String
    On Error GoTo DecryptUsername_Error

    UsernameString = Right(UsernameString, Len(UsernameString) - 2)
    For i = 1 To Len(UsernameString)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(UsernameString, i, 1)) - 28 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(UsernameString, i, 1)) - 28 Mod i / 4)
        End If
    Next
    DecryptUsername = TempVar

    On Error GoTo 0
    Exit Function

DecryptUsername_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DecryptUsername of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function DecryptCount(CountString) As Integer
    On Error GoTo DecryptCount_Error

    CountString = Right(CountString, Len(CountString) - 3)
    For i = 1 To Len(CountString)
        If i <= 100 Then
            TempVar = TempVar & Chr(Asc(Mid(CountString, i, 1)) - 32 Mod i)
        Else
            TempVar = TempVar & Chr(Asc(Mid(CountString, i, 1)) - 32 Mod i / 6)
        End If
    Next
    DecryptCount = TempVar

    On Error GoTo 0
    Exit Function

DecryptCount_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure DecryptCount of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Sub ExitFromRHLock()
    On Error GoTo ExitFromRHLock_Error

    Form1.Icon1.DeleteIcon
    Unload Form6
    Unload Form5
    Unload Form4
    Unload Form3
    Unload Form2
    Unload Form1
    End

    On Error GoTo 0
    Exit Sub

ExitFromRHLock_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ExitFromRHLock of Module RHLockModule01. Please report this to the author of this program."
End Sub
Public Sub ActionLogOff()
    Dim x
    On Error GoTo ActionLogOff_Error

    x = ExitWindowsEx(EWX_LOGOFF, 0&)

    On Error GoTo 0
    Exit Sub

ActionLogOff_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ActionLogOff of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub ActionRestart()
    Dim x
    On Error GoTo ActionRestart_Error

    x = ExitWindowsEx(EWX_REBOOT, 0&)

    On Error GoTo 0
    Exit Sub

ActionRestart_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ActionRestart of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub ActionShutDown()
    Dim x
    On Error GoTo ActionShutDown_Error

    x = ExitWindowsEx(EWX_SHUTDOWN, 0&)

    On Error GoTo 0
    Exit Sub

ActionShutDown_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ActionShutDown of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub ActionAlarm(AFilename)
    On Error GoTo ActionAlarm_Error

    PlaySound AFilename, ByVal 0&, SND_FILENAME Or SND_ASYNC

    On Error GoTo 0
    Exit Sub

ActionAlarm_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ActionAlarm of Module RHLockModule01. Please report this to the author of this program."
End Sub

Public Sub HideTaskBar()
    Dim rtn
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub
Public Sub ShowTaskbar()
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

Public Function DialogHookFunction(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim ComDlg As cNewDialog
    Set ComDlg = HookedDialog
    If Not (ComDlg Is Nothing) Then
        DialogHookFunction = ComDlg.DialogHook(hDlg, msg, wParam, lParam)
    End If
End Function
Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub
Public Function NullTrim(s) As String
    Dim i As Integer
    i = InStr(s, vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
    s = Trim$(s)
    NullTrim = s
End Function
Function GetDWORD(ByVal HKEY As Long, ByVal StrPath As String, ByVal StrValueName As String) As Long
    Dim TempResult As Long
    Dim TempValueType As Long
    Dim TempBuf As Long
    Dim TempDataBufSize As Long
    Dim TempVar As Long
    Dim KeyHandle As Long
    On Error GoTo GetDWORD_Error

    TempVar = RegOpenKey(HKEY, StrPath, KeyHandle)
    TempDataBufSize = 4
    TempResult = RegQueryValueEx(KeyHandle, StrValueName, 0&, TempValueType, TempBuf, TempDataBufSize)
    If TempResult = ERROR_SUCCESS Then
        If TempValueType = REG_DWORD Then
            GetDWORD = TempBuf
        End If
    End If
    TempVar = RegCloseKey(KeyHand)
    If GetDWORD = "" Then
        GetDWORD = "Not Exist"
    End If

    On Error GoTo 0
    Exit Function

GetDWORD_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure GetDWORD of Module RHLockModule01. Please report this to the author of this program."
End Function
Public Function GetString(HKEY As Long, StrPath As String, StrValue As String)
    Dim KeyHandle As Long
    Dim TempDataType As Long
    Dim TempResult As Long
    Dim TempStrBuf As String
    Dim TempDataBufSize As Long
    Dim TempZeroPos As Integer
    On Error GoTo GetString_Error

    R = RegOpenKey(HKEY, StrPath, KeyHandle)
    TempResult = RegQueryValueEx(KeyHandle, StrValue, 0&, lValueType, ByVal 0&, TempDataBufSize)
    If lValueType = REG_SZ Then
        TempStrBuf = String(TempDataBufSize, " ")
        TempResult = RegQueryValueEx(KeyHandle, StrValue, 0&, 0&, ByVal TempStrBuf, TempDataBufSize)
        If TempResult = ERROR_SUCCESS Then
            TempZeroPos = InStr(TempStrBuf, Chr$(0))
            If TempZeroPos > 0 Then
                GetString = Left$(TempStrBuf, TempZeroPos - 1)
            Else
                GetString = TempStrBuf
            End If
        End If
    End If
    If GetString = "" Then
        GetString = "NOT EXISTS"
    End If

    On Error GoTo 0
    Exit Function

GetString_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure GetString of Module RHLockModule01. Please report this to the author of this program."
End Function
Function SaveDWORD(ByVal HKEY As Long, ByVal StrPath As String, ByVal StrValueName As String, ByVal LongData As Long)
    Dim TempResult As Long
    Dim KeyHandle As Long
    Dim TempVar As Long
    TempVar = RegCreateKey(HKEY, StrPath, KeyHandle)
    TempResult = RegSetValueEx(KeyHandle, StrValueName, 0&, REG_DWORD, LongData, 4)
    TempVar = RegCloseKey(KeyHandle)
    If TempVar = 0 Then
        SaveDWORD = "Success"
    Else
        SaveDWORD = "Not Success"
    End If
End Function
Public Sub SaveKey(HKEY As Long, StrPath As String)
    Dim KeyHand&
    Dim TempVar
    TempVar = RegCreateKey(HKEY, StrPath, KeyHand&)
    TempVar = RegCloseKey(KeyHand&)
End Sub
Public Sub SavePasswords()
    On Error GoTo SavePasswords_Error

    Screen.MousePointer = 11
    Dim PasswordCount As String
    Dim SelectedPos As String
    Dim UsernameLoop As Integer
    SaveKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys"
    PasswordCount = EncryptCount(Form1.lstPasswordList.ListItems.Count + 1980)
    SelectedPos = EncryptCount(Form1.lstPasswordList.SelectedItem.Index + 1984)
    SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID", CStr(PasswordCount)
    SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "", CStr(SelectedPos)
    For UsernameLoop = 1 To Form1.lstPasswordList.ListItems.Count
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & UsernameLoop, EncryptUsername(Form1.lstPasswordList.ListItems.Item(UsernameLoop).Text)
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & UsernameLoop, EncryptPassword(Form1.lstPasswordList.ListItems.Item(UsernameLoop).SubItems(1))
    Next UsernameLoop
    Screen.MousePointer = 0

    On Error GoTo 0
    Exit Sub

SavePasswords_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SavePasswords of Module RHLockModule01. Please report this to the author of this program."
End Sub
Public Sub ReadPasswords()

    On Error GoTo ReadPasswords_Error

    Form1.MousePointer = 11
    With Form1
        Dim UsernameLoop As Integer
        PasswordCount = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
        SelectedPos = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "")

        If UCase(PasswordCount) = "NOT EXISTS" Or UCase(SelectedPos) = "NOT EXISTS" Then GoTo IgnoreRegistryExistence

        SelectedPos = DecryptCount(SelectedPos) - 1984
        PasswordCount = DecryptCount(PasswordCount) - 1980

        For UsernameLoop = 1 To PasswordCount
            .lstPasswordList.ListItems.Add , , DecryptUsername(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & UsernameLoop))
            .lstPasswordList.ListItems.Item(.lstPasswordList.ListItems.Count).SubItems(1) = DecryptPassword(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & UsernameLoop))
        Next UsernameLoop
        .lstPasswordList.ListItems.Item(SelectedPos).Bold = True
        Form1.MousePointer = 0
    End With
    Exit Sub

IgnoreRegistryExistence:
    Form1.lstPasswordList.ListItems.Clear
    Form1.MousePointer = 0
    On Error GoTo 0
    Exit Sub

ReadPasswords_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ReadPasswords of Module RHLockModule01. Please report this to the author of this program."
End Sub
Public Sub SaveRHSettings()
    On Error GoTo SaveRHSettings_Error

    Form1.MousePointer = 11
    SaveKey HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings"
    With Form1
        Dim AppPath As String
        AppPath = App.Path
        If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
        AppPath = """" & AppPath & "[RH] Lock.exe" & """" & " /L"
        If .chkStartup.Value = 1 Then
            SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "[RH] Lock Startup Loader", AppPath
        Else
            SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "[RH] Lock Startup Loader", "NOT LOADED:" & AppPath
        End If
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 01", (.chkStartup.Value + 18)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 02", (.chkUseSplash.Value + 7)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 03", (.chkHotkey.Value + 1980)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 04", (.chkPassSensitive.Value + 18)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 05", (.chkUseMultipleUser.Value + 7)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 06", (.chkMaskPassword.Value + 1980)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 07", (.cmbBackground.ListIndex + 18)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 08", (.chkShowLockText.Value + 7)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 09", .txtBGLocation.Text
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 10", (.chkUseMaxTry.Value + 1980)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 11", (.cmbNumberOfTry.ListIndex + 18)
        SaveString HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 12", (.cmbAction.ListIndex + 7)
    End With
    Form1.MousePointer = 0

    On Error GoTo 0
    Exit Sub

SaveRHSettings_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SaveRHSettings of Module RHLockModule01. Please report this to the author of this program."
End Sub
Public Sub ReadRHSettings()
    Form1.MousePointer = 11
    On Error GoTo GetOut
    With Form1
        .chkStartup.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 01") - 18
        .chkUseSplash.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 02") - 7
        .chkHotkey.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 03") - 1980
        .chkPassSensitive.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 04") - 18
        .chkUseMultipleUser.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 05") - 7
        .chkMaskPassword.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 06") - 1980
        .cmbBackground.ListIndex = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 07") - 18
        .chkShowLockText.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 08") - 7
        .txtBGLocation.Text = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 09")
        .chkUseMaxTry.Value = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 10") - 1980
        .cmbNumberOfTry.ListIndex = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 11") - 18
        .cmbAction.ListIndex = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 12") - 7
    End With
    Form1.MousePointer = 0
    Exit Sub

GetOut:
    Form1.MousePointer = 0
    Exit Sub
End Sub
Public Function SaveString(HKEY As Long, StrPath As String, StrValue As String, StrData As String)
    Dim KeyHandle As Long
    Dim TempVar As Long
    On Error GoTo SaveString_Error

    TempVar = RegCreateKey(HKEY, StrPath, KeyHandle)
    TempVar = RegSetValueEx(KeyHandle, StrValue, 0, REG_SZ, ByVal StrData, Len(StrData))
    TempVar = RegCloseKey(KeyHandle)
    If TempVar = 0 Then
        SaveString = "Success"
    Else
        SaveString = "Not Success"
    End If

    On Error GoTo 0
    Exit Function

SaveString_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SaveString of Module RHLockModule01. Please report this to the author of this program."
End Function
