VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInputPassword 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2085
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   5100
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   5100
      Begin prjChameleon.chameleonButton cmdOK 
         Default         =   -1  'True
         Height          =   420
         Left            =   3690
         TabIndex        =   5
         Top             =   1170
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         BTYPE           =   3
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         MPTR            =   0
         MICON           =   "Form2.frx":000C
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1035
         PasswordChar    =   "#"
         TabIndex        =   4
         Top             =   1620
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1035
         TabIndex        =   2
         Top             =   1170
         Width           =   2535
      End
      Begin prjChameleon.chameleonButton chameleonButton1 
         Height          =   285
         Left            =   3690
         TabIndex        =   6
         Top             =   1650
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         MPTR            =   0
         MICON           =   "Form2.frx":0028
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   135
         Picture         =   "Form2.frx":0044
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[RH] Lock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Index           =   3
         Left            =   3105
         TabIndex        =   8
         Top             =   45
         Width           =   1905
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   855
         X2              =   5085
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This computer is locked by [RH] Lock. To restore this computer, please insert your password."
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   2
         Left            =   945
         TabIndex        =   7
         Top             =   585
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password :"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   1665
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username :"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   1215
         Width           =   825
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblLockInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This computer is locked by [RH] Lock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   585
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Image imgLockIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   0
      Picture         =   "Form2.frx":1D0E
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgBackground 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private K() As cControlFlater, i As Integer

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Const MF_BYPOSITION = &H400
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200
Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Dim UserTry As Integer
Public Sub LockTheComputer(TheWindowsHandle As Long)
    Dim QWickIPyck

   On Error GoTo LockTheComputer_Error

    QWickIPyck = SetWindowPos(TheWindowsHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    SystemMenu% = GetSystemMenu(hwnd, 0)
    Res% = RemoveMenu(SystemMenu%, 6, MF_BYPOSITION)

   On Error GoTo 0
   Exit Sub

LockTheComputer_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure LockTheComputer of Form Form2. Please report this to the author of this program."
End Sub

Private Sub chameleonButton1_Click()
   On Error GoTo chameleonButton1_Click_Error

    Form2.picInputPassword.Visible = False
    Form2.picInputPassword.Enabled = False

   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure chameleonButton1_Click of Form Form2. Please report this to the author of this program."
End Sub

Private Sub chameleonButton1_KeyDown(KeyCode As Integer, Shift As Integer)
    LockTheComputer Form2.hwnd
End Sub


Private Sub cmdOK_Click()
    Dim PasswordCount
    Dim UsernameLoop As Integer
    Dim TempUsername
    Dim TempPassword

    Dim UserPassOK As Boolean
    Dim InputUsername As String
    Dim InputPassword As String
    Dim Filepath As String
   On Error GoTo cmdOK_Click_Error

    UserPassOK = False

    If Len(Form2.txtPassword.Text) < 3 And Len(Form2.txtUsername.Text) < 3 Then
        GoTo OKEY
        Exit Sub
    End If

    InputUsername = Form2.txtUsername.Text
    InputPassword = Form2.txtPassword.Text


    If Len(Form2.txtPassword.Text) >= 3 And Len(Form2.txtUsername.Text) >= 3 Then
        Form2.MousePointer = 11
        PasswordCount = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
        SelectedPos = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "")
        If UCase(PasswordCount) = "NOT EXISTS" Or UCase(SelectedPos) = "NOT EXISTS" Then Exit Sub

        PasswordCount = DecryptCount(PasswordCount) - 1980
        SelectedPos = DecryptCount(SelectedPos) - 1984

        If Form1.chkUseMultipleUser.Value = 1 Then
            For UsernameLoop = 1 To PasswordCount
                TempUsername = DecryptUsername(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & UsernameLoop))
                TempPassword = DecryptPassword(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & UsernameLoop))

                If Form1.chkPassSensitive.Value = 0 Then
                    InputUsername = UCase(InputUsername)
                    InputPassword = UCase(InputPassword)

                    TempUsername = UCase(TempUsername)
                    TempPassword = UCase(TempPassword)
                End If

                If InputUsername = TempUsername And InputPassword = TempPassword Then
                    UserPassOK = True
                    GoTo OKEY
                Else
                    UserPassOK = False
                End If
            Next UsernameLoop
        Else
            TempUsername = DecryptUsername(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & SelectedPos))
            TempPassword = DecryptPassword(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & SelectedPos))
            If InputUsername = TempUsername And InputPassword = TempPassword Then
                UserPassOK = True
                GoTo OKEY
            Else
                UserPassOK = False
            End If
        End If

OKEY:
        If UserPassOK = True Then
            Form2.MousePointer = 0
            Unload Form2
            ShowTaskbar
            EnableCAD
            Exit Sub
        Else
            Form2.MousePointer = 0
            If Form1.chkUseMultipleUser.Value = 1 Then
                UserTry = UserTry + 1
                MsgBox "Wrong default username or password! Please enter your default username and password", vbExclamation, "[RH] Lock"
            Else
                UserTry = UserTry + 1
                MsgBox "Wrong username or password! Please enter correct username and password", vbExclamation, "[RH] Lock"
            End If
            If Form1.chkUseMaxTry.Value = 1 Then
                If UserTry >= Form1.cmbNumberOfTry.Text Then
                    Select Case Form1.cmbAction.ListIndex
                        Case 0
                            Form2.MousePointer = 0
                            ActionShutDown
                            EnableCAD
                            chameleonButton1_Click
                            Exit Sub
                        Case 1
                            Form2.MousePointer = 0
                            ActionRestart
                            EnableCAD
                            chameleonButton1_Click
                            Exit Sub
                        Case 2
                            Form2.MousePointer = 0
                            ActionLogOff
                            EnableCAD
                            chameleonButton1_Click
                            Exit Sub
                        Case 3
                            Filepath = App.Path
                            If Right(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
                            Filepath = Filepath & "[RH] Alarm.wav"
                            Form2.MousePointer = 0
                            If FileExists(Filepath) = True Then ActionAlarm Filepath
                            chameleonButton1_Click
                            Exit Sub
                    End Select
                End If
            End If
        End If
        Form2.MousePointer = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdOK_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure cmdOK_Click of Form Form2. Please report this to the author of this program."
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    LockTheComputer Form2.hwnd
End Sub


Private Sub Form_Click()
   On Error GoTo Form_Click_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With

   On Error GoTo 0
   Exit Sub

Form_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure Form_Click of Form Form2. Please report this to the author of this program."
End Sub

Private Sub Form_DblClick()
    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LockTheComputer Form2.hwnd
    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With
End Sub

Private Sub Form_Load()
    Dim CTL As Control

    For Each CTL In Me.Controls
        Select Case TypeName(CTL)
            Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar"
                ReDim Preserve K(i)
                Set K(i) = New cControlFlater
                K(i).Attach CTL
                i = i + 1
        End Select
    Next CTL

    Me.imgLockIcon.Visible = False
    Me.lblLockInfo.Visible = False
    If Form1.chkShowLockText.Value = 1 Then
        Me.imgLockIcon.Visible = True
        Me.lblLockInfo.Visible = True
    End If
    Me.imgBackground.Picture = LoadPicture("")
    Me.imgBackground.Stretch = False
    Me.imgBackground.Top = 0
    Me.imgBackground.Left = 0

    If Form1.chkMaskPassword.Value = 1 Then
        Form2.txtPassword.PasswordChar = "#"
    ElseIf Form1.chkMaskPassword.Value = 0 Then
        Form2.txtPassword.PasswordChar = ""
    End If
End Sub


Private Sub imgBackground_Click()
   On Error GoTo imgBackground_Click_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With

   On Error GoTo 0
   Exit Sub

imgBackground_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure imgBackground_Click of Form Form2. Please report this to the author of this program."
End Sub


Private Sub imgBackground_DblClick()
   On Error GoTo imgBackground_DblClick_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With

   On Error GoTo 0
   Exit Sub

imgBackground_DblClick_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure imgBackground_DblClick of Form Form2. Please report this to the author of this program."
End Sub


Private Sub imgBackground_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo imgBackground_MouseDown_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With

   On Error GoTo 0
   Exit Sub

imgBackground_MouseDown_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure imgBackground_MouseDown of Form Form2. Please report this to the author of this program."
End Sub


Private Sub imgLockIcon_Click()
   On Error GoTo imgLockIcon_Click_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With
   On Error GoTo 0
   Exit Sub

imgLockIcon_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure imgLockIcon_Click of Form Form2. Please report this to the author of this program."
End Sub

Private Sub lblLockInfo_Click()
   On Error GoTo lblLockInfo_Click_Error

    If Form2.picInputPassword.Visible = True Then Exit Sub
    With Form2
        CenterInputPassBox
        .picInputPassword.Visible = True
        .picInputPassword.Enabled = True
    End With

   On Error GoTo 0
   Exit Sub

lblLockInfo_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure lblLockInfo_Click of Form Form2. Please report this to the author of this program."
End Sub

Private Sub Timer1_Timer()
    LockTheComputer Form2.hwnd
End Sub


Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub


Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    LockTheComputer Form2.hwnd
End Sub


Private Sub txtUsername_GotFocus()
    txtUsername.SelStart = 0
    txtUsername.SelLength = Len(txtUsername.Text)
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
    LockTheComputer Form2.hwnd
End Sub


