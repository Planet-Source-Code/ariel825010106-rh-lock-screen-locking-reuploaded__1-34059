VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[RH] Lock - Input Password"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   " Input default Username and Password "
      Height          =   2355
      Left            =   82
      TabIndex        =   0
      Top             =   45
      Width           =   4980
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1935
         PasswordChar    =   "#"
         TabIndex        =   3
         ToolTipText     =   "Confirm the password here"
         Top             =   1890
         Width           =   2895
      End
      Begin VB.TextBox txtDefaultPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1935
         PasswordChar    =   "#"
         TabIndex        =   2
         ToolTipText     =   "Enter username's password here"
         Top             =   1485
         Width           =   2895
      End
      Begin VB.TextBox txtDefaultUsername 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1935
         TabIndex        =   1
         ToolTipText     =   "Enter Default Username here"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   $"Form4.frx":000C
         Height          =   690
         Index           =   3
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Confirm p&assword :"
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   6
         Top             =   1980
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Default &password :"
         Height          =   210
         Index           =   1
         Left            =   465
         TabIndex        =   5
         Top             =   1575
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Default &username :"
         Height          =   210
         Index           =   0
         Left            =   495
         TabIndex        =   4
         Top             =   1170
         Width           =   1365
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   510
      Left            =   3330
      TabIndex        =   8
      Top             =   2520
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
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
      MICON           =   "Form4.frx":0094
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   510
      Left            =   1530
      TabIndex        =   9
      Top             =   2520
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
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
      MICON           =   "Form4.frx":00B0
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private K() As cControlFlater, i As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
   On Error GoTo cmdOK_Click_Error

    With Form4
        Dim SelectedPos As String
        Dim SearchLoop As Integer
        Dim TempUser As String
        Dim TempPass As String
        If Len(.txtDefaultUsername.Text) < 3 Then
            MsgBox "Default Username contains at least three characters.", vbInformation, "Username"
            .txtDefaultUsername.SelStart = 0
            .txtDefaultUsername.SelLength = Len(.txtDefaultUsername.Text)
            .txtDefaultUsername.SetFocus
            GoTo GetOut
        End If
        If Len(.txtDefaultPassword.Text) < 3 Then
            MsgBox "Default Password contains at least three characters.", vbInformation, "Password"
            .txtDefaultPassword.SelStart = 0
            .txtDefaultPassword.SelLength = Len(.txtDefaultPassword.Text)
            .txtDefaultPassword.SetFocus
            GoTo GetOut
        End If
        If .txtDefaultPassword.Text <> .txtConfirmPassword.Text Then
            MsgBox "Wrong password confirmation.", vbInformation, "Confirm Password"
            .txtConfirmPassword.SelStart = 0
            .txtConfirmPassword.SelLength = Len(.txtConfirmPassword.Text)
            .txtConfirmPassword.SetFocus
            GoTo GetOut
        End If

        InputUser = .txtDefaultUsername.Text
        InputPass = .txtDefaultPassword.Text

        SelectedPos = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "")
        If UCase(SelectedPos) = "NOT EXISTS" Then GoTo GetOut
        SelectedPos = DecryptCount(SelectedPos) - 1984

        TempUser = DecryptUsername(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & SelectedPos))
        TempPass = DecryptPassword(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & SelectedPos))

        If InputUser = TempUser And InputPass = TempPass Then
            Form1.cmdShowList.Caption = "OK"
            Unload Me
            Exit Sub
        Else
            MsgBox "Your input doesn't match with Default Username & Password!" & vbCrLf & vbCrLf & "If you dont know what Default Username & Password is, you must remember the Username that printed BOLD on the list." & vbCrLf & "If you never set Default Username & Password manually, you can enter your first password because in some cases the Default Username & Password is the first Username & Password on the list." & vbCrLf & "And please check your typing, because each input is case-sensitive.", vbInformation, "Default Username & Password"
            Unload Me
            Exit Sub
        End If

    End With
    Me.MousePointer = 0
    Exit Sub
GetOut:
    Me.MousePointer = 0
    Exit Sub

   On Error GoTo 0
   Exit Sub

cmdOK_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure cmdOK_Click of Form Form4. Please report this to the author of this program."
End Sub

Private Sub Form_Load()
    Dim CTL As Control
    For Each CTL In Me.Controls
        Select Case TypeName(CTL)
            Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox", "PictureBox"
                ReDim Preserve K(i)
                Set K(i) = New cControlFlater
                K(i).Attach CTL
                i = i + 1
        End Select
    Next CTL
End Sub


