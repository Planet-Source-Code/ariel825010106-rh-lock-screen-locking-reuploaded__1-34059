VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[RH] Lock - Add password"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   " Add password "
      Height          =   2490
      Left            =   82
      TabIndex        =   0
      Top             =   45
      Width           =   4980
      Begin VB.CheckBox chkAsDefault 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Make this username and password as default"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1215
         TabIndex        =   7
         ToolTipText     =   "Check here if you want this username & password as default"
         Top             =   1530
         Width           =   3615
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1935
         PasswordChar    =   "#"
         TabIndex        =   6
         ToolTipText     =   "Confirm the password here"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtNewPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1935
         PasswordChar    =   "#"
         TabIndex        =   4
         ToolTipText     =   "Enter username's password here"
         Top             =   675
         Width           =   2895
      End
      Begin VB.TextBox txtNewUsername 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1935
         TabIndex        =   2
         ToolTipText     =   "Enter your new username here"
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   $"Form3.frx":000C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Confirm p&assword :"
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Username's &password :"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   765
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "New &username :"
         Height          =   210
         Index           =   0
         Left            =   660
         TabIndex        =   1
         Top             =   360
         Width           =   1200
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   510
      Left            =   3330
      TabIndex        =   9
      Top             =   2655
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
      MICON           =   "Form3.frx":00D4
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   510
      Left            =   1530
      TabIndex        =   10
      Top             =   2655
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
      MICON           =   "Form3.frx":00F0
   End
End
Attribute VB_Name = "Form3"
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

    Form3.MousePointer = 11
    Dim ThisIsTempVar
    Dim TempVar As Integer
    With Form3
        If Len(.txtNewUsername.Text) < 3 Then
            MsgBox "Username must be at least three characters.", vbInformation, "Username"
            .txtNewUsername.SelStart = 0
            .txtNewUsername.SelLength = Len(.txtNewUsername.Text)
            .txtNewUsername.SetFocus
            GoTo GetOut
        End If
        If Len(.txtNewPassword.Text) < 3 Then
            MsgBox "Password must be at least three characters.", vbInformation, "Password"
            .txtNewPassword.SelStart = 0
            .txtNewPassword.SelLength = Len(.txtNewPassword.Text)
            .txtNewPassword.SetFocus
            GoTo GetOut
        End If
        If .txtNewPassword.Text <> .txtConfirmPassword.Text Then
            MsgBox "Wrong password confirmation.", vbInformation, "Confirm Password"
            .txtConfirmPassword.SelStart = 0
            .txtConfirmPassword.SelLength = Len(.txtConfirmPassword.Text)
            .txtConfirmPassword.SetFocus
            GoTo GetOut
        End If

        With Form1
            ThisIsTempVar = .lstPasswordList.ListItems.Add(.lstPasswordList.ListItems.Count + 1, , Form3.txtNewUsername)
            .lstPasswordList.ListItems.Item(.lstPasswordList.ListItems.Count).SubItems(1) = Form3.txtNewPassword.Text

            If .lstPasswordList.ListItems.Count <= 1 Then
                .lstPasswordList.ListItems(1).Bold = True
                .lstPasswordList.ListItems(1).Selected = True
            Else
                If Form3.chkAsDefault.Value = 1 Then
                    For TempVar = 1 To .lstPasswordList.ListItems.Count
                        .lstPasswordList.ListItems.Item(TempVar).Bold = False
                    Next TempVar
                    .lstPasswordList.ListItems.Item(Form1.lstPasswordList.ListItems.Count).Bold = True
                    .lstPasswordList.ListItems.Item(Form1.lstPasswordList.ListItems.Count).Selected = True
                End If
            End If
        End With
        MsgBox "New username and password successfully added.", vbInformation, "Password"
    End With
    Unload Form3
    Exit Sub
GetOut:
    Form3.MousePointer = 0
    Exit Sub

   On Error GoTo 0
   Exit Sub

cmdOK_Click_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure cmdOK_Click of Form Form3. Please report this to the author of this program."
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


