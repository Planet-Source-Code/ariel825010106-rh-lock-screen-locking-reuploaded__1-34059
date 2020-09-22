VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2070
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5835
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
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   0
      Left            =   0
      Picture         =   "Form6.frx":000C
      ScaleHeight     =   1800
      ScaleWidth      =   5805
      TabIndex        =   0
      Top             =   0
      Width           =   5835
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[RH] Lock is Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3960
      TabIndex        =   1
      Top             =   1845
      Width           =   1875
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    If App.PrevInstance = True Then End
    Dim A, B, C, D
    A = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
    B = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "")
    C = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 01") - 18
    D = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 02") - 7
    
    If UCase(A) = "NOT EXISTS" Or UCase(B) = "NOT EXISTS" Or UCase(C) = "NOT EXISTS" Or UCase(D) = "NOT EXISTS" Then InstallDefaultSettings
    Dim Z
    Dim UserCommand As String
    Z = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 02") - 7
    If CInt(Z) <> 0 Then
        Form6.Show
    End If
        Load Form1
    UserCommand = UCase(Command)
    If UserCommand = "/T" Then
        Load Form1
        GoToTray
        Exit Sub
    ElseIf UserCommand = "/L" Then
        Load Form1
        GoToTray
        ActivateLock
        Exit Sub
    End If
    Load Form1
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form6. Please report this to the author of this program."
End Sub
