VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{393A9FC2-A107-4CAD-B13B-77B06BA97134}#1.0#0"; "[RH] SYSTRAY.OCX"
Object = "{252BC880-5111-4AFE-95F4-0201E70F34CA}#1.0#0"; "MCLHOTKEY.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[RH] Lock"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGeneral 
      Caption         =   " [ General ] "
      Height          =   3660
      Left            =   90
      TabIndex        =   0
      Top             =   2340
      Width           =   5775
      Begin VB.CheckBox chkHotkey 
         Appearance      =   0  'Flat
         Caption         =   "A&ctivate [RH] Lock through hotkey press [ Ctrl + Shift + Alt + F12 ]"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   1125
         Width           =   5460
      End
      Begin VB.CheckBox chkUseSplash 
         Appearance      =   0  'Flat
         Caption         =   "S&how splash screen on [RH] Lock startup"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   765
         Value           =   1  'Checked
         Width           =   5460
      End
      Begin VB.CheckBox chkStartup 
         Appearance      =   0  'Flat
         Caption         =   "L&oad [RH] Lock on computer startup (recommended)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Value           =   1  'Checked
         Width           =   5460
      End
      Begin VB.Label Label1 
         Caption         =   "NOTE : Disable this option if the hotkey [ Ctrl + Shift + Alt + F12 ] already reserved by other applications."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   450
         TabIndex        =   9
         Top             =   1440
         Width           =   5010
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   " | About | "
      Enabled         =   0   'False
      Height          =   3660
      Left            =   90
      TabIndex        =   36
      Top             =   2340
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox lstProgInfo 
         Appearance      =   0  'Flat
         Height          =   870
         ItemData        =   "Form1.frx":49E2
         Left            =   180
         List            =   "Form1.frx":4A13
         TabIndex        =   39
         Top             =   2610
         Width           =   5415
      End
      Begin prjChameleon.chameleonButton cmdRyokoHirosue 
         Height          =   465
         Left            =   180
         TabIndex        =   37
         ToolTipText     =   "Information about Ryoko Hirosue"
         Top             =   270
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "Ryoko Hirosue"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   12632064
         FCOL            =   0
         FCOLO           =   0
         MPTR            =   0
         MICON           =   "Form1.frx":4CC2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Form1.frx":4CDE
         ForeColor       =   &H80000008&
         Height          =   1770
         Index           =   5
         Left            =   180
         TabIndex        =   38
         Top             =   765
         Width           =   5415
      End
   End
   Begin VB.Frame fraPassword 
      Caption         =   " | Password | "
      Enabled         =   0   'False
      Height          =   3660
      Left            =   90
      TabIndex        =   10
      Top             =   2340
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkMaskPassword 
         Appearance      =   0  'Flat
         Caption         =   "Ma&sk my password (eg. ""myp455"" into ""######"")"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   3285
         Value           =   1  'Checked
         Width           =   5460
      End
      Begin prjChameleon.chameleonButton cmdShowList 
         Height          =   375
         Left            =   4275
         TabIndex        =   13
         Top             =   1125
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Show list"
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
         MICON           =   "Form1.frx":4F59
      End
      Begin VB.Frame Frame2 
         Caption         =   " Password(s) list "
         Height          =   1905
         Left            =   135
         TabIndex        =   14
         Top             =   1215
         Visible         =   0   'False
         Width           =   5505
         Begin MSComctlLib.ListView lstPasswordList 
            Height          =   1365
            Left            =   135
            TabIndex        =   15
            Top             =   360
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   2408
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Username :"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Password :"
               Object.Width           =   2540
            EndProperty
         End
         Begin prjChameleon.chameleonButton cmdAdd 
            Height          =   375
            Left            =   4230
            TabIndex        =   16
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "[ + A&dd ]"
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
            MICON           =   "Form1.frx":4F75
         End
         Begin prjChameleon.chameleonButton cmdRemoveSel 
            Height          =   375
            Left            =   4230
            TabIndex        =   17
            Top             =   810
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "[ - &Remove ]"
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
            MICON           =   "Form1.frx":4F91
         End
         Begin prjChameleon.chameleonButton cmdClear 
            Height          =   375
            Left            =   4230
            TabIndex        =   18
            Top             =   1350
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "[ &Clear ]"
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
            MICON           =   "Form1.frx":4FAD
         End
      End
      Begin VB.CheckBox chkUseMultipleUser 
         Appearance      =   0  'Flat
         Caption         =   "&Enable multiple username, and password (password for others)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   765
         Value           =   1  'Checked
         Width           =   5460
      End
      Begin VB.CheckBox chkPassSensitive 
         Appearance      =   0  'Flat
         Caption         =   "&My password(s) are case-sensitive (eg. uppercase and undercase)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   405
         Value           =   1  'Checked
         Width           =   5460
      End
   End
   Begin VB.Frame fraLock 
      Caption         =   " | Lock | "
      Enabled         =   0   'False
      Height          =   3660
      Left            =   90
      TabIndex        =   20
      Top             =   2340
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkUseMaxTry 
         Appearance      =   0  'Flat
         Caption         =   "&Use maximum try (eg. Three tries and shutdown)"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   270
         TabIndex        =   30
         Top             =   2610
         Width           =   3840
      End
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   135
         TabIndex        =   29
         Top             =   2610
         Width           =   5505
         Begin VB.ComboBox cmbAction 
            Height          =   330
            ItemData        =   "Form1.frx":4FC9
            Left            =   3375
            List            =   "Form1.frx":4FD9
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   2040
         End
         Begin VB.ComboBox cmbNumberOfTry 
            Height          =   330
            ItemData        =   "Form1.frx":5009
            Left            =   1215
            List            =   "Form1.frx":5067
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Kind of action :"
            Height          =   210
            Index           =   4
            Left            =   2205
            TabIndex        =   34
            Top             =   405
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Maximum try :"
            Height          =   210
            Index           =   3
            Left            =   135
            TabIndex        =   32
            Top             =   405
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Background picture "
         Height          =   1455
         Left            =   135
         TabIndex        =   24
         Top             =   1080
         Width           =   5505
         Begin VB.TextBox txtBGLocation 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1935
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   225
            Width           =   2940
         End
         Begin VB.PictureBox picPrevBG 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   1140
            Left            =   135
            ScaleHeight     =   1110
            ScaleWidth      =   1695
            TabIndex        =   25
            Top             =   225
            Width           =   1725
            Begin VB.Image imgBGPicture 
               Height          =   1095
               Left            =   315
               Top             =   0
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin prjChameleon.chameleonButton cmdBGBrowse 
            Height          =   330
            Left            =   4950
            TabIndex        =   27
            Top             =   225
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            BTYPE           =   3
            TX              =   "..."
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
            MICON           =   "Form1.frx":50DA
         End
         Begin VB.Label Label1 
            Caption         =   "NOTE : If the picture not exists, a picture of Ryoko Hirosue will be used."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1890
            TabIndex        =   28
            Top             =   630
            Width           =   3525
         End
      End
      Begin VB.CheckBox chkShowLockText 
         Appearance      =   0  'Flat
         Caption         =   "S&how information that this computer is being locked"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   765
         Width           =   5460
      End
      Begin VB.ComboBox cmbBackground 
         Height          =   330
         ItemData        =   "Form1.frx":50F6
         Left            =   2430
         List            =   "Form1.frx":5106
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   315
         Width           =   3210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&When [RH] Lock active show :"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   405
         Width           =   2190
      End
   End
   Begin MCLHotkey.VBHotKey VBHotKey1 
      Left            =   5355
      Top             =   45
      _ExtentX        =   794
      _ExtentY        =   794
      AltKey          =   -1  'True
      ShiftKey        =   -1  'True
      CtrlKey         =   -1  'True
      VKey            =   123
      WinKey          =   0   'False
      Enabled         =   0   'False
   End
   Begin RHSystray.Icon Icon1 
      Left            =   5355
      Top             =   45
      _ExtentX        =   900
      _ExtentY        =   820
   End
   Begin VB.PictureBox picSystray 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   90
      Picture         =   "Form1.frx":5159
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   41
      Top             =   6030
      Visible         =   0   'False
      Width           =   270
   End
   Begin prjChameleon.chameleonButton cmdGeneral 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "General settings"
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&General"
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
      MICON           =   "Form1.frx":52A3
   End
   Begin prjChameleon.chameleonButton cmdPassword 
      Height          =   375
      Left            =   1590
      TabIndex        =   3
      ToolTipText     =   "Password-related settings"
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Password"
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
      MICON           =   "Form1.frx":52BF
   End
   Begin prjChameleon.chameleonButton cmdLock 
      Height          =   375
      Left            =   3105
      TabIndex        =   4
      ToolTipText     =   "Lock Interface settings"
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Lock"
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
      MICON           =   "Form1.frx":52DB
   End
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   375
      Left            =   4635
      TabIndex        =   5
      ToolTipText     =   "Information about [RH] Lock"
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&About"
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
      MICON           =   "Form1.frx":52F7
   End
   Begin VB.PictureBox Picture1 
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
      Left            =   45
      Picture         =   "Form1.frx":5313
      ScaleHeight     =   1800
      ScaleWidth      =   5865
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   510
      Left            =   4140
      TabIndex        =   35
      Top             =   6075
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "Exit [RH] Lock"
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
      MICON           =   "Form1.frx":A5C7
   End
   Begin prjChameleon.chameleonButton cmdMinimizetoTray 
      Height          =   510
      Left            =   2340
      TabIndex        =   40
      Top             =   6075
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
      BTYPE           =   3
      TX              =   "Go &to tray"
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
      MICON           =   "Form1.frx":A5E3
   End
   Begin VB.Menu mnuRHLockMenu 
      Caption         =   "[RH] Lock Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuActivateLock 
         Caption         =   "Activate [RH]  &Lock"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Open [RH] Lock settings"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnload 
         Caption         =   "E&xit [RH] Lock"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SettingChanged As Boolean
Dim PassListChanged As Boolean

Private K() As cControlFlater, i As Integer
Public Sub PleaseProcessBackground(ThePictureControl As Control)
   On Error GoTo PleaseProcessBackground_Error

    If ThePictureControl.Width > Screen.Width Then
        If ThePictureControl.Width > ThePictureControl.Height Then
            VarRatio = ThePictureControl.Width / Screen.Width
            ThePictureControl.Width = ThePictureControl.Width / VarRatio
            ThePictureControl.Height = ThePictureControl.Height / VarRatio
        End If

        If ThePictureControl.Height > ThePictureControl.Width Then
            VarRatio = ThePictureControl.Height / Screen.Height
            ThePictureControl.Height = ThePictureControl.Height / VarRatio
            ThePictureControl.Width = ThePictureControl.Width / VarRatio
        End If
    End If
    If ThePictureControl.Height > Screen.Height Then
        If ThePictureControl.Height > ThePictureControl.Width Then
            VarRatio = ThePictureControl.Width / Screen.Width
            ThePictureControl.Width = ThePictureControl.Height / VarRatio
            ThePictureControl.Height = ThePictureControl.Width / VarRatio
        End If
        If ThePictureControl.Width > ThePictureControl.Height Then
            VarRatio = ThePictureControl.Height / Screen.Height
            ThePictureControl.Height = ThePictureControl.Height / VarRatio
            ThePictureControl.Width = ThePictureControl.Width / VarRatio
        End If
    End If

    If ThePictureControl.Width >= Screen.Width Then ThePictureControl.Left = 0
    If ThePictureControl.Height >= Screen.Height Then ThePictureControl.Top = 0

    If ThePictureControl.Width < Screen.Width Then
        ThePictureControl.Left = (Screen.Width / 2) - (ThePictureControl.Width / 2)
    End If

    If ThePictureControl.Height < Screen.Height Then
        ThePictureControl.Top = (Screen.Height / 2) - (ThePictureControl.Height / 2)
    End If

   On Error GoTo 0
   Exit Sub

PleaseProcessBackground_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure PleaseProcessBackground of Form Form1. Please report this to the author of this program."
End Sub

Public Sub SetBG0()
   On Error GoTo SetBG0_Error

    Form2.MousePointer = 11
    Dim BGFilename As String
    Form2.imgBackground.Visible = False
    BGFilename = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
    If FileExists(BGFilename) = False Then
        SetBG3
        Exit Sub
    End If
    With Form2
        .imgBackground.Picture = LoadPicture(BGFilename)
        .imgBackground.Stretch = False
        PleaseProcessBackground .imgBackground
        .imgBackground.Stretch = True
    End With
    Form2.imgBackground.Visible = True
    Form2.MousePointer = 0

   On Error GoTo 0
   Exit Sub

SetBG0_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SetBG0 of Form Form1. Please report this to the author of this program."
End Sub
Public Sub SetBG1()
   On Error GoTo SetBG1_Error

    Form2.MousePointer = 11
    Dim BGFilename As String
    Form2.imgBackground.Visible = False
    BGFilename = Form1.txtBGLocation.Text
    If FileExists(BGFilename) = False Then
        SetBG3
        Exit Sub
    End If
    Form2.imgBackground.Picture = LoadPicture(BGFilename)
    Form2.imgBackground.Stretch = False
    PleaseProcessBackground Form2.imgBackground
    Form2.imgBackground.Stretch = True

    Form2.imgBackground.Visible = True
    Form2.MousePointer = 0

   On Error GoTo 0
   Exit Sub

SetBG1_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SetBG1 of Form Form1. Please report this to the author of this program."
End Sub
Public Sub SetBG2()
   On Error GoTo SetBG2_Error

    Form2.MousePointer = 11
    With Form2
        .imgBackground.Picture = LoadPicture("")
        .imgBackground.Top = 0
        .imgBackground.Left = 0
        .imgBackground.Stretch = False
        .imgBackground.Visible = True
    End With
    Form2.MousePointer = 0

   On Error GoTo 0
   Exit Sub

SetBG2_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SetBG2 of Form Form1. Please report this to the author of this program."
End Sub
Public Sub SetBG3()
   On Error GoTo SetBG3_Error

    Form2.MousePointer = 11
    Form2.imgBackground.Visible = False
    With Form5
        Form2.imgBackground.Picture = LoadPicture("")
        If Second(Now) < 30 Then
            Form2.imgBackground.Picture = .picRyoko3.Picture
        End If

        If Second(Now) > 30 Then
            Form2.imgBackground.Picture = .picRyoko4.Picture
        End If

        PleaseProcessBackground Form2.imgBackground
        Form2.imgBackground.Top = 0
        Form2.imgBackground.Stretch = False
        Form2.imgBackground.Visible = True
    End With
    Form2.MousePointer = 0

   On Error GoTo 0
   Exit Sub

SetBG3_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure SetBG3 of Form Form1. Please report this to the author of this program."
End Sub

Public Function ViewPictureFile(PictureFilename As String) As Boolean
    Dim VarRatio
   On Error GoTo ViewPictureFile_Error

    ViewPictureFile = False
    If Len(PictureFilename) <= 2 Then Exit Function
    If FileExists(PictureFilename) = False Then Exit Function
    Form1.MousePointer = 11
    With Form1.imgBGPicture
        .Visible = False
        .Stretch = False
        .Picture = LoadPicture(PictureFilename)
        If .Width > picPrevBG.ScaleWidth Then
            If .Width > .Height Then
                VarRatio = .Width / picPrevBG.ScaleWidth
                .Width = .Width / VarRatio
                .Height = .Height / VarRatio
            End If

            If .Height > .Width Then
                VarRatio = .Height / picPrevBG.ScaleHeight
                .Height = .Height / VarRatio
                .Width = .Width / VarRatio
            End If
        End If
        If .Height > picPrevBG.ScaleHeight Then
            If .Height > .Width Then
                VarRatio = .Width / picPrevBG.ScaleWidth
                .Width = .Height / VarRatio
                .Height = .Width / VarRatio
            End If
            If .Width > .Height Then
                VarRatio = .Height / picPrevBG.ScaleHeight
                .Height = .Height / VarRatio
                .Width = .Width / VarRatio
            End If
        End If

        If .Width - 50 <> 0 And .Height - 50 <> 0 Then
            .Width = .Width - 50
            .Height = .Height - 50
        End If

        If .Width >= picPrevBG.ScaleWidth Then .Left = 0
        If .Height >= picPrevBG.ScaleHeight Then .Top = 0

        If .Width < picPrevBG.ScaleWidth Then
            .Left = (picPrevBG.ScaleWidth / 2) - (.Width / 2)
        End If

        If .Height < picPrevBG.ScaleHeight Then
            .Top = (picPrevBG.ScaleHeight / 2) - (.Height / 2)
        End If
        .Stretch = True
        .Visible = True
    End With
    ViewPictureFile = True
    Form1.MousePointer = 0

   On Error GoTo 0
   Exit Function

ViewPictureFile_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure ViewPictureFile of Form Form1. Please report this to the author of this program."
End Function
Private Sub chkHotkey_Click()
    Form1.VBHotKey1.Enabled = False
    If Form1.chkHotkey.Value = 1 Then Form1.VBHotKey1.Enabled = True
    SettingChanged = True
End Sub


Private Sub chkMaskPassword_Click()
    SettingChanged = True
End Sub

Private Sub chkPassSensitive_Click()
    SettingChanged = True
End Sub


Private Sub chkShowLockText_Click()
    SettingChanged = True
End Sub

Private Sub chkStartup_Click()
    SettingChanged = True
End Sub

Private Sub chkUseMaxTry_Click()
    SettingChanged = True
    If chkUseMaxTry.Value = 1 Then
        Frame3.Enabled = True
    End If

    If chkUseMaxTry.Value = 0 Then
        Frame3.Enabled = False
    End If
End Sub

Private Sub chkUseMultipleUser_Click()
    SettingChanged = True
End Sub

Private Sub chkUseSplash_Click()
    SettingChanged = True
End Sub

Private Sub cmbAction_Change()
    SettingChanged = True
End Sub

Private Sub cmbBackground_Change()
    SettingChanged = True
End Sub

Private Sub cmbNumberOfTry_Change()
    SettingChanged = True
End Sub

Private Sub cmdAbout_Click()
    With Form1
        .fraGeneral.Enabled = False
        .fraPassword.Enabled = False
        .fraLock.Enabled = False
        .fraAbout.Enabled = True

        .fraGeneral.Visible = False
        .fraPassword.Visible = False
        .fraLock.Visible = False
        .fraAbout.Visible = True
    End With
End Sub

Private Sub cmdAdd_Click()
    BeforeAdd = Form1.lstPasswordList.ListItems.Count
    Form3.Show vbModal, Form1
    AfterAdd = Form1.lstPasswordList.ListItems.Count
    If BeforeAdd < AfterAdd Then
        SettingChanged = True
        PassListChanged = True
    End If
End Sub

Private Sub cmdBGBrowse_Click()
    Dim MyDialog As cNewDialog
    Dim OldFilename As String
    Set MyDialog = New cNewDialog
    OldFilename = Form1.txtBGLocation.Text
    With MyDialog
        .DialogTitle = "Open background picture"
        .FileFlags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
        .Filter = "All supported format|*.jpg;*.jpeg;*.jpe;*.gif;*.bmp|All files|*.*"
        .hwnd = Form1.hwnd
        .ShowOpen
        If ViewPictureFile(.FileName) = False Then
            Form1.txtBGLocation.Text = OldFilename
        Else
            Form1.txtBGLocation.Text = .FileName
        End If
    End With

    Set MyDialog = Nothing
    SettingChanged = True
End Sub

Private Sub cmdClear_Click()
    Dim AskUser
    AskUser = MsgBox("Are you sure you want to clear all username and password?", vbQuestion + vbYesNo, "Clear")
    If AskUser = vbYes Then
        DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys"
        SaveKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys"
        PasswordCount = EncryptCount(Form1.lstPasswordList.ListItems.Count + 1980)
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID", "0"
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "", "0"
        lstPasswordList.ListItems.Remove (lstPasswordList.SelectedItem.Index)
        PassListChanged = True
        SettingChanged = True
    End If
End Sub

Private Sub cmdGeneral_Click()
    With Form1
        .fraGeneral.Enabled = True
        .fraPassword.Enabled = False
        .fraLock.Enabled = False
        .fraAbout.Enabled = False

        .fraGeneral.Visible = True
        .fraPassword.Visible = False
        .fraLock.Visible = False
        .fraAbout.Visible = False
    End With
End Sub


Private Sub cmdLock_Click()
    With Form1
        .fraGeneral.Enabled = False
        .fraPassword.Enabled = False
        .fraLock.Enabled = True
        .fraAbout.Enabled = False

        .fraGeneral.Visible = False
        .fraPassword.Visible = False
        .fraLock.Visible = True
        .fraAbout.Visible = False
    End With
End Sub

Private Sub cmdMinimizetoTray_Click()
    GoToTray
End Sub

Private Sub cmdOK_Click()
    If SettingChanged = True Then
        SaveRHSettings
        SettingChanged = False
    End If
    ExitFromRHLock
End Sub

Private Sub cmdPassword_Click()
    Dim PasswordCount
    With Form1
        .fraGeneral.Enabled = False
        .fraPassword.Enabled = True
        .fraLock.Enabled = False
        .fraAbout.Enabled = False

        .fraGeneral.Visible = False
        .fraPassword.Visible = True
        .fraLock.Visible = False
        .fraAbout.Visible = False
        PasswordCount = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
        If UCase(PasswordCount) = "NOT EXISTS" Or UCase(SelectedPos) = "NOT EXISTS" Then
            cmdShowList_Click
            Exit Sub
        End If
        PasswordCount = DecryptCount(PasswordCount) - 1980
        If PasswordCount < 1 Then cmdShowList_Click
    End With
End Sub


Private Sub cmdRemoveSel_Click()
    Dim AskUser
    AskUser = MsgBox("Are you sure you want to delete this username and password?", vbQuestion + vbYesNo, "Remove")
    If AskUser = vbYes Then
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ClsRH.ID" & Form1.lstPasswordList.SelectedItem.Index
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "LoadClass" & Form1.lstPasswordList.SelectedItem.Index
        lstPasswordList.ListItems.Remove (lstPasswordList.SelectedItem.Index)
        PassListChanged = True
        SettingChanged = True
    End If
End Sub

Private Sub cmdRyokoHirosue_Click()
    ShellExecute Form1.hwnd, "open", "http://www.Ryoko-Hirosue.org", "", "", 1
End Sub

Private Sub cmdShowList_Click()
    Dim TempVar As Integer
    Dim BoldCount As Integer
    Dim PasswordCount

    If LCase(cmdShowList.Caption) = "&show list" Then
        PasswordCount = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\WinProcSys", "ArcID")
        If UCase(PasswordCount) = "NOT EXISTS" Then GoTo ShowTheList
        PasswordCount = DecryptCount(PasswordCount) - 1980
        If PasswordCount < 1 Then GoTo ShowTheList
        Form4.Show vbModal, Form1
        If UCase(cmdShowList.Caption) = "OK" Then
ShowTheList:
            cmdShowList.Caption = "&Show List"
            ReadPasswords
            Frame2.Visible = True
            Frame2.Enabled = True
            cmdShowList.Caption = "&Hide list"
            Exit Sub
        End If
    End If

    If LCase(cmdShowList.Caption) = "&hide list" Then
        For TempVar = 1 To lstPasswordList.ListItems.Count
            If lstPasswordList.ListItems.Item(TempVar).Bold = True Then BoldCount = BoldCount + 1
        Next TempVar
        If lstPasswordList.ListItems.Count < 1 Then
            MsgBox "No username & password detected! Please add at least one password.", vbInformation, "Username & Password"
            cmdAdd_Click
            Exit Sub
        End If
        If BoldCount = 0 Then
            MsgBox "No default Username and Password found! Please select a Username / Password and double-click it to select it as a default Username and Password.", vbInformation, "Default Username and Password"
            Exit Sub
        End If
        If PassListChanged = True Then
            SavePasswords
        End If
        lstPasswordList.ListItems.Clear
        Frame2.Visible = False
        Frame2.Enabled = False
        cmdShowList.Caption = "&Show list"
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    Dim C
   On Error GoTo Form_Load_Error

    C = GetString(HKEY_LOCAL_MACHINE, "Software\Ryoko Hirosue\Lock\Settings", "[RH] Lock - Settings 02") - 7
    C = CInt(C)
    Form6.Label1.Caption = "Reading user settings..."
    ReadRHSettings
    If C <> 0 Then ReadPasswords
    If C <> 0 Then Form1.lstPasswordList.ListItems.Clear

    Dim CTL As Control
    Form6.Label1.Caption = "Setting GUI..."
    If C <> 0 Then Form5.Show
    If C <> 0 Then Unload Form5

    For Each CTL In Me.Controls
        Select Case TypeName(CTL)
            Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox", "PictureBox"
                ReDim Preserve K(i)
                Set K(i) = New cControlFlater
                If C <> 0 Then Form6.Label1.Caption = "Setting " & CTL & "..."
                K(i).Attach CTL
                i = i + 1
        End Select
    Next CTL

    Form6.Label1.Caption = "Finalizing settings..."
    chkUseMaxTry_Click
    D = ViewPictureFile(Form1.txtBGLocation.Text)
    If Form1.chkHotkey.Value = 1 Then Form1.VBHotKey1.Enabled = True
    Unload Form6
    Form1.Show

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error occured. The error number is : " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1. Please report this to the author of this program."
End Sub
Private Sub Icon1_MouseDown(nButton As Integer)
    If nButton = 2 Then
        PopupMenu mnuRHLockMenu
    End If
End Sub


Private Sub lstPasswordList_DblClick()
    Form1.MousePointer = 11
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    With Form1
        If .lstPasswordList.SelectedItem.Index < 0 Or .lstPasswordList.SelectedItem.Index > .lstPasswordList.ListItems.Count Then Exit Sub

        TempVar2 = .lstPasswordList.SelectedItem.Index
        For TempVar = 1 To .lstPasswordList.ListItems.Count
            .lstPasswordList.ListItems.Item(TempVar).Bold = False
        Next TempVar
        .lstPasswordList.ListItems.Item(TempVar2).Bold = True
        .lstPasswordList.ListItems.Item(TempVar2).Selected = True
    End With
    SettingChanged = True
    Form1.MousePointer = 0
End Sub


Private Sub mnuActivateLock_Click()
    ActivateLock
End Sub

Private Sub mnuSettings_Click()
    With Form1
        .Show
        .Icon1.DeleteIcon
    End With
End Sub

Private Sub mnuUnload_Click()
    ExitFromRHLock
End Sub

Private Sub txtBGLocation_Change()
    SettingChanged = True
End Sub


Private Sub VBHotKey1_HotkeyPressed()
    mnuActivateLock_Click
End Sub


