VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1905
   ClientLeft      =   -2.45715e5
   ClientTop       =   -99960
   ClientWidth     =   3090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRyoko4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   1575
      Picture         =   "Form5.frx":000C
      ScaleHeight     =   660
      ScaleWidth      =   840
      TabIndex        =   2
      Top             =   585
      Width           =   870
   End
   Begin VB.PictureBox picRyoko3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   660
      Picture         =   "Form5.frx":19AEF
      ScaleHeight     =   660
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   592
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   $"Form5.frx":491F9
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1350
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUMP FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   0
      Width           =   2670
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
