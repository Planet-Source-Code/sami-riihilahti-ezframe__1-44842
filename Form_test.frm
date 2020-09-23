VERSION 5.00
Begin VB.Form Form_test 
   Caption         =   "EZFrame 1.0"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin EZFrameControl.EZFrame Frame_settings 
      Height          =   3795
      Left            =   360
      Top             =   360
      Width           =   6735
      _extentx        =   11880
      _extenty        =   6694
      caption         =   "Settings"
      textboxheight   =   18
      alignment       =   2
      font            =   "Form_test.frx":000C
      Begin VB.CommandButton CMD_bcrandom 
         Caption         =   "Randomize"
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox Combo_a 
         Height          =   360
         ItemData        =   "Form_test.frx":0034
         Left            =   2280
         List            =   "Form_test.frx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   1815
      End
      Begin VB.CommandButton CMD_tcrandom 
         Caption         =   "Randomize"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton CMD_fcrandom 
         Caption         =   "Randomize"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   2700
         Width           =   1815
      End
      Begin VB.CommandButton CMD_tbcrandom 
         Caption         =   "Randomize"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1860
         Width           =   1815
      End
      Begin VB.TextBox Text_settings 
         Height          =   360
         Left            =   2280
         TabIndex        =   7
         Text            =   "Settings"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Combo_tbh 
         Height          =   360
         ItemData        =   "Form_test.frx":005A
         Left            =   2280
         List            =   "Form_test.frx":0070
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         Height          =   255
         Left            =   300
         TabIndex        =   17
         Top             =   3180
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TextColor"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FrameColor"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxColor"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "TextBoxHeight"
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   1500
         Width           =   1455
      End
   End
   Begin EZFrameControl.EZFrame EZFrame1 
      Height          =   2115
      Left            =   360
      Top             =   4260
      Width           =   6735
      _extentx        =   11880
      _extenty        =   3731
      framecolor      =   8421504
      textboxcolor    =   16744576
      caption         =   "About EZFrame"
      textboxheight   =   18
      textcolor       =   16777215
      alignment       =   2
      font            =   "Form_test.frx":008B
      Begin VB.Label Label4 
         Caption         =   "By ElectroZ"
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   1620
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Version 1.0"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   1260
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "Free to use. Free to modify."
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   900
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Lightweight VB control demonstrating user control topology."
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   540
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD_bcrandom_Click()
    Frame_settings.BackColor = Int(Rnd * 16777216)
End Sub

Private Sub CMD_fcrandom_Click()
    Frame_settings.FrameColor = Int(Rnd * 16777216)
End Sub

Private Sub CMD_tbcrandom_Click()
    Frame_settings.TextBoxColor = Int(Rnd * 16777216)
End Sub

Private Sub CMD_tcrandom_Click()
    Frame_settings.TextColor = Int(Rnd * 16777216)
End Sub

Private Sub Combo_a_Click()
    Frame_settings.Alignment = Combo_a.ListIndex
End Sub

Private Sub Combo_tbh_Click()
    Frame_settings.TextBoxHeight = Val(Combo_tbh.Text)
End Sub

Private Sub Text_settings_Change()
    Frame_settings.Caption = Text_settings.Text
End Sub
