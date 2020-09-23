VERSION 5.00
Object = "{475DDC32-C5E4-460C-B00C-A0DAC2FE6D71}#3.0#0"; "bsFrame.ocx"
Begin VB.Form Form1 
   Caption         =   "bsFrame demo"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About the control"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Frames are enabled"
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "A standard, boring-ass intrinsic Frame control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "Hi, I'm an intrinsic check box inside a standard, boring-ass intrinsic frame control"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
   Begin bsFrameOCX.bsFrame bsFrame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3413
      Caption         =   "Change this bsFrame control"
      FlatBorderColour=   -2147483632
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check2 
         Caption         =   "This intrinsic check box is much happier, now it's in a stylish bsFrame control"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4560
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "BadSoft"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Green"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "System colours"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   360
         List            =   "frmMain.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":0484
         Left            =   360
         List            =   "frmMain.frx":04A0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame style"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border style"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check3_Click()
   bsFrame1.Enabled = Check3.Value
   Frame1.Enabled = Check3.Value
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub Combo1_Click()
   bsFrame1.BorderStyle = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
   bsFrame1.FrameStyle = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
   'Set to system colours
   With bsFrame1
      .HighlightColour = vb3DHighlight
      .HighlightDKColour = vb3DLight
      .ShadowColour = vb3DShadow
      .ShadowDKColour = vb3DDKShadow
      .BackColour = vbButtonFace
      .FlatBorderColour = vbBlack
      .Captioncolour = vbButtonText
   End With
End Sub

Private Sub Command2_Click()
   'Set to green...
   'NOTE: If using hex values like below, it's imperative that
   'you put a 2 after the H. Otherwise, things get messy. Don't
   'ask me why.
   With bsFrame1
      .HighlightColour = &HC0FFC0
      .HighlightDKColour = &H80FF80
      .ShadowColour = &H2C000
      .ShadowDKColour = &H28000
      .BackColour = &H2FF00
      .FlatBorderColour = &H24000
      .Captioncolour = &H28000
   End With
End Sub

Private Sub Command3_Click()
   'Set to blue...
   With bsFrame1
      .HighlightColour = &HFFC0C0
      .HighlightDKColour = &HFF8080
      .ShadowColour = &H2C00000
      .ShadowDKColour = &H2800000
      .BackColour = &H2FF0000
      .FlatBorderColour = &H2400000
      .Captioncolour = &H2FFC0C0
   End With
End Sub

Private Sub Command4_Click()
   'BadSoft's custom colours.
   With bsFrame1
      .HighlightColour = &HFFFFFF
      .HighlightDKColour = &HFFEEEE
      .ShadowColour = &H2FFAAAA
      .ShadowDKColour = &H2FF9999
      .BackColour = &H2FFDEDE
      .FlatBorderColour = &H2FFAAAA
      .Captioncolour = &H0
   End With
End Sub

Private Sub Command5_Click()
   bsFrame1.About
End Sub

Private Sub Form_Load()
   Combo1.ListIndex = 6
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub
