VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5550
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   0
      Width           =   855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Help File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   5040
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   230
         Picture         =   "frmHelp.frx":0000
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0442
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Image1.Picture = frmMain.Icon
End Sub
