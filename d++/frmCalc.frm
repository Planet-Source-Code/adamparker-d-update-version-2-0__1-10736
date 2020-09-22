VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   2970
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   4080
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2970
      Left            =   0
      ScaleHeight     =   2970
      ScaleWidth      =   840
      TabIndex        =   20
      Top             =   0
      Width           =   840
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Calculator"
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
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   230
         Picture         =   "frmCalc.frx":0000
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   2880
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   2880
      TabIndex        =   14
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   3480
      TabIndex        =   15
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      TabIndex        =   18
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   2880
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   19
      Top             =   105
      Width           =   3000
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1994 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
Option Explicit
Dim Op1, Op2
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim LastInput
Dim OpFlag
Dim TempReadout

Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

Private Sub Form_Load()
    frmCalc.Icon = frmMain.Icon
    Image1 = frmMain.Icon
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
End Sub

Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub

