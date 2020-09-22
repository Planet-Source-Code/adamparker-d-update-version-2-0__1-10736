VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "D++ Compiler"
   ClientHeight    =   6975
   ClientLeft      =   3285
   ClientTop       =   2145
   ClientWidth     =   8205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   6975
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6720
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Picture         =   "frmMain.frx":064C
            Text            =   "D++ Compiler"
            TextSave        =   "D++ Compiler"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "3:00 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6390
      Left            =   0
      ScaleHeight     =   6390
      ScaleWidth      =   840
      TabIndex        =   3
      Top             =   330
      Width           =   840
      Begin VB.Image Image1 
         Height          =   480
         Left            =   230
         Picture         =   "frmMain.frx":0B90
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D++ Main Code Window"
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
         Height          =   675
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1422
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":163A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":174E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            Object.ToolTipText     =   "Compile"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDebug 
      Height          =   1335
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4260
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.TextBox txtText 
      Height          =   3855
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Image icoHelp 
      Height          =   240
      Left            =   5400
      Picture         =   "frmMain.frx":1862
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoPaste 
      Height          =   240
      Left            =   5040
      Picture         =   "frmMain.frx":1964
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoCopy 
      Height          =   240
      Left            =   4560
      Picture         =   "frmMain.frx":1A66
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoCut 
      Height          =   240
      Left            =   4200
      Picture         =   "frmMain.frx":1B68
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoSave 
      Height          =   240
      Left            =   3840
      Picture         =   "frmMain.frx":1C6A
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoOpen 
      Height          =   240
      Left            =   3360
      Picture         =   "frmMain.frx":1D6C
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image icoNew 
      Height          =   240
      Left            =   3000
      Picture         =   "frmMain.frx":1E6E
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAS 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnulne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "Time/Date"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewDebug 
         Caption         =   "&Debug Window"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewCalc 
         Caption         =   "Calculator"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectCompile 
         Caption         =   "&Compile"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuProjectRun 
         Caption         =   "&Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuProjectStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelparr 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About D++ Compiler..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim debugactive As Boolean, save As Boolean

Private Sub Form_Load()
SetMenuBitmap Me, Array(0, 0), icoNew.Picture
SetMenuBitmap Me, Array(0, 2), icoOpen.Picture
SetMenuBitmap Me, Array(0, 5), icoSave.Picture
SetMenuBitmap Me, Array(1, 0), icoCut.Picture
SetMenuBitmap Me, Array(1, 1), icoCopy.Picture
SetMenuBitmap Me, Array(1, 2), icoPaste.Picture
SetMenuBitmap Me, Array(4, 0), icoHelp.Picture
Image1.Picture = frmMain.Icon
Me.Show
SetDllLocation
debugactive = False
save = False
If FileExist(GetSystemDirectory & "\DPPAPP.dll") = False Then
    MsgBox "DPPAPP.DLL not found in system folder!", vbCritical, "File Not Found!"
    End
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Line1.X2 = Width
Line2.X2 = Width
If debugactive = True Then
txtDebug.Top = ScaleHeight - txtDebug.Height - 255
txtDebug.Width = ScaleWidth - 955
txtText.Width = ScaleWidth - 955
txtText.Height = ScaleHeight - (txtDebug.Height + 150) - 505
Else
txtText.Width = ScaleWidth - 955
txtText.Height = ScaleHeight - (Toolbar1.Height + 50) - 255
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
Clipboard.SetText txtText.SelText
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
Clipboard.SetText txtText.SelText
txtText.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
txtText.SelText = Clipboard.GetText
End Sub

Private Sub mnuFileClose_Click()
save = False
Me.Caption = "D++"
txtText.Text = ""
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileNew_Click()
save = False
Me.Caption = "D++"
txtText.Text = ""
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo endit
CommonDialog1.Filter = "D++ Files (*.dpp)|*.dpp|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Open D++ File"
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
txtText.Text = ReadFile(CommonDialog1.FileName)
Me.Caption = "D++ Compiler - [" & CommonDialog1.FileTitle & "]"
save = True
endit:
End Sub

Private Sub mnuFileSave_Click()
If save = True Then
Open CommonDialog1.FileName For Output As #1
Print #1, txtText.Text
Close #1
Me.Caption = "D++ Compiler - [" & CommonDialog1.FileTitle & "]"
Else
mnuFileSaveAS_Click
End If
End Sub

Private Sub mnuFileSaveAS_Click()
On Error GoTo endit
CommonDialog1.Filter = "D++ Files (*.dpp)|*.dpp|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Save D++ File"
CommonDialog1.CancelError = True
CommonDialog1.ShowSave
If FileExist(CommonDialog1.FileName) Then
    overwrite = MsgBox("File Exists!  Overwrite?", 276, "File Found!")
    If overwrite = 6 Then
        Open CommonDialog1.FileName For Output As #1
        Print #1, txtText.Text
        Close #1
    Else
        Exit Sub
    End If
End If
Me.Caption = "D++ Compiler - [" & CommonDialog1.FileTitle & "]"
save = True
endit:
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuHelparr_Click()
frmHelp.Show
End Sub

Private Sub mnuProjectCompile_Click()
debugactive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize
mnuProjectStop.Enabled = True
txtDebug.Text = ""
AddDebug ">D++ Debug"
AddDebug ">"
Make
AddDebug ">D++ Application Finished"
mnuProjectStop.Enabled = False
End Sub

Private Sub mnuProjectRun_Click()
debugactive = True
txtDebug.Visible = True
mnuViewDebug.Checked = True
Form_Resize
mnuProjectStop.Enabled = True
txtDebug.Text = ""
AddDebug ">D++ Debug"
AddDebug ">"
Compile
AddDebug ">D++ Application Finished"
mnuProjectStop.Enabled = False
End Sub

Private Sub mnuTimeDate_Click()
txtText.SelText = Time & "/" & Date
End Sub

Private Sub mnuViewCalc_Click()
frmCalc.Show
End Sub

Private Sub mnuViewDebug_Click()
If mnuViewDebug.Checked = True Then
debugactive = False
txtDebug.Visible = False
Form_Resize
mnuViewDebug.Checked = False
Else
debugactive = True
txtDebug.Visible = True
Form_Resize
mnuViewDebug.Checked = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Help"
            mnuHelparr_Click
        Case "Run"
            mnuProjectRun_Click
        Case "Compile"
            mnuProjectCompile_Click
    End Select


End Sub
