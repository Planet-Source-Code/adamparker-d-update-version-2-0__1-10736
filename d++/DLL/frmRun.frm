VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D++ APP"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock winsck 
      Left            =   1560
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Text            =   "Input Responces"
      Top             =   3840
      Width           =   7215
   End
   Begin VB.TextBox txtIn 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6480
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   2640
   End
   Begin VB.Label input1 
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean, flag1 As Boolean
Dim val1, val2, val3, val4, val5, val6, val7, chartext, chartext1
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Private Sub Form_Activate()
txtType.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo Errorh
    
    Me.Show
    
    Open App.Path + "\" + App.EXEName + ".EXE" For Binary As #1
    
    FileSize = LOF(1)
    FileData$ = Space$(LOF(1))
    
    Get #1, , FileData$

    For i = 1 To FileSize
        If Mid(FileData$, i, 4) = "DPP:" Then
            i = i + 4
            FileChunk$ = String(1000, 0)
            Get #1, i, FileChunk$
            txtIn.Text = FileChunk$
            Linkit
            Exit Sub
        End If
    Next i
    
    Close #1
    
Errorh:
MsgBox "Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Error"
End Sub

Sub Typeit(Typed As String)
End Sub


Sub Linkit()
For i = 1 To Len(txtIn.Text)

    If LCase(Mid(txtIn.Text, i, 11)) = LCase("screenout " & Chr(34)) Then
        i = i + 11
        d = i + 256
        Do Until Mid(txtIn.Text, i, 2) = Chr(34) & ";"
        If i = d Then
            MsgBox "Expected ';' at " & i & "; Found end of program.", vbCritical, "Error"
            End
        End If
        a Mid(txtIn.Text, i, 1)
        i = i + 1
        Loop
        
    ElseIf LCase(Mid(txtIn.Text, i, 11)) = LCase("screenput """) Then
        i = i + 11
        d = i + 256
        chartext = ""
        Do Until Mid(txtIn.Text, i, 2) = """;"
        If i = d Then
            MsgBox "Expected ';' at " & i & "; Found end of program.", vbCritical, "Error"
            End
        End If
        chartext = chartext & Mid(txtIn.Text, i, 1)
        i = i + 1
        Loop
        txtText.Text = txtText.Text & chartext
        
    ElseIf LCase(Mid(txtIn.Text, i, 10)) = LCase("screenout ") Then
        i = i + 10
        d = i + 256
        Do Until Mid(txtIn.Text, i, 1) = ";"
        If i = d Then
            MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
            End
        End If
        val3 = val3 & Mid(txtIn.Text, i, 1)
        i = i + 1
        Loop
        If val3 = val1 Then
            a val2
        Else
            MsgBox "Syntax Error: variable not defined.", vbCritical, "Syntax Error"
            End
        End If

    ElseIf LCase(Mid(txtIn.Text, i, 10)) = LCase("screenput ") Then
        i = i + 10
        d = i + 256
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            chartext1 = chartext1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        If chartext1 = val1 Then
            txtText.Text = txtText.Text & val2
        Else
            MsgBox "Syntax Error: variable not defined.", vbCritical, "Syntax Error"
            End
        End If
        
    ElseIf LCase(Mid(txtIn.Text, i, 9)) = LCase("screenin ") Then
        val4 = ""
        i = i + 9
        d = i + 256
        Do Until Mid(txtIn.Text, i, 3) = ", """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Error"
                End
            End If
            val1 = val1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        i = i + 3
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            val4 = val4 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        val2 = InputBox(val4, "Input Value", "Value")
        
        
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = LCase("title """) Then
        i = i + 7
        d = i + 256
        Me.Caption = ""
    
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            Me.Caption = Me.Caption & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        App.Title = Me.Caption
            
    ElseIf LCase(Mid(txtIn.Text, i, 4)) = LCase("time") Then
        i = i + 4
        d = i + 256
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            i = i + 1
        Loop
            txtText.Text = txtText.Text & Time
            

    ElseIf LCase(Mid(txtIn.Text, i, 8)) = LCase("delete """) Then
        i = i + 8
        d = i + 256
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            val5 = val5 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        If FileExist(val5) = False Then
            MsgBox "Run Time Error: File not found", vbCritical, "Run Time Error"
            End
        Else
            Kill val5
        End If
        
    ElseIf LCase(Mid(txtIn.Text, i, 7) = "delete ") Then
        i = i + 7
        d = i + 256
        val6 = ""
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            val6 = val6 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        If val6 = val1 Then
            If FileExist(val2) = False Then
                MsgBox "Error!  File not found!", vbCritical, "File Not Found"
            Else
                Kill val2
            End If
        Else
            MsgBox "Syntax Error: Variable not defined: " & val6, vbCritical, "Syntax Error"
            End
        End If
        
    ElseIf Mid(txtIn.Text, i, 1) = "<" Then
        i = i + 1
        Do Until Mid(txtIn.Text, i, 1) = ">"
            i = i + 1
        Loop
        i = i + 1
        
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("box """) Then
        i = i + 5
        d = i + 256
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box1 = box1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        i = i + 4
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box2 = box2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        MsgBox box1, vbExclamation, box2
        
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("pause") Then
        i = i + 5
        d = i + 256
        val7 = ""
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            val7 = val7 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        Pause val7
        
    ElseIf LCase(Mid(txtIn.Text, i, 6)) = LCase("clear;") Then
        txtText.Text = ""
        
    ElseIf LCase(Mid(txtIn.Text, i, 8)) = LCase("pause05;") Then
        Pause 0.5

    ElseIf LCase(Mid(txtIn.Text, i, 7)) = LCase("pause1;") Then
        Pause 1
        
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = LCase("pause2;") Then
        Pause 2
        
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = LCase("pause3;") Then
        Pause 3
        
    ElseIf LCase(Mid(txtIn.Text, i, 4)) = LCase("end;") Then
        End
        
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = LCase("screen;") Then
        txtText.Text = txtText.Text & vbCrLf
                
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("wav """) Then
        i = i + 5
        d = i + 256
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            val5 = val5 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        If FileExist(val5) = False Then
            MsgBox "Run Time Error: File not found", vbCritical, "Run Time Error"
            End
        Else
            Playwav (val5)
        End If
                
    ElseIf LCase(Mid(txtIn.Text, i, 6)) = LCase("ontop;") Then
        FormOnTop Me
    
    ElseIf LCase(Mid(txtIn.Text, i, 9)) = LCase("notontop;") Then
        FormNotOnTop Me
        
                
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("add """) Then
        i = i + 5
        d = i + 256
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box1 = box1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        i = i + 4
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box2 = box2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(box1) + CDbl(box2)
        
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("sub """) Then
        i = i + 5
        d = i + 256
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box1 = box1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        i = i + 4
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box2 = box2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(box1) - CDbl(box2)
        
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("mul """) Then
        i = i + 5
        d = i + 256
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box1 = box1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        i = i + 4
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box2 = box2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(box1) * CDbl(box2)
        
        
    Else
        If Mid(txtIn.Text, i, 1) <> "" And Mid(txtIn.Text, i, 1) = vbCrLf And Mid(txtIn.Text, i, 1) <> " " And Mid(txtIn.Text, i, 1) <> "    " Then
            MsgBox "Syntax Error: Invalid syntax at " & i & ". (" & Mid(txtIn.Text, i, 1) & ")", vbCritical, "Syntax Error"
        End If
    End If
    
Next i
End Sub

Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub a(TextToPut)
On Error Resume Next
'txtText.SelText = vbCrLf
For sd = 1 To Len(TextToPut)
txtText.SelStart = Len(txtText)
txtText.SelText = Mid(TextToPut, sd, 1)
DoEvents
Pause 0.01
Next sd
txtText.SelStart = Len(txtText)
End Sub

Function FileExist(ByVal FileName As String) As Boolean
    Dim fileFile As Integer
    fileFile = FreeFile
    On Error Resume Next
    Open FileName For Input As fileFile
    If Err Then
        FileExist = False
    Else
        Close fileFile
        FileExist = True
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
    txtText.SelStart = Len(txtIn.Text)
End Sub

Private Sub txtText_GotFocus()
txtType.SetFocus
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Call Typeit(txtType.Text)
txtType = ""
End If
End Sub
