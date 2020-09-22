Attribute VB_Name = "Module1"
' --- Functions for GetSystemDirectory ---
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' --- Functions for GetDesktopDirectory ---
' Special folder values for SHGetSpecialFolderLocation and
' SHGetSpecialFolderPath (Shell32.dll v4.71)
' Retrieves the path of a special folder.
' The docs say it returns NOERROR if successful, or an
' OLE-defined error result otherwise, *but* with both
' Shell32.dll v4.71 and v4.72 I have only seen it return 1
' if successful, or 0 otherwise.
Private Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" (ByVal hwndOwner As Long, ByVal pszPath As String, ByVal nFolder As Long, ByVal fCreate As Boolean) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Const NOERROR = 0
Private Const MAX_PATH = 260
Private Const CSIDL_DESKTOPDIRECTORY = &H10

Private DLLFILE As String

Sub SetDllLocation()
DLLFILE = GetSystemDirectory & "\DPPAPP.dll"
End Sub

Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Function GetDesktopDirectory(hWnd As Long)
Dim pidl As Long, sPath As String * MAX_PATH, nFolder As Long

nFolder = CSIDL_DESKTOPDIRECTORY
' If the version of Shell32.dll is < v4.71 then
' SHGetSpecialFolderPath won't be exported and we'll get VB error '453.
   On Error GoTo NotExported

   ' Since we're not sure what the call's return value is, we'll
   ' just check where the first Null char is in the path below.
 
  Call SHGetSpecialFolderPath(hWnd, sPath, nFolder, 0)
   ' Return the path (if any)
   If InStr(sPath, vbNullChar) > 1 Then
      GetDesktopDirectory = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      Exit Function
   End If

NotExported:
   ' Get the pointer to the folder's item ID list from
   ' it's specified folder ID, returns 0 on success
   If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) _
     = NOERROR Then
      If pidl Then
         ' Get the path from the pointer to the item id list,
         ' returns True on success.
         If SHGetPathFromIDList(pidl, sPath) Then
            ' Return the path
           GetDesktopDirectory = Left$(sPath, InStr(sPath, _
                vbNullChar) - 1)
         End If
         ' Free the memory the shell allocated for the pidl
         Call CoTaskMemFree(pidl)
      End If
   End If
End Function

Public Function GetSystemDirectory() As String
Dim strBuffer As String, lngReturn As String
strBuffer = Space(255)
lngReturn = apiGetSystemDirectory(strBuffer, Len(strBuffer))
GetSystemDirectory = Left(strBuffer, lngReturn)
End Function

Sub AddDebug(TextToAdd As String)
frmMain.txtDebug.Text = frmMain.txtDebug.Text & TextToAdd & vbCrLf
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

Function ReadFile(ByVal sFileName As String) As String
    Dim fhFile As Integer
    fhFile = FreeFile
    Open sFileName For Binary As #fhFile
    ReadFile = Input$(LOF(fhFile), fhFile)
    Close #fhFile
End Function

Sub Make()
On Error GoTo Errorh

errorbug = 0

If frmMain.txtText.Text = "" Then
    AddDebug ">Compile Error: Not valid program: Null"
    AddDebug ">Application Terminated."
Else
    
    frmMain.CommonDialog1.Filter = "EXE Files (*.exe)|*.exe|All Files (*.*)|*.*"
    frmMain.CommonDialog1.DialogTitle = "Save D++ File"
    frmMain.CommonDialog1.CancelError = True
    frmMain.CommonDialog1.ShowSave

    APPFILE = frmMain.CommonDialog1.FileName
    
    If FileExist(APPFILE) Then
        overwrite = MsgBox("Overwrite Existing .EXE File?", 276, "File Found!")
        If overwrite = 6 Then
            Kill APPFILE
            AddDebug ">EXE overwritten."
        Else
            AddDebug ">Canceled overwrite."
            AddDebug ">Application Terminated."
            Exit Sub
        End If
    End If
    
    FileCopy DLLFILE, APPFILE
    PUTINF = "DPP:" + frmMain.txtText.Text
    File1$ = APPFILE
    File2$ = DLLFILE
    
    AddDebug ">Compiling Project..."
    
    Open File1$ For Output As #1        'Open Application
    Open File2$ For Binary As #2        'Open DLL File
    Do While Not EOF(2)
        filedata = Input$(2000, #2)
        msg = filedata
        msg2 = msg2 + msg
        Print #1, msg2;
        msg2 = ""
        If Len(msg) > 2000 Then
            msg = ""
        End If
    Loop
    AddDebug ">Writting Data..."
    Print #1, PUTINF                    'Application
    Close #2                            'Close DLL File
    Close #1                            'Close Application
    
    AddDebug ">" & APPFILE & " was complied successfully."
    
    Pause 0.01
    Shell APPFILE, vbNormalFocus
    
End If
Errorh:
    If Err.Number = 0 Then Exit Sub
    errorbug = errorbug + 1
    AddDebug ">Error #" & Err.Number & ": " & Err.Description
    AddDebug ">Application Terminated."
    Exit Sub
End Sub

Sub Compile()
On Error GoTo Errorh

errorbug = 0

If frmMain.txtText.Text = "" Then
    AddDebug ">Error: Not valid program: Null"
    AddDebug ">Application Terminated."
Else

    APPFILE = GetDesktopDirectory(frmMain.hWnd) & "\D++APP1.EXE"
    
    If FileExist(APPFILE) Then
        overwrite = MsgBox("Overwrite Existing .EXE File?", 276, "File Found!")
        If overwrite = 6 Then
            Kill APPFILE
            AddDebug ">EXE overwritten."
        Else
            AddDebug ">Canceled overwrite."
            AddDebug ">Application Terminated."
            Exit Sub
        End If
    End If
    
    FileCopy DLLFILE, APPFILE
    Putfil = "DPP:" + frmMain.txtText.Text
    File1$ = APPFILE
    File2$ = DLLFILE
    
    AddDebug ">Compiling Project..."
    
    Open File1$ For Output As #1        'Open Application
    Open File2$ For Binary As #2        'Open DLL File
    Do While Not EOF(2)
        filedata = Input$(2000, #2)
        msg = filedata
        msg2 = msg2 + msg
        Print #1, msg2;
        msg2 = ""
        If Len(msg) > 2000 Then
            msg = ""
        End If
    Loop
    AddDebug ">Writing Data..."
    Print #1, Putfil                    'Application
    Close #2                            'Close DLL File
    Close #1                            'Close Application
    
    AddDebug ">" & APPFILE & " was complied successfully."
    
    Pause 0.01
    Shell APPFILE, vbNormalFocus
    
End If
Errorh:
    If Err.Number = 0 Then Exit Sub
    errorbug = errorbug + 1
    AddDebug ">Error #" & Err.Number & ": " & Err.Description
    Pause 0.03
    AddDebug ">Application Terminated."
    Exit Sub
End Sub

