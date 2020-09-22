VERSION 5.00
Begin VB.Form frmLinker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MakeDLL"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "No"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "Should this project be compiled into a DLL file?"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
'' VBLinker - A replacement Visual Basic linker ''
''          ©2004 DanSoft Australia.            ''

'' This is part of MakeDLL. To use this linker, ''
'' go into your VB folder (c:\program files\m...''
'' ...icrosoft visual studio\vb98, and rename   ''
'' "LINK.EXE" to "LINK1.EXE". Compile this prog ''
'' into that folder, and call it 'LINK.EXE'.    ''
'' Now use the MakeDLL Visual Basic AddIn to    ''
'' make a DLL in Visual Basic! Enjoy!           ''
''''''''''''''''''''''''''''''''''''''''''''''''''

'If you like this, PLEASE vote on Planet Source Code
' (look at the text files included in the ZIP file)

Option Explicit

'The API's required
Private Declare Function CreateProcess Lib "Kernel32" Alias _
                                            "CreateProcessA" ( _
    ByVal lpAppName As Long, _
    ByVal lpCmdLine As String, _
    ByVal lpProcAttr As Long, _
    ByVal lpThreadAttr As Long, _
    ByVal lpInheritedHandle As Long, _
    ByVal lpCreationFlags As Long, _
    ByVal lpEnv As Long, _
    ByVal lpCurDir As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInfo As PROCESS_INFORMATION _
) As Long
     
Private Declare Function WaitForSingleObject Lib "Kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
) As Long
    
Private Declare Function CloseHandle Lib "Kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Consts required
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

'Types reqiured
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Integer
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub cmdNo_Click()
ShellWait App.Path & "\link1.exe " & Command()
End
End Sub

Private Sub cmdYes_Click()
Dim strCmdLine As String
Dim strDefFile As String
Dim intTemp As Integer
'get the definition file name
strDefFile = DialogFile(Me.hWnd, 1, "Navigate to your project directory and open file", "", "Definition Files" & Chr(0) & "*.def", CurDir, "def")
If strDefFile = "" Then cmdNo_Click

'make new command line, so it compiles as a .dll file
strCmdLine = Command()
strCmdLine = Replace(strCmdLine, "/ENTRY:__vbaS", "/ENTRY:DLLMain")
strCmdLine = Replace(strCmdLine, "/BASE:0x400000", "/BASE:0x10000000")
strCmdLine = strCmdLine & " /DLL /DEF:""" & strDefFile & """"
cmdNo.Visible = False
cmdYes.Visible = False
lblDesc.Caption = "Compiling, please wait..."
Refresh
ShellWait App.Path & "\link1.exe " & strCmdLine
End
End Sub

Function ShellWait(CmdLine As String, Optional ByVal _
                        bShowApp As Boolean = False) As Boolean
    'Run a process, and wait for it to finish.
    
    
    Dim uProc As PROCESS_INFORMATION
    Dim uStart As STARTUPINFO
    Dim lRetVal As Long
    
    uStart.cb = Len(uStart)
    uStart.wShowWindow = Abs(bShowApp)
    uStart.dwFlags = 1
    
    lRetVal = CreateProcess(0&, CmdLine, 0&, 0&, 1&, _
                            NORMAL_PRIORITY_CLASS, 0&, 0&, _
                            uStart, uProc)
    lRetVal = WaitForSingleObject(uProc.hProcess, INFINITE)
    lRetVal = CloseHandle(uProc.hProcess)
    ShellWait = (lRetVal <> 0)
End Function

'Shows OpenFile dialog
'// szFilename = DialogFile(Me.hWnd, 1, "Open", "MyFileName.doc", "Documents" & Chr(0) & "*.doc" & Chr(0) & "All files" & Chr(0) & "*.*", App.Path, "doc")
Public Function DialogFile(hWnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String
    Dim x As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String
    OFN.lStructSize = Len(OFN)
    OFN.hWnd = hWnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

   If wMode = 1 Then
        OFN.Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        x = GetOpenFileName(OFN)
    Else
        OFN.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        x = GetSaveFileName(OFN)
    End If

    If x <> 0 Then
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
        DialogFile = szFile
    Else
        DialogFile = ""
    End If
End Function

