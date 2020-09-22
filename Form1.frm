VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   " FILE WIPER"
   ClientHeight    =   2745
   ClientLeft      =   5670
   ClientTop       =   3780
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "READ ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Height          =   1935
      Left            =   2640
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "CLICK THIS BUTTON TO SELECT A FILE TO WIPE AND THEN WIPE THE FILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------'
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constants for "CreateFile" function "dwDesiredAccess" parameter  ]'
Private Const GENERIC_READ  As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000

'[  constants for "CreateFile" function "dwShareMode" parameter  ]'
Private Const FILE_SHARE_READ  As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

'[  constants for "CreateFile" function "dwCreateDisposition" parameter  ]'
Private Const CREATE_ALWAYS     As Long = 2
Private Const CREATE_NEW        As Long = 1
Private Const OPEN_ALWAYS       As Long = 4
Private Const OPEN_EXISTING     As Long = 3
Private Const TRUNCATE_EXISTING As Long = 5

'[  constants for "CreateFile" function "dwFlagsAndAttributes" parameter  ]'
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const FILE_FLAG_NO_BUFFERING    As Long = &H20000000
Private Const FILE_FLAG_OVERLAPPED      As Long = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS As Long = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS   As Long = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH   As Long = &H80000000

'[  constants for "SetFilePointer" function "dwMoveMethod" parameter  ]'
Private Const FILE_BEGIN   As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END     As Long = 2


'------------------------------------------------------------------------------'
'  FileSystemObject Constants                                                  '
'------------------------------------------------------------------------------'


'[  constants for "GetSpecialFolder" function "folderspec" parameter  ]'
Private Const GSF_WINDOWSFOLDER As Long = 0
Private Const GSF_SYSTEMFOLDER As Long = 1
Private Const GSF_TEMPORARYFOLDER As Long = 2

'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'

Private Declare Function CloseHandle _
        Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreateFile _
        Lib "kernel32.dll" Alias "CreateFileA" ( _
          ByVal lpFileName As String, _
          ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, _
          ByVal lpSecurityAttributes As Long, _
          ByVal dwCreationDisposition As Long, _
          ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long _
        ) As Long
        
        Private Declare Function SetFilePointer _
        Lib "kernel32.dll" ( _
          ByVal iFileHandler As Long, _
          ByVal lDistanceToMove As Long, _
          ByRef lpDistanceToMoveHigh As Long, _
          ByVal dwMoveMethod As Long _
        ) As Long
        
        Private Declare Function WriteFile _
        Lib "kernel32.dll" ( _
          ByVal iFileHandler As Long, _
          ByRef lpBuffer As Any, _
          ByVal nNumberOfBytesToWrite As Long, _
          ByRef lpNumberOfBytesWritten As Long, _
          ByVal lpOverlapped As Long _
        ) As Long
 
Private Sub Command1_Click()
Dim Msg As String

  On Error Resume Next
  dlgBrowse.ShowOpen

  Msg = "A file can no longer be recovered after being wiped." & vbCrLf & _
        "Are you sure you wish to permanently remove the following file?" & _
        vbCrLf & vbCrLf & _
        dlgBrowse.FileName

  If Err.Number = 0 Then
    If MsgBox(Msg, vbQuestion + vbYesNo) = vbYes Then
      If WipeFile(dlgBrowse.FileName, True) Then
        MsgBox "The selected file has been successfully wiped."
      Else
        MsgBox "The selected file cannot be wiped."
      End If
    End If
  End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Form_Load()
  dlgBrowse.InitDir = App.Path 'GetPrimaryDrive
  dlgBrowse.Filter = "(All files)|*.*|"
  dlgBrowse.Flags = cdlOFNHideReadOnly + cdlOFNLongNames + _
                    cdlOFNPathMustExist + cdlOFNFileMustExist
  Me.Show
End Sub
Public Function GetPrimaryDrive() As String
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  GetPrimaryDrive = Left$(FSO.GetSpecialFolder(GSF_WINDOWSFOLDER), 2) & "\"
End Function
Public Function WipeFile(ByVal FileName As String, ByVal Rename As Boolean) _
                         As Boolean

  Dim FileName2    As String
  Dim FileSize     As Long
  Dim hFile        As Long
  Dim Ctr          As Long
  Dim Pattern()    As Byte
  Dim BytesWritten As Long

  On Error Resume Next

  '[  set file attribute to normal and retrieve file size  ]'
  SetAttr FileName, vbNormal
  FileSize = FileLen(FileName)

  '[  rename file and move to windows temporary folder  ]'
  If Rename Then
    FileName2 = GenerateTempFile
    Kill FileName2
    Name FileName As FileName2
  Else
    FileName2 = FileName
  End If

  On Error GoTo FileError

  '[  open file with disk caching disabled  ]'
  hFile = CreateFile(FileName2, GENERIC_WRITE, 0, 0, OPEN_EXISTING, _
                     FILE_FLAG_WRITE_THROUGH + FILE_FLAG_DELETE_ON_CLOSE + _
                     FILE_FLAG_SEQUENTIAL_SCAN, 0)

  '[  if file opened successfully, then wipe file  ]'
  If hFile <> -1 Then

    ReDim Pattern(1 To FileSize, 1 To 3)

    '[  assign bit patterns  ]'
    For Ctr = 1 To FileSize
      Pattern(Ctr, 1) = &H55  '[  bit pattern 01010101  ]'
      Pattern(Ctr, 2) = &HAA  '[  bit pattern 10101010  ]'
      Pattern(Ctr, 3) = &H0   '[  bit pattern 00000000  ]'
    Next Ctr

    '[  write bit patterns to file  ]'
    For Ctr = 1 To 3
      SetFilePointer hFile, 0, 0, FILE_BEGIN
      WriteFile hFile, Pattern(1, Ctr), FileSize, BytesWritten, 0
    Next Ctr

    '[  close and delete file  ]'
    CloseHandle hFile
    WipeFile = True

    Exit Function

  End If

FileError:
  WipeFile = False

End Function
Public Function GenerateTempFile() As String

  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  '[  set current drive to that of the windows folder  ]'
  ChDir Left$(FSO.GetSpecialFolder(GSF_WINDOWSFOLDER), 2) & "\"

  '[  get path of temporary folder  ]'
  GenerateTempFile = FSO.GetSpecialFolder(GSF_TEMPORARYFOLDER) & "\" & _
                     UCase$(FSO.GetTempName)

End Function

