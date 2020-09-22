VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9d.ocx"
Begin VB.Form frmLink 
   BackColor       =   &H00000000&
   Caption         =   " LINK"
   ClientHeight    =   3750
   ClientLeft      =   4185
   ClientTop       =   4005
   ClientWidth     =   8730
   Icon            =   "FrmMainer.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   3750
   ScaleWidth      =   8730
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _cx             =   6376
      _cy             =   6165
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   3285
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      Visible         =   0   'False
      X1              =   8520
      X2              =   7680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   7680
      X2              =   8520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      Visible         =   0   'False
      X1              =   7680
      X2              =   7680
      Y1              =   3600
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      Visible         =   0   'False
      X1              =   7680
      X2              =   8520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   7680
      X2              =   7680
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   7680
      X2              =   8520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   8520
      X2              =   8520
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      Visible         =   0   'False
      X1              =   8520
      X2              =   8520
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IF YOU LIKE THIS PROGRAM CHECK OUT MY WEBSITE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label LblLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.sds-software-maker.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
ShockwaveFlash1.Movie = App.Path & "\sds_dg.swf"
ShockwaveFlash1.Loop = True
ShockwaveFlash1.Menu = False
End Sub

Private Sub LblLink_Click()
   Dim lWindow As Long
    Call ShellExecute(lWindow, "open", "http://www.sds-software-maker.com/index.html", vbNullString, vbNullString, 5)
    End Sub
Private Sub LblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblLink.FontBold = True
    LblLink.FontUnderline = True
    LblLink.ForeColor = vbRed
    Me.MousePointer = 99
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblLink.FontBold = False
    LblLink.FontUnderline = False
    LblLink.ForeColor = &H80FF&
    Me.MousePointer = 0
    Dim l As Integer
For l = 0 To 3
Line1(l).Visible = False
Next l
End Sub
Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index As Integer
Dim p As Integer
p = Index
For p = 0 To 3
Line2(p).Visible = True
Next p
Line1(0).Visible = False
Line1(1).Visible = False
Line1(2).Visible = False
Line1(3).Visible = False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index As Integer

 Label2.MousePointer = 99
 Label2.MouseIcon = LoadResPicture(101, vbResCursor)

Dim l As Integer
l = Index
For l = 0 To 3
Line1(l).Visible = True
Next l
End Sub
