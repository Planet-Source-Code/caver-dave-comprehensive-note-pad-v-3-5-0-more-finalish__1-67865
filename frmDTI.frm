VERSION 5.00
Begin VB.Form frmDTI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ShortCut 2 Desktop"
   ClientHeight    =   1950
   ClientLeft      =   5175
   ClientTop       =   1560
   ClientWidth     =   8730
   Icon            =   "frmDTI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":0CCA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":1994
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":265E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "CREATE FILE ASSOCIATION"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CREATE DESKTOP SHORTCUT"
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtFileAssociation 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   4305
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?????"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CREATE FILE ASSOCIATION"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtShortCutName 
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   105
      Width           =   5895
   End
   Begin VB.TextBox txtIconLocation 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   825
      Width           =   5895
   End
   Begin VB.TextBox txtTargetPath 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   465
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE DESKTOP SHORTCUT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Default File Extension"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Default ShortCut Name"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Default Icon Location"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Default Target Path"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmDTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wShell As New IWshShell_Class
Dim wShortcut As IWshShortcut_Class
Private Sub Check1_Click()
If Check1.Value = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Call CreateIcon
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtTargetPath.Text = App.Path & "\swapr.exe"
txtIconLocation.Text = App.Path & "\swapr.exe"
txtShortCutName.Text = "Comprehensive Note Pad v3.5.0.lnk"
'txtShortCutName.Text = "Comprehensive Note Pad v2.75.lnk"
End Sub

Private Sub CreateIcon()
   Set wShortcut = wShell.CreateShortcut(wShell.SpecialFolders.Item(0) & "\" & txtShortCutName.Text)
    wShortcut.TargetPath = txtTargetPath.Text
    wShortcut.IconLocation = txtIconLocation.Text
    wShortcut.Save
End Sub

