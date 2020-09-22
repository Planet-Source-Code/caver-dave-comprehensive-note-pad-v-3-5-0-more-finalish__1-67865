VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " READ ME"
   ClientHeight    =   6315
   ClientLeft      =   5505
   ClientTop       =   2655
   ClientWidth     =   5970
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5970
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   5535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":0CCA
      Top             =   120
      Width           =   5775
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
      Left            =   4680
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "A file wiped with this program can no longer be recovered after being wiped." & vbCrLf & _
        "Are you sure you wish to permanently" & vbCrLf & vbCrLf & "remove the files that you are going to select?" & vbCrLf & vbCrLf & _
        "DO NOT WIPE ANY FILES THAT YOU ARE NOT CERTAIN ABOUT, or FILES ASSOCIATED WITH WINDOWS OR POSSIBLY FILES THAT ARE NEEDED BY THE SYSTEM" & vbCrLf & vbCrLf & _
        "CAVI CAUTUM::- USER BEWARE YOU COULD CORRUPT YOUR SYSTEM!"
'Populus caveo
' cavi cautum
End Sub

