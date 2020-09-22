VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   4980
   ClientTop       =   2895
   ClientWidth     =   5745
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4440
      Top             =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "by caver dave"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPREHENSIVE NOTEPAD VERSION 3.5.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   20
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   3135
      Left            =   360
      Top             =   360
      Width           =   5055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   15
      Height          =   3375
      Left            =   240
      Top             =   240
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   2655
      Left            =   600
      Top             =   600
      Width           =   4575
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      BorderWidth     =   10
      Height          =   2895
      Left            =   480
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 2500
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 2500 Then
Timer1.Enabled = False
Unload Me
frmMain.Show
End If
End Sub
