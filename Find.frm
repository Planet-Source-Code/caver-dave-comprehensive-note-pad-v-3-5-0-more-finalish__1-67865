VERSION 5.00
Begin VB.Form Find1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Find"
   ClientHeight    =   1665
   ClientLeft      =   7545
   ClientTop       =   2190
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111.765
   ScaleMode       =   0  'User
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Picture         =   "Find.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox cboRep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Replace..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Picture         =   "Find.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Find"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Find.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Picture         =   "Find.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "Find.frx":1EF2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "R&eplace All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "F&ind"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "R&eplace All"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Picture         =   "Find.frx":247C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Replace With:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Find1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Replace_All
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ToolTipText = "REPLACE ALL ..."
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command2_Click()
    If frmMain.RichTextBox1(frmMain.Text8.Text).SelText = "" Then
        frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = 0
    End If
        intBegSearch = frmMain.RichTextBox1(frmMain.Text8.Text).SelStart + 2

    If Command2.Caption = "&Find" Then
        Find
        'Command2.Caption = "&Find Next"
    Else
        FindNext
    End If

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.ToolTipText = "FIND..."
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.ToolTipText = "REPLACE ..."
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.ToolTipText = "CANCEL FIND"
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command5_Click()
    vReplace
End Sub

Private Sub Command4_Click()
    Command2.Enabled = True
    Command3.Visible = True
    Command3.Enabled = True
    Command1.Visible = False
    Command1.Enabled = False
    Command5.Visible = False
    Command5.Enabled = False
    cboRep.Visible = False
    Label2.Visible = False
    frmMain.SetFocus

        If cboRep.Text = "" Then
        Else
            cboRep.AddItem cboRep.Text
        End If
    Me.Hide
End Sub

Private Sub Command3_Click()
    Command3.Visible = False
    Command3.Enabled = False
    Command1.Visible = True
    Command1.Enabled = True
    Command5.Visible = True
    Command5.Enabled = True
    Command2.Enabled = False
    cboRep.Visible = True
    Label2.Visible = True
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.ToolTipText = "REPLACE ..."
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.ToolTipText = "EXIT THIS DIALOGUE, PLEASE USE THE DOOR!"
Command6.MousePointer = 99
Command6.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Form_Load()
'   SET ALWAYS ON TOP(TRUE)
    MakeAlwaysOnTop Me, True
    
End Sub


Private Sub Text1_Change()
Command2.Caption = "&Find"
If Text1.Text = "" Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If
End Sub


