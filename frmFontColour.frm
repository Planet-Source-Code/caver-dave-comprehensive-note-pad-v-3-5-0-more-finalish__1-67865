VERSION 5.00
Begin VB.Form frmFontColour 
   Caption         =   " FONT COLOUR PICKER"
   ClientHeight    =   5925
   ClientLeft      =   6000
   ClientTop       =   2190
   ClientWidth     =   7380
   Icon            =   "frmFontColour.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7380
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   6120
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      Picture         =   "frmFontColour.frx":0CCA
      ScaleHeight     =   4695
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label lblColorHex 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   2535
   End
End
Attribute VB_Name = "frmFontColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X As Integer
Dim Y As Integer

Dim Red As Integer, Green As Integer, Blue As Integer
Dim HRed As String, HGreen As String, HBlue As String
Dim colorHex As String


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Picture1.MousePointer = 99
Picture1.MouseIcon = LoadResPicture(102, vbResCursor)

lblColorHex = Hex(Picture1.Point(X, Y))
Picture2.BackColor = Picture1.Point(X, Y)

Blue = (Picture2.BackColor \ &H10000) Mod &H100
Green = (Picture2.BackColor \ &H100) Mod &H100
Red = Picture2.BackColor Mod &H100
Label1.Caption = "RGB(" & Red & ", " & Green & ", " & Blue & ")     "


HRed = Format(Hex(Red), "00")
HGreen = Format(Hex(Green), "00")
HBlue = Format(Hex(Blue), "00")
colorHex = HRed & HGreen & HBlue
Label2.Caption = "Hex: #" & colorHex

End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
'Unload Form1
End Sub
Private Sub Picture1_Click()
frmMain.RichTextBox1(frmMain.Text8.Text).SelColor = Picture2.BackColor
frmMain.Text3.ForeColor = Picture2.BackColor
If frmMain.Check1(9).Value = 1 Then
frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = 0 'Set the start pos of the selection
frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(frmMain.RichTextBox1(frmMain.Text8.Text)) 'Set length of the selection
frmMain.RichTextBox1(frmMain.Text8.Text).SelColor = Picture2.BackColor
ElseIf frmMain.Check1(9).Value = 0 Then
frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
frmMain.RichTextBox1(frmMain.Text8.Text).SelColor = Picture2.BackColor
End If
Unload Me
End Sub


