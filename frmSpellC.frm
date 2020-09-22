VERSION 5.00
Begin VB.Form frmSpellC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " SPELL CHECKER"
   ClientHeight    =   4530
   ClientLeft      =   9525
   ClientTop       =   2205
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSpellC.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   3030
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   420
      Width           =   2835
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   1200
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   7620
      TabIndex        =   8
      Top             =   2640
      Width           =   1635
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   7560
      TabIndex        =   7
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      Picture         =   "frmSpellC.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox lstSug 
      Height          =   1035
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2835
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "REPLACE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmSpellC.frx":09D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   1680
      Picture         =   "frmSpellC.frx":0F5E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "IGNORE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "frmSpellC.frx":14E8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   120
      X2              =   2880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   1560
      X2              =   1560
      Y1              =   2760
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      X1              =   1560
      X2              =   1560
      Y1              =   2760
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2880
      X2              =   2880
      Y1              =   2760
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   120
      X2              =   2880
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   2880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   2880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   2880
      X2              =   2880
      Y1              =   4440
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   2880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   2760
   End
   Begin VB.Label lblSug 
      Caption         =   "Suggestions:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label lblNF 
      Caption         =   "Wrong word:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      X1              =   120
      X2              =   2880
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmSpellC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lasti As Long

Private Sub cmdAdd_Click()
Open App.Path & "\dict.cnp" For Append As #1
Print #1, Text1.Text
Close #1
List2.AddItem Text1.Text
DoEvents: DoEvents: DoEvents
Timer1.Enabled = True
cmdIgnore.Enabled = False
cmdReplace.Enabled = False
cmdAdd.Enabled = False
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.ToolTipText = "ADD TO WORD LIST"
cmdAdd.MousePointer = 99
cmdAdd.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub cmdIgnore_Click()
Timer1.Enabled = True
cmdIgnore.Enabled = False
cmdReplace.Enabled = False
cmdAdd.Enabled = False
End Sub

Private Sub cmdIgnore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdIgnore.ToolTipText = "IGNORE WORD"
cmdIgnore.MousePointer = 99
cmdIgnore.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdQuit.ToolTipText = "EXIT THIS DIALOGUE, PLEASE USE THE DOOR!"
cmdQuit.MousePointer = 99
cmdQuit.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub cmdReplace_Click()
If lstSug.Text <> "" Then
frmMain.RichTextBox1(frmMain.Text8.Text).Text = Replace(frmMain.RichTextBox1(frmMain.Text8.Text).Text, Text1.Text, lstSug.Text, , , vbTextCompare)
Else
frmMain.RichTextBox1(frmMain.Text8.Text).Text = Replace(frmMain.RichTextBox1(frmMain.Text8.Text).Text, Text1.Text, lstSug.List(0), , , vbTextCompare)
End If
cmdIgnore.Enabled = False
cmdReplace.Enabled = False
cmdAdd.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub cmdReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdReplace.ToolTipText = "REPLACE ..."
cmdReplace.MousePointer = 99
cmdReplace.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Form_Load()
If Dir(App.Path & "\dict.cnp") = "" Then
MsgBox "dic.txt missing!"
Unload Me
End If
Open App.Path & "\dict.cnp" For Input As #1
Do While Not EOF(1)
Line Input #1, l
List2.AddItem l
Loop
Close #1

Dim txt As String
txt = frmMain.RichTextBox1(frmMain.Text8.Text).Text
For i = 0 To Len(txt)
On Error GoTo l1
List1.AddItem Split(txt, " ", , vbTextCompare)(i)
Next i
l1:
lasti = 0
If List2.ListCount = List1.ListCount Then
MsgBox "Complete."
Timer1.Enabled = False
Unload Me
Exit Sub
End If
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Dim w As String
If lasti = List1.ListCount Then
MsgBox "Complete."
Timer1.Enabled = False
Unload Me
Exit Sub
End If

For i = lasti To List1.ListCount - 1
    w = List1.List(i)
        For ii = 0 To List2.ListCount - 1
            If w = List2.List(ii) Then
            GoTo l1
            Exit For
            End If
        Next ii
        Timer1.Enabled = False
      Text1.Text = w
      cmdIgnore.Enabled = True
      cmdAdd.Enabled = True
      
      lasti = lasti + 1
      lstSug.Clear

      For j = 0 To List2.ListCount - 1
      If Len(w) > 2 Then
        If Left(w, Len(w) - 1) = Left(List2.List(j), Len(w) - 1) Then
            lstSug.AddItem List2.List(j)
            GoTo l2
         ElseIf Left(w, Len(w) - 2) = Left(List2.List(j), Len(w) - 2) Then
               lstSug.AddItem List2.List(j)
            GoTo l2
         ElseIf Right(w, Len(w) - 2) = Right(List2.List(j), Len(w) - 2) Then
               lstSug.AddItem List2.List(j)
            GoTo l2
            ElseIf Right(w, Len(w) - 2) = Right(List2.List(j), Len(w) - 2) Then
               lstSug.AddItem List2.List(j)
            GoTo l2
        End If
      End If
l2:
      Next j
      If lstSug.ListCount > 0 Then
      cmdReplace.Enabled = True
      End If
Exit Sub
l1:
  lasti = lasti + 1
Next i
'frmMain.RichTextBox1(frmMain.Text8.Text)
End Sub

