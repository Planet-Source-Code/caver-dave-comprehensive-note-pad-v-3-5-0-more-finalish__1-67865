VERSION 5.00
Begin VB.Form frmWordLister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " WORD LIST MAKER"
   ClientHeight    =   8040
   ClientLeft      =   5670
   ClientTop       =   1785
   ClientWidth     =   4425
   Icon            =   "frmWordLister.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   4425
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   6945
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S&AVE"
      Height          =   615
      Left            =   2520
      MouseIcon       =   "frmWordLister.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmWordLister.frx":0E1C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   3480
      MouseIcon       =   "frmWordLister.frx":13A6
      MousePointer    =   99  'Custom
      Picture         =   "frmWordLister.frx":14F8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   6690
      ItemData        =   "frmWordLister.frx":1A82
      Left            =   120
      List            =   "frmWordLister.frx":1A84
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox txtWLBf 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "WORD COUNT"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "TYPE LIST NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
End
Attribute VB_Name = "frmWordLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const myerrfilepath = 380
Private Sub Command1_Click()
Dim X$, bah$
Dim i As Integer
'  ReDim astrWords(2) As String
'    Dim intLCV As Long
Call sRtb2Txt
'    Call subGetWords(txtWLBf.Text, astrWords())
'    List1.Clear
'    For intLCV = 0 To UBound(astrWords)
'        List1.AddItem astrWords(intLCV)
'    Next intLCV
List1.Clear
txtWLBf = Trim(txtWLBf)


For i = 1 To Len(txtWLBf)
    X$ = Mid(txtWLBf, i, 1)


    If X$ = " " Then
        List1.AddItem bah$
        bah$ = ""
        GoTo yo
    End If

    bah$ = bah$ & X$
yo:
Next i

List1.AddItem bah$

Call Remove_Duplicate(List1)

Text1.Text = List1.ListCount
End Sub
Private Sub sRtb2Txt()
frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = 0 'Set the start pos of the selection
frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(frmMain.RichTextBox1(frmMain.Text8.Text))

txtWLBf.Text = frmMain.RichTextBox1(frmMain.Text8.Text).Text
End Sub
Sub Remove_Duplicate(ListBx As ListBox)

    Dim X

    Do
        ListBx.Text = ListBx.List(X)
        If Not ListBx.ListIndex = X Then ListBx.RemoveItem X
        If ListBx.ListIndex = X Then X = X + 1
    Loop Until X > ListBx.ListCount - 1

    ListBx.ListIndex = 0
    ListBx.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
'save file
Dim v As Integer
Dim Filehandle As Integer

Filehandle = FreeFile


Open App.Path & "\" & Text2.Text & ".txt" For Output As #Filehandle   ' Open file for output.
For v = 0 To Text1.Text - 1
Print #Filehandle, List1.List(v)
Next v
Close #Filehandle
'*********************************************************************************
'*** Print# -operations with list or combo boxes                               ***
'*** input - opens the file to the program                                     ***
'*** output - saves the current data and over writes the open file i.e save as ***
'*** append - adds extra data to the file                                      ***
'*********************************************************************************
End Sub


Private Sub Form_Load()
Dim Msg As String
On Error GoTo fubar
Call Command1_Click
fubar:
If (Err.Number = myerrfilepath) Then
    Msg = "NO TEXT OR WORDS TO LIST"
    If MsgBox(Msg) = vbOK Then
      frmMain.SetFocus
    End If
  End If
  Exit Sub

End Sub
