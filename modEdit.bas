Attribute VB_Name = "modEdit"
Option Explicit
    Public strSearchFor As String
    Public strFirstW As String


    Public intFoundPos As Integer
    Public intFoundPosUR As Integer
    Public intBegSearch As Integer
    Public strSearch As String
'*****************************************************************************
'AlwaysOnTop                                                               '**
'*****************************************************************************
    Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long          '**
    Const HWND_TOPMOST = -1                                                    '**
    Const HWND_NOTOPMOST = -2                                                  '**
    Const SWP_NOMOVE As Long = &H2                                             '**
    Const SWP_NOSIZE As Long = &H1                                             '**
Public Sub MakeAlwaysOnTop(TheForm As Form, SetOnTop As Boolean)           '**
    Dim lflag                                                              '**
    If SetOnTop Then                                                       '**
        lflag = HWND_TOPMOST                                               '**
    Else                                                                   '**
        lflag = HWND_NOTOPMOST                                             '**
    End If                                                                 '**
    SetWindowPos TheForm.hWnd, lflag, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE '**
End Sub
Public Sub sFind()
    Dim X As Integer
      Find1.Caption = " FIND"
            If frmMain.RichTextBox1(frmMain.Text8.Text).Text = "" Then Exit Sub
                Find1.Command2.Caption = "&Find"
                SelectWord
                Find1.Text1.AddItem strFirstW
                X = Find1.Text1.ListCount
                Find1.Text1.Text = Find1.Text1.List(X - 1)
                strFirstW = ""
                'Find1.Show
            If frmMain.RichTextBox1(frmMain.Text8.Text).SelText <> "" Then
                intBegSearch = frmMain.RichTextBox1(frmMain.Text8.Text).SelStart + 2
                strSearchFor = frmMain.RichTextBox1(frmMain.Text8.Text).SelText
                intFoundPos = InStr(1, frmMain.RichTextBox1(frmMain.Text8.Text).SelText, strSearchFor, 1)   'search for text
                Find1.Text1.Text = frmMain.RichTextBox1(frmMain.Text8.Text).SelText
                Find1.Show
            Else
                ctext Find1
                Find1.Show
            End If
            End Sub
Public Sub sReplace()
Dim X As Integer
        Find1.Caption = " FIND AND REPLACE"
                SelectWord
                Find1.Text1.AddItem strFirstW
                X = Find1.Text1.ListCount
                Find1.Text1.Text = Find1.Text1.List(X - 1)
                strFirstW = ""
                Find1.Command2.Caption = "&Find"
                Find1.Command3.Visible = False
                Find1.Command3.Enabled = False
                Find1.Command1.Visible = True
                Find1.Command1.Enabled = True
                Find1.Command5.Visible = True
                Find1.Command5.Enabled = True
                Find1.cboRep.Visible = True
                Find1.Label2.Visible = True
                Find1.cboRep.Clear
                Find1.Show
End Sub

'........................................
'Name: ctext
'Description: Clears all text boxes on a form
'........................................
Sub ctext(G As Form)
Dim CT As Control
Dim nm As String
On Local Error Resume Next
For Each CT In G
    nm = CT.Name
   If TypeOf CT Is RichTextBox Then CT.Text = ""
Next
End Sub
Public Function GETWord(File As String) As String
    Dim i As Integer
    For i = 1 To Len(File) Step 1
        If Mid(File, i, 1) = " " Then
            i = i - 1
            Exit For
        End If
    Next
    GETWord = Mid$(File, 1, i)
    strFirstW = GETWord
End Function
Public Function SelectWord() 'gets the word on or near the cursor
Dim SStart, SEnd As Integer
    'these characters will not be part of the start of the word
    frmMain.RichTextBox1(frmMain.Text8.Text).Span " ,;{}[]''()/\-=:.?!<>*+#^%$&", False, True
    SStart = frmMain.RichTextBox1(frmMain.Text8.Text).SelStart
    'these characters will not be part of the end of the word
    frmMain.RichTextBox1(frmMain.Text8.Text).Span " ,;{}[]''()/\-=:.?!<>*+#^%$&", True, True
    SEnd = frmMain.RichTextBox1(frmMain.Text8.Text).SelStart + frmMain.RichTextBox1(frmMain.Text8.Text).SelLength
If frmMain.RichTextBox1(frmMain.Text8.Text).SelText = "" Then Exit Function
    strFirstW = Trim(frmMain.RichTextBox1(frmMain.Text8.Text).SelText)
    frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = 0
End Function
Public Function Find() 'finds the first word
    Const conBtns As Integer = vbOKOnly + vbInformation _
            + vbDefaultButton1 + vbApplicationModal
    Const conMsg As String = "The search text was not found."
            
        strSearch = frmMain.RichTextBox1(frmMain.Text8.Text).Text
  
        strSearchFor = Find1.Text1.Text
        intFoundPos = InStr(1, strSearch, strSearchFor, 1)    'search for text

        If intFoundPos = 0 Then     'if text was not found
            MakeAlwaysOnTop Find1, False
            MsgBox conMsg, conBtns, "Find"
            MakeAlwaysOnTop Find1, True
            Find1.Command2.Caption = "&Find"
            

        Else            'if text was found
            frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = intFoundPos - 1  'highlight text
            frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(strSearchFor)
            frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
            Find1.Command2.Caption = "&Find Next"
        End If

End Function

Public Function FindNext() 'finds the next word ...
    Const conBtns As Integer = vbOKOnly + vbInformation _
        + vbDefaultButton1 + vbApplicationModal
    Const conMsg As String = "The search has been completed."

        intFoundPos = InStr(intBegSearch, strSearch, strSearchFor, 1)
        If intFoundPos = 0 Then 'if text was not found
            MakeAlwaysOnTop Find1, False
            MsgBox conMsg, conBtns, "Find Next"
            Find1.Command2.Caption = "&Find"
            MakeAlwaysOnTop Find1, True
        Else 'if text was found
            frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = intFoundPos - 1 'highlight text
            frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(strSearchFor)
   
            frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
        End If

End Function
Public Function vReplace() 'replace one word at a time


    Const conBtns As Integer = vbOKOnly + vbInformation _
            + vbDefaultButton1 + vbApplicationModal
    Const conMsg As String = "The search text was not found."
    
    strSearchFor = Find1.Text1.Text
   
    intFoundPos = InStr(1, frmMain.RichTextBox1(frmMain.Text8.Text).Text, strSearchFor, 1)   'search for text
        If intFoundPos = 0 Then          'if string not found
            MakeAlwaysOnTop Find1, False
            MsgBox conMsg, conBtns, "Replace"
            MakeAlwaysOnTop Find1, True
            Exit Function
        Else
            frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = intFoundPos - 1  'highlight text
            frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(strSearchFor)
            frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
            frmMain.RichTextBox1(frmMain.Text8.Text).SelText = Find1.cboRep.Text 'replace text with contents of combobox
        End If
End Function
Public Function Replace_All() 'replace all words in text box


Dim X As Integer
    Dim XC As Integer
    Const conBtns As Integer = vbOKOnly + vbInformation _
            + vbDefaultButton1 + vbApplicationModal
    Const conMsg As String = "The search text was not found."
   
    strSearchFor = Find1.Text1.Text
        Do
            intBegSearch = frmMain.RichTextBox1(frmMain.Text8.Text).SelStart + 2
            intFoundPos = InStr(1, frmMain.RichTextBox1(frmMain.Text8.Text).Text, strSearchFor, 1)
            If intFoundPos = 0 Then
                'XC = XC + 1
                MakeAlwaysOnTop Find1, False
                MsgBox "There were " & XC & " replacements made", , "Replace All"
                MakeAlwaysOnTop Find1, True
                Exit Function
            End If
            frmMain.RichTextBox1(frmMain.Text8.Text).SelStart = intFoundPos - 1 'highlight text
            frmMain.RichTextBox1(frmMain.Text8.Text).SelLength = Len(strSearchFor)
            frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
            frmMain.RichTextBox1(frmMain.Text8.Text).SelText = Find1.cboRep.Text
            XC = XC + 1
        Loop
End Function


