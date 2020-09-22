Attribute VB_Name = "Module2"
Option Explicit

Public Function IsPalindrome(ByVal strTest As String, _
                             Optional ByVal CaseSensitive As Boolean = False, _
                             Optional ByVal IgnoreSpaces As Boolean = True) As Boolean

    Dim CutLen As Long

'Optional parameters
    If Not CaseSensitive Then
'Simplify case if requried (local only because of 'ByVal')
        strTest = LCase$(strTest)
    End If
'
    If IgnoreSpaces Then
'Loop because removing blanks is tricky if there are multiple blanks
        Do While InStr(strTest, " ")
            strTest = Replace(strTest, " ", "")
        Loop
    End If
'
    If Len(strTest) / 2 = Len(strTest) \ 2 Then
        CutLen = Len(strTest) / 2 'Even numbered mirrored text ABCCBA
      Else
        CutLen = Len(strTest) \ 2 'Odd numbered pivoting text  ABCBA
'integer division means that neither half contains the pivot character.
    End If
    IsPalindrome = (Left$(strTest, CutLen)) = (StrReverse(Right$(strTest, CutLen)))

End Function

'This function converts numbers to their corresponding
'roman values
Public Function NumericToRoman(ByVal Value As Long) As String

Dim iPos As Integer, sBuffer As String, iReference As Integer
Dim sLowChar As String, sMidChar As String, sHighChar As String

On Error Resume Next
sBuffer = String$(Value \ 1000, "M")
Value = Value Mod 1000

iReference = 100
Do Until iReference = 0
If iReference = 100 Then
sHighChar = "M"
sMidChar = "D"
sLowChar = "C"
ElseIf iReference = 10 Then
sHighChar = "C"
sMidChar = "L"
sLowChar = "X"
Else
sHighChar = "X"
sMidChar = "V"
sLowChar = "I"
End If
iPos = Value \ iReference
If (iPos > 0) And (iPos < 4) Then
sBuffer = sBuffer & String$(iPos, sLowChar)
ElseIf iPos = 4 Then
sBuffer = sBuffer & sLowChar & sMidChar
ElseIf iPos = 5 Then
sBuffer = sBuffer & sMidChar
ElseIf (iPos > 5) And (iPos < 9) Then
sBuffer = sBuffer & sMidChar & String$(iPos - 5, sLowChar)
ElseIf iPos = 9 Then
sBuffer = sBuffer & sLowChar & sHighChar
End If
Value = Value - iReference * iPos
iReference = iReference \ 10
Loop
NumericToRoman = sBuffer

End Function


