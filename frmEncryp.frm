VERSION 5.00
Begin VB.Form frmEncryp 
   Caption         =   " ENCRYPT"
   ClientHeight    =   7785
   ClientLeft      =   7260
   ClientTop       =   1500
   ClientWidth     =   4350
   Icon            =   "frmEncryp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   4350
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   1380
      Picture         =   "frmEncryp.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Enrypt File"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   3240
      Picture         =   "frmEncryp.frx":1644
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Clear Text"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   3240
      Picture         =   "frmEncryp.frx":17AA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "End Application"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Top             =   7320
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8280
      TabIndex        =   6
      ToolTipText     =   "Enter file name to encrypt"
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   120
      Picture         =   "frmEncryp.frx":19FC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Clear All"
      Top             =   6360
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   645
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox TextPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   16
      PasswordChar    =   "?"
      TabIndex        =   1
      ToolTipText     =   "Choose a password 16 char. max."
      Top             =   540
      Width           =   2685
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   765
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   120
      Picture         =   "frmEncryp.frx":1ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Minimise"
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Text chars: 0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Text Length"
      Top             =   960
      Width           =   4050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypt File Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      TabIndex        =   13
      ToolTipText     =   "Text Length"
      Top             =   165
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Image:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   765
   End
End
Attribute VB_Name = "frmEncryp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************
'  Written by GioRock  *
'***********************
'***********************
'  Created by GioRock  *
'***********************

' Define all bitmap data
' to store into the file.
Private Type BITMAPFILEHEADER
    bfType As Integer           ' Specifies the file type. It must be BM.
    bfSize As Long              ' Specifies the size, in bytes, of the bitmap file.
    bfReserved1 As Integer      ' Reserved; must be zero.
    bfReserved2 As Integer      ' Reserved; must be zero.
    bfOffBits As Long           ' Specifies the offset, in bytes, from the BITMAPFILEHEADER
                                ' structure to the bitmap bits.
End Type
Private Type BITMAPINFOHEADER
    biSize As Long              ' Specifies the number of bytes required by the structure.
    biWidth As Long             ' Specifies the width of the bitmap, in pixels.
    biHeight As Long            ' Specifies the height of the bitmap, in pixels.
                                ' If biHeight is positive, the bitmap is a bottom-up DIB
                                ' and its origin is the lower left corner.
                                ' If biHeight is negative, the bitmap is a top-down DIB
                                ' and its origin is the upper left corner.
    biPlanes As Integer         ' Specifies the number of planes for the target device.
                                ' This value must be set to 1.
    biBitCount As Integer       ' Specifies the number of bits per pixel.
                                ' This value must be 1, 4, 8, 16, 24, or 32.
    biCompression As Long       ' Specifies the type of compression for a compressed
                                ' bottom-up bitmap (top-down DIBs cannot be compressed).
                                ' It can be one of the following values:
                                ' Value Description
                                ' BI_RGB:  An uncompressed format.
                                ' BI_RLE8: A run-length encoded (RLE) format for bitmaps
                                ' with 8 bits per pixel. The compression format is a
                                ' two-byte format consisting of a count byte followed
                                ' by a byte containing a color index.
                                ' For more information, see the following Remarks section.
                                ' BI_RLE4: An RLE format for bitmaps with 4 bits per pixel.
                                ' The compression format is a two-byte format consisting
                                ' of a count byte followed by two word-length color indices. For more information, see the following Remarks section.
                                ' BI_BITFIELDS: Specifies that the bitmap is not compressed
                                ' and that the color table consists of three doubleword
                                ' color masks that specify the red, green, and blue
                                ' components, respectively, of each pixel.
                                ' This is valid when used with 16- and 32-bits-per-pixel
                                ' bitmaps.
    biSizeImage As Long         ' Specifies the size, in bytes, of the image.
                                ' This may be set to 0 for BI_RGB bitmaps.
    biXPelsPerMeter As Long     ' Specifies the horizontal resolution, in pixels per meter,
                                ' of the target device for the bitmap. An application can
                                ' use this value to select a bitmap from a resource group
                                ' that best matches the characteristics of the current
                                ' device.
    biYPelsPerMeter As Long     ' Specifies the vertical resolution, in pixels per meter,
                                ' of the target device for the bitmap.
    biClrUsed As Long           ' Specifies the number of color indices in the color table
                                ' that are actually used by the bitmap. If this value is
                                ' zero, the bitmap uses the maximum number of colors
                                ' corresponding to the value of the biBitCount member for
                                ' the compression mode specified by biCompression.
                                ' If biClrUsed is nonzero and the biBitCount member is less
                                ' than 16, the biClrUsed member specifies the actual number
                                ' of colors the graphics engine or device driver accesses.
                                ' If biBitCount is 16 or greater, then biClrUsed member
                                ' specifies the size of the color table used to optimize
                                ' performance of Windows color palettes. If biBitCount
                                ' equals 16 or 32, the optimal color palette starts
                                ' immediately following the three doubleword masks.
                                ' If the bitmap is a packed bitmap (a bitmap in which the
                                ' bitmap array immediately follows the BITMAPINFO header and
                                ' which is referenced by a single pointer), the biClrUsed
                                ' member must be either 0 or the actual size of the color
                                ' table.
    biClrImportant As Long      ' Specifies the number of color indices that are considered
                                ' important for displaying the bitmap. If this value is
                                ' zero, all colors are important.
End Type

Private Const BF_TYPE = &H4D42  ' It must be BM.
Private Const HEADERLEN = &H36  ' Total length of File header,
                                ' Len(BITMAPFILEHEADER) + Len(BITMAPINFOHEADER).
Private Const BI_SIZE = &H28    ' Len(BITMAPINFOHEADER) structure.
Private Const COLOR256 = &H8    ' Define that we works with 256 color.

' Define color palette
' to store into the file.
Private Type PALETTEENTRY
    peRed As Byte               ' Specifies a red intensity value for the palette entry.
    peGreen As Byte             ' Specifies a green intensity value for the palette entry.
    peBlue As Byte              ' Specifies a blue intensity value for the palette entry.
    peFlags As Byte             ' Specifies how the palette entry is to be used.
                                ' The peFlags member may be set to NULL or one of the
                                ' following values:
                                ' PC_EXPLICIT: Specifies that the low-order word of the
                                ' logical palette entry designates a hardware palette index.
                                ' This flag allows the application to show the contents of
                                ' the display device palette.
                                ' PC_NOCOLLAPSE: Specifies that the color be placed in an
                                ' unused entry in the system palette instead of being
                                ' matched to an existing color in the system palette.
                                ' If there are no unused entries in the system palette, the
                                ' color is matched normally. Once this color is in the
                                ' system palette, colors in other logical palettes can be
                                ' matched to this color.
                                ' PC_RESERVED: Specifies that the logical palette entry be
                                ' used for palette animation. This flag prevents other
                                ' windows from matching colors to the palette entry since
                                ' the color frequently changes. If an unused system-palette
                                ' entry is available, the color is placed in that entry.
                                ' Otherwise, the color is not available for animation.
End Type
Private Type LOGPALETTE
    palVersion As Integer            ' Specifies the Windows version number for the
                                     ' structure (currently &H300).
    palNumEntries As Integer         ' Specifies the number of entries in the logical color
                                     ' palette.
    palPalEntry(255) As PALETTEENTRY ' Specifies an array of PALETTEENTRY structures that
                                     ' define the color and usage of each entry in the
                                     ' logical palette.
End Type

Private Const PAL_VERSION_ = &H300   ' It must be &H300.
Private Const PAL_COLOR_256 = &H100  ' Maximum number of PALETTEENTRY.
Private Const PAL_LEN_256 = &H400    ' Length of current PALETTEENTRY array.

' Retrieve current Windows Directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Const myerrfilepath = 53

Private Function CryptText(ByVal sTextToCrypt As String, ByVal sPassword As String) As String
Dim l As Long
Dim sTempText As String
Dim sTmpPwd As String
Dim sCryptedText As String
Dim ch1 As String * 1
Dim ch2 As String * 1
Dim chResult As String * 1
'***********************
'  Created by GioRock  *
'***********************

    ' This show a typical Algorithm to Crypt a message with password
    ' You can create your custom algorithm too, if you are able to do this...
    
    ' Reverse Text to Crypt
    sTempText = StrReverse(sTextToCrypt)
    
    ' Create capable Buffer to store Crypted Text
    sCryptedText = String$(Len(sTempText), 32)
    
    ' Create capable Buffer to store Password Text
    sTmpPwd = String$(Len(sTempText), 32)
    ' Remember: Buffer String is faster then String chain like commented sample below
    ' sTmpPwd = sTmpPwd + ......
    
    ' Here create a String by Password Text of same length of Text to Crypt
    For l = 1 To Len(sTempText) Step Len(sPassword)
        ' I scramble Password slightly right rather reversing it
        Mid$(sTmpPwd, l, Len(sPassword)) = IIf(l Mod 3 = 0, sPassword, StrReverse(sPassword))
    Next l
    
    ' Crypto Algorithm
    For l = 1 To Len(sTempText)
        ' Get one char at time by Text and Password
        ch1 = Mid$(sTempText, l, 1)
        ch2 = Mid$(sTmpPwd, l, 1)
        ' Uses Xor operator to obtain one char from two
        ' basing on 255 total font index chars at time
        ' Here you can modify this Algorithm and view the
        ' unpredictable resultant.....
        chResult = Chr$(Abs(255 Xor Asc(ch1) Xor Asc(ch2)))
        ' Store result in buffer
        Mid$(sCryptedText, l, 1) = chResult
    Next l
        
    ' Return Crypted String
    CryptText = sCryptedText

End Function



Private Sub CenterPic()
    'Decide where place Picturebox in Form
    With Picture1
        .Move (Me.ScaleWidth - .Width) / 2, (Me.ScaleHeight - .Height) / 2
    End With
End Sub

Private Sub ReSampleText256(sText As String)
'***********************
'  Created by GioRock  *
'***********************
    
    sText = RTrim$(sText)
    
    ' Adjust Text length in manner
    ' to divide it in the middle
    ' Add space since rest (Len(sText) / 2 = 0)
    Do While Len(sText) Mod 2 <> 0
        sText = sText + " "
    Loop

End Sub

Private Sub CalculatePicDimension256(ByVal sText As String)
Dim w As Single, h As Single
'***********************
'  Created by GioRock  *
'***********************

    ' Calculate height first
    h = Sqr(Len(sText) / 2)
    ' Width and Height are the same
    w = h
    
    ' Set Picture Dimension
    Picture1.Width = w
    Picture1.Height = h
    
End Sub


Private Function InitPalette256() As LOGPALETTE
Dim i As Integer
Dim LP As LOGPALETTE

    ' Store all PALETTEENTRY color
    With LP
        .palVersion = PAL_VERSION_
        .palNumEntries = PAL_COLOR_256
        ' I choose a Gray Scale palette color
        ' You can create custom palette color
        ' modifing RGB bytes at your choice
        ' This not compromise the result
        ' but change bitmap final aspect only
        For i = 0 To PAL_COLOR_256 - 1
            .palPalEntry(i).peRed = CByte(i)
            .palPalEntry(i).peGreen = CByte(i)
            .palPalEntry(i).peBlue = CByte(i)
            .palPalEntry(i).peFlags = CByte(0)
        Next i
    End With
    
    InitPalette256 = LP
    
End Function

Private Function WidthBytes(ByVal lWide As Long, lBits As Long) As Long
    ' Standard function to retrieve Width Bytes in Bitmap
    WidthBytes = ((((lWide * lBits) + &H1F) And &HFFE0) / 8)
End Function

Private Sub Command1_Click()
Dim i As Integer, hff As Integer
Dim lPixW As Long, lPixH As Long
Dim lWB As Integer
Dim sText As String
Dim sNotepad As String
Dim BFH As BITMAPFILEHEADER
Dim BIH As BITMAPINFOHEADER
Dim LP As LOGPALETTE
'***********************
'  Created by GioRock  *
'***********************

    ' Check for valid Password length
    If Len(TextPassword.Text) = 0 Then
        MsgBox "Password required!!!", vbInformation
        TextPassword.SetFocus
        Exit Sub
    End If
    
    ' Check for minimum number of characters in Text to diplay bitmap correctly
    ' I set this to 16 chrs. Sqr(16) = 4 / 2 = 2, bitmap width = 2, bitmap height = 2
    ' So, minimum picture size is 2x2px.
    If Len(frmMain.RichTextBox1(frmMain.Text8.Text).Text) < TextPassword.MaxLength Then
        MsgBox CStr(TextPassword.MaxLength) + " characters minimum required!!!", vbInformation
        frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Store text to Ctypt in a temporary variable
    sText = frmMain.RichTextBox1(frmMain.Text8.Text).Text
    
    ' Hide picture momentarily
    ' during process
    Picture1.Visible = False
    
    ' Adjust Text length in manner
    ' to divide it in the middle
    ReSampleText256 sText
    ' Calculate picture dimension
    ' by Text length
    CalculatePicDimension256 sText
    ' Center picture to display
    CenterPic
    
    ' Go to Crypto algorithm function
    sText = CryptText(sText, TextPassword.Text)
    
    ' Get all pallette entry
    LP = InitPalette256()
    
    ' Store bitmap width and bitmap height
    ' in these variable
    lPixW = Picture1.ScaleWidth
    lPixH = Picture1.ScaleHeight
    
    ' Calculate Width Bytes with standard function
    ' like manuals
    lWB = WidthBytes(lPixW, COLOR256)
    
    ' All Referred in header declarations explanation
    With BFH
        .bfType = BF_TYPE
        .bfOffBits = HEADERLEN + PAL_LEN_256
        .bfSize = .bfOffBits + lWB * lPixH
        .bfReserved1 = 0
        .bfReserved2 = 0
    End With
    
    ' and so this
    With BIH
        .biSize = BI_SIZE
        .biWidth = lPixW
        .biHeight = lPixH
        .biPlanes = 1
        .biBitCount = COLOR256
        .biSizeImage = lPixW * lPixH
        .biClrImportant = 0
    End With
    
    ' Delete File if exist
    If Dir$(App.Path + "\" + Text1.Text + ".bmp") <> "" Then
        Kill App.Path + "\" + Text1.Text + ".bmp"
    End If
    
    ' Save all work in Bitmap Binary File
    hff = FreeFile
    Open App.Path + "\" + Text1.Text + ".bmp" For Binary Access Write As #hff
        ' First BITMAPFILEHEADER structure
        Put #hff, , BFH
        ' Second BITMAPINFOHEADER structure
        Put #hff, , BIH
        ' Then PALETTEENTRY
        For i = 0 To PAL_COLOR_256 - 1
            Put #hff, , LP.palPalEntry(i)
        Next i
        ' Last Crypted Text
        Put #hff, , sText
    Close #hff
    
    ' Load Bitmap Binary File to display in Picture
    Picture1.Picture = LoadPicture(App.Path + "\" + Text1.Text + ".bmp")
    
    ' Now you can show the result
    Picture1.Visible = True
    Picture1.Refresh
    
    ' Enable Button UnCrypt
    'Command2.Enabled = True
    
'    ' Check for Notepad existence
'    If Check1.Value = vbChecked Then
'        sNotepad = GetWinDir() + "Notepad.exe"
'        If Dir$(sNotepad) <> "" Then
'            ' Shell Notepad to view data transformation
'            Shell sNotepad + " " + App.Path + "\" + Text1.Text + ".bmp", vbNormalFocus
'        Else
'            Screen.MousePointer = vbDefault
'            MsgBox "Notepad not found!!!", vbInformation
'        End If
'    End If
'
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Command2_Click()
Dim i As Integer, hff As Integer
Dim lPixW As Long, lPixH As Long
Dim lWB As Integer
Dim sText As String
Dim BFH As BITMAPFILEHEADER
Dim BIH As BITMAPINFOHEADER
Dim LP As LOGPALETTE
'***********************
'  Created by GioRock  *
'***********************

    ' Check for Crypto Bitmap
    If IsNull(Picture1.Picture) Then
        MsgBox "Picture required!!!", vbInformation
        frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
        Exit Sub
    End If
    
    ' Check for valid Password length
    If Len(TextPassword.Text) = 0 Then
        MsgBox "Password required!!!", vbInformation
        TextPassword.SetFocus
        Exit Sub
    End If
    
    ' Check for valid Crypto Bitmap File
    If Dir$(App.Path + "\" + Text2.Text + ".bmp") = "" Then
        MsgBox "Bitmap File required!!!", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    frmMain.RichTextBox1(frmMain.Text8.Text).Text = ""
 
    hff = FreeFile
    Open App.Path + "\" + Text2.Text + ".bmp" For Binary Access Read As #hff
        ' Skip BITMAPFILEHEADER structure
        Get #hff, , BFH
        ' Skip BITMAPINFOHEADER structure
        Get #hff, , BIH
        ' Check for a valid Crypto Bitmap
        If BIH.biBitCount = COLOR256 Then
            ' Skip PALETTEENTRY
            For i = 0 To PAL_COLOR_256 - 1
                Get #hff, , LP.palPalEntry(i)
            Next i
            ' (FileLen - skipped strucure length) = Len(Crypted Text)
            ' Create Buffer to store data
            sText = String$(FileLen(App.Path + "\" + Text2.Text + ".bmp") - (HEADERLEN + PAL_LEN_256), 32)
            ' Catch our Crypto Text
            Get #hff, , sText
        Else
            Screen.MousePointer = vbDefault
            MsgBox "Not a valid Crypto Bitmap!!!", vbInformation
            Close #hff
            Exit Sub
        End If
    Close #hff
    
    ' Go to UnCrypt Algorithm
    frmMain.RichTextBox1(frmMain.Text8.Text).Text = UnCryptText(sText, TextPassword.Text)
   
    'Command1.Enabled = True
    
    Screen.MousePointer = vbDefault
  Text3.Text = Text2.Text
End Sub

Private Sub Command3_Click()
    frmMain.RichTextBox1(frmMain.Text8.Text).Text = ""
    frmMain.RichTextBox1(frmMain.Text8.Text).SetFocus
End Sub

Private Sub Command4_Click()
Call Form_Load
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
' Load Bitmap Binary File to display in Picture
    Picture1.Picture = LoadPicture(App.Path + "\" + Text2.Text + ".bmp")
    'Command2.Enabled = True
End Sub

Private Sub Command7_Click()
frmMain.RichTextBox1(frmMain.Text8.Text).Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
TextPassword.Text = ""
Command1.Enabled = False
'Command2.Enabled = False
'Command4.Enabled = False
'Command6.Enabled = False
Picture1.Picture = Picture2.Image
End Sub

Private Sub Command8_Click()
Me.WindowState = vbMinimized
Me.Caption = ""
End Sub

Private Sub Form_Load()

On Error GoTo fubar
Dim Msg As String
Dim hff As Integer

'Dim TopCorner As Integer
'  Dim LeftCorner As Integer
'  'centres the form on the screen
'  If Me.WindowState <> 0 Then Exit Sub
'
'  TopCorner = (Screen.Height - Me.Height) \ 2
'  LeftCorner = (Screen.Width - Me.Width) \ 2
'  Me.Move LeftCorner, TopCorner

    
    hff = FreeFile
    Open App.Path + "\" + Text3.Text + ".txt" For Input As #hff '
        frmMain.RichTextBox1(frmMain.Text8.Text).Text = Input(LOF(hff), #hff)
    Close #hff
    
    'CenterPic
  'Grand Prix Red Rose
      Exit Sub
fubar:
      If (Err.Number = myerrfilepath) Then
        Msg = "you must open a file to begin" _
        & vbCrLf & "when the program loads"
        If MsgBox(Msg) = vbOK Then
          frmEncryp.Show
          
        End If
      End If
      Exit Sub
  
End Sub

Private Function UnCryptText(ByVal sCryptedText As String, ByVal sPassword As String) As String
Dim l As Long
Dim sTempCryptedText As String
Dim sTmpPwd As String
Dim sUnCryptedText As String
Dim ch1 As String * 1
Dim ch2 As String * 1
Dim chResult As String * 1
'***********************
'  Created by GioRock  *
'***********************
    
    ' Same as Crypto Algorithm but in reversed order
    ' that's all

    sTempCryptedText = sCryptedText
    
    sUnCryptedText = String$(Len(sTempCryptedText), 32)
    
    sTmpPwd = String$(Len(sTempCryptedText), 32)
    For l = 1 To Len(sTempCryptedText) Step Len(sPassword)
        Mid$(sTmpPwd, l, Len(sPassword)) = IIf(l Mod 3 = 0, sPassword, StrReverse(sPassword))
    Next l
    
    For l = 1 To Len(sTempCryptedText)
        ch1 = Mid$(sTempCryptedText, l, 1)
        ch2 = Mid$(sTmpPwd, l, 1)
        chResult = Chr$(Abs(255 Xor Asc(ch1) Xor Asc(ch2)))
        Mid$(sUnCryptedText, l, 1) = chResult
    Next l
    
    UnCryptText = RTrim$(StrReverse(sUnCryptedText))

End Function

Private Sub Form_Paint()
If Me.WindowState = 0 Then
Me.Caption = "CryptoPic"
End If
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Command1.Enabled = True
ElseIf Text1.Text = "" Then
Command1.Enabled = False
End If

End Sub

'Private Sub Text2_Change()
'If Text2.Text <> "" Then
'Command6.Enabled = True
'ElseIf Text2.Text = "" Then
'Command6.Enabled = False
'End If
'End Sub
'
'Private Sub Text3_Change()
'If Text3.Text <> "" Then
'Command4.Enabled = True
'ElseIf Text3.Text = "" Then
'Command4.Enabled = False
'End If
'End Sub

Private Sub TextCrypto_Change()
    Label2.Caption = "Text chars: " + CStr(Len(frmMain.RichTextBox1(frmMain.Text8.Text).Text))
    If Len(frmMain.RichTextBox1(frmMain.Text8.Text).Text) = 0 Then
'        Command1.Enabled = False
'    Else
'        Command1.Enabled = True
    End If
End Sub

Private Function GetWinDir() As String
Dim sWD As String

    ' Make a Buffer to store
    ' Windows Directory
    sWD = String$(128, 32)
    
    ' Call API to get path
    GetWindowsDirectory sWD, 128
    
    ' Check for zero terminated string
    If InStr(sWD, Chr$(0)) <> 0 Then
        ' strip null char if exist
        sWD = Left$(sWD, InStr(sWD, Chr$(0)) - 1)
    End If
    
    ' Add separator "\" to path if not
    GetWinDir = RTrim$(IIf(Right$(sWD, 1) = "\", sWD, sWD + "\"))
    
End Function



