VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Comprehensive Note Pad"
   ClientHeight    =   10665
   ClientLeft      =   225
   ClientTop       =   345
   ClientWidth     =   14865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   14865
   Begin VB.CommandButton Command11 
      Height          =   375
      Index           =   3
      Left            =   4920
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   780
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Index           =   2
      Left            =   4560
      Picture         =   "frmMain.frx":0E14
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   780
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Index           =   1
      Left            =   4200
      Picture         =   "frmMain.frx":0F5E
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   780
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Index           =   0
      Left            =   3825
      Picture         =   "frmMain.frx":10A8
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   780
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Height          =   375
      Left            =   14400
      Picture         =   "frmMain.frx":11F2
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   660
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   13320
      Picture         =   "frmMain.frx":177C
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   9600
      Picture         =   "frmMain.frx":18C6
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   11880
      Picture         =   "frmMain.frx":1A10
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   6840
      TabIndex        =   179
      Top             =   1680
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   11280
      Picture         =   "frmMain.frx":1F9A
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   28
      Left            =   10680
      Picture         =   "frmMain.frx":2524
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   20
      Left            =   240
      Picture         =   "frmMain.frx":2AAE
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1200
      TabIndex        =   170
      Text            =   "4"
      Top             =   8085
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1200
      TabIndex        =   169
      Text            =   "3"
      Top             =   8445
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1200
      TabIndex        =   168
      Text            =   "2"
      Top             =   8820
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   27
      Left            =   10200
      Picture         =   "frmMain.frx":3038
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   19
      Left            =   8880
      Picture         =   "frmMain.frx":35C2
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   1200
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   18
      Left            =   9270
      TabIndex        =   155
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   13920
      Picture         =   "frmMain.frx":3B4C
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   660
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   26
      Left            =   8400
      Picture         =   "frmMain.frx":40D6
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   25
      Left            =   7920
      Picture         =   "frmMain.frx":4660
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   23
      Left            =   6720
      Picture         =   "frmMain.frx":4BEA
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   24
      Left            =   7200
      Picture         =   "frmMain.frx":5174
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Index           =   22
      Left            =   14160
      Picture         =   "frmMain.frx":56FE
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   21
      Left            =   1200
      Picture         =   "frmMain.frx":5C88
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   9360
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   12000
      TabIndex        =   131
      Text            =   "1"
      Top             =   10065
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   10680
      TabIndex        =   130
      Text            =   "1"
      Top             =   10065
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   129
      Text            =   "hi lite word"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   127
      Text            =   "word cnt"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   20
      Left            =   240
      Picture         =   "frmMain.frx":6212
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   6120
      Width           =   375
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   390
      Left            =   2160
      TabIndex        =   76
      Top             =   1740
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8520
      MouseIcon       =   "frmMain.frx":679C
      MousePointer    =   99  'Custom
      TabIndex        =   73
      Text            =   "0"
      ToolTipText     =   "LINE SPACING: 0 = NORMAL, 1 =  1.5 SPACING, 2 = DOUBLE SPACED"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   18
      Left            =   8040
      Picture         =   "frmMain.frx":68EE
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   17
      Left            =   1200
      Picture         =   "frmMain.frx":6E78
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   16
      Left            =   720
      Picture         =   "frmMain.frx":7402
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   15
      Left            =   240
      Picture         =   "frmMain.frx":798C
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":7F16
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   14
      Left            =   11655
      Picture         =   "frmMain.frx":84A0
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   120
      Width           =   375
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   1
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":8A2A
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   13
      Left            =   14400
      Picture         =   "frmMain.frx":8AAC
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   12
      Left            =   1200
      Picture         =   "frmMain.frx":9036
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   9
      Left            =   6120
      Picture         =   "frmMain.frx":95C0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   18
      Left            =   5370
      Picture         =   "frmMain.frx":9B4A
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   17
      Left            =   5745
      Picture         =   "frmMain.frx":A0D4
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   16
      Left            =   11715
      TabIndex        =   61
      Top             =   435
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   16
      Left            =   10680
      Picture         =   "frmMain.frx":A65E
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   330
      Index           =   15
      Left            =   10740
      TabIndex        =   59
      Top             =   465
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   11
      Left            =   4935
      Picture         =   "frmMain.frx":ABE8
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   14
      Left            =   4995
      TabIndex        =   57
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   14
      Left            =   9720
      Picture         =   "frmMain.frx":B172
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   15
      Left            =   9360
      Picture         =   "frmMain.frx":B6FC
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   13
      Left            =   9780
      TabIndex        =   55
      Top             =   450
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   330
      Index           =   12
      Left            =   9420
      TabIndex        =   54
      Top             =   465
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   13
      Left            =   13200
      Picture         =   "frmMain.frx":BC86
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   11
      Left            =   12600
      Picture         =   "frmMain.frx":C210
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   12
      Left            =   3240
      Picture         =   "frmMain.frx":C79A
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   11
      Left            =   3285
      TabIndex        =   48
      Top             =   435
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   10
      Left            =   12225
      Picture         =   "frmMain.frx":CD24
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   9
      Left            =   2160
      Picture         =   "frmMain.frx":D2AE
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   8
      Left            =   2520
      Picture         =   "frmMain.frx":D838
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   7
      Left            =   2880
      Picture         =   "frmMain.frx":DDC2
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   9
      Left            =   12285
      TabIndex        =   46
      Top             =   450
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   8
      Left            =   2580
      TabIndex        =   45
      Top             =   450
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   345
      Index           =   7
      Left            =   2940
      TabIndex        =   44
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   3
      Left            =   6705
      Picture         =   "frmMain.frx":E34C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   4
      Left            =   7065
      Picture         =   "frmMain.frx":E8D6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   5
      Left            =   7425
      Picture         =   "frmMain.frx":EE60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   6
      Left            =   10320
      Picture         =   "frmMain.frx":F3EA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   6
      Left            =   10380
      TabIndex        =   43
      Top             =   450
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   330
      Index           =   5
      Left            =   7485
      TabIndex        =   42
      Top             =   465
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   4
      Left            =   7125
      TabIndex        =   41
      Top             =   435
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   3
      Left            =   6765
      TabIndex        =   40
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   2
      Left            =   4560
      Picture         =   "frmMain.frx":F974
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   1
      Left            =   4185
      Picture         =   "frmMain.frx":FEFE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   330
      Index           =   2
      Left            =   4620
      TabIndex        =   39
      Top             =   465
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   1
      Left            =   4245
      TabIndex        =   38
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   0
      Left            =   3825
      Picture         =   "frmMain.frx":10488
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   0
      Left            =   3885
      TabIndex        =   37
      Top             =   450
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   1200
      Top             =   9840
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "SaMpLe"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   480
      TabIndex        =   32
      Top             =   1920
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   31
      Top             =   10365
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "frmMain.frx":10A12
            TextSave        =   "20/08/2007"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "frmMain.frx":10FAC
            TextSave        =   "16:45"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Comprehensive Note Pad an open source text editor"
            TextSave        =   "Comprehensive Note Pad an open source text editor"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   30
      Text            =   "1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   9840
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   10
      Left            =   14400
      Picture         =   "frmMain.frx":11546
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   8
      Left            =   11280
      Picture         =   "frmMain.frx":11AD0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      MouseIcon       =   "frmMain.frx":1205A
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Text            =   "10"
      Top             =   1230
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      MouseIcon       =   "frmMain.frx":121AC
      MousePointer    =   99  'Custom
      Sorted          =   -1  'True
      TabIndex        =   11
      Text            =   "Arial"
      Top             =   1230
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   7
      Left            =   13920
      Picture         =   "frmMain.frx":122FE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   6
      Left            =   13920
      Picture         =   "frmMain.frx":12888
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   5
      Left            =   720
      Picture         =   "frmMain.frx":12E12
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":1339C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   3
      Left            =   1560
      Picture         =   "frmMain.frx":13926
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   2
      Left            =   480
      Picture         =   "frmMain.frx":13EB0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   840
      Picture         =   "frmMain.frx":1443A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":149C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2160
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   2160
      Width           =   11985
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8
         X2              =   8
         Y1              =   0
         Y2              =   56
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   0
         Picture         =   "frmMain.frx":14F4E
         Top             =   0
         Width           =   11925
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   10
      Left            =   11340
      TabIndex        =   47
      Top             =   435
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   360
      Index           =   17
      Left            =   8100
      TabIndex        =   75
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   19
      Left            =   1200
      Picture         =   "frmMain.frx":1734C
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8520
      TabIndex        =   74
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox rtbBuffer 
      Height          =   7335
      Left            =   2160
      TabIndex        =   137
      Top             =   2640
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12938
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":178D6
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   2
      Left            =   2160
      TabIndex        =   138
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":1795E
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   3
      Left            =   2160
      TabIndex        =   139
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":179E0
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   4
      Left            =   2160
      TabIndex        =   140
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17A62
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   5
      Left            =   2160
      TabIndex        =   141
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17AE4
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   6
      Left            =   2160
      TabIndex        =   142
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17B66
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   7
      Left            =   2160
      TabIndex        =   143
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17BE8
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   8
      Left            =   2160
      TabIndex        =   144
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17C6A
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   9
      Left            =   2160
      TabIndex        =   145
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17CEC
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Index           =   10
      Left            =   2160
      TabIndex        =   146
      Top             =   2640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":17D6E
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   128
      Text            =   "11"
      Top             =   6960
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17DF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   21
      Left            =   1080
      Picture         =   "frmMain.frx":1838A
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   1920
      ScaleHeight     =   1155
      ScaleWidth      =   6165
      TabIndex        =   160
      Top             =   2760
      Width           =   6225
      Begin VB.CheckBox chkAware 
         Caption         =   "Ignore Spaces"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   163
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkAware 
         Caption         =   "Case Sensitive"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   162
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   360
         TabIndex        =   161
         Text            =   "Palindrome"
         Top             =   405
         Width           =   5775
      End
      Begin VB.Label Label10 
         Caption         =   "Check if word or phrase is a palindrome"
         Height          =   255
         Left            =   360
         TabIndex        =   165
         Top             =   45
         Width           =   4695
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000FF&
         Caption         =   "P A L I N"
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   30
         TabIndex        =   164
         Top             =   75
         Width           =   135
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000FF&
         Height          =   1455
         Left            =   0
         TabIndex        =   166
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox Text16 
      Height          =   7335
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   188
      Top             =   2640
      Width           =   12495
   End
   Begin VB.CommandButton Command9 
      Height          =   375
      Left            =   11880
      TabIndex        =   189
      Top             =   1080
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   1200
      TabIndex        =   197
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check3"
      Height          =   255
      Left            =   3360
      TabIndex        =   198
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   1935
      Left            =   1920
      ScaleHeight     =   1875
      ScaleWidth      =   2805
      TabIndex        =   180
      Top             =   3960
      Width           =   2865
      Begin VB.CommandButton Command19 
         Height          =   375
         Left            =   2160
         Picture         =   "frmMain.frx":18914
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   185
         Text            =   "1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Enter line number to go to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   187
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H000000FF&
         Caption         =   "G O T O   L  I N E"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   30
         TabIndex        =   181
         Top             =   75
         Width           =   135
      End
      Begin VB.Label Label21 
         BackColor       =   &H000000FF&
         Height          =   1935
         Left            =   0
         TabIndex        =   182
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   24
      X1              =   -120
      X2              =   3720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   24
      X1              =   -840
      X2              =   3720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disable Control Tabbing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   191
      Top             =   960
      Width           =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   23
      X1              =   12360
      X2              =   12360
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   22
      X1              =   11760
      X2              =   11760
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Insert Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   176
      Top             =   7800
      Width           =   1560
   End
   Begin VB.Label Label16 
      Caption         =   "Cols"
      Height          =   255
      Left            =   720
      TabIndex        =   173
      Top             =   8100
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Rows"
      Height          =   255
      Left            =   720
      TabIndex        =   172
      Top             =   8460
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Width"
      Height          =   255
      Left            =   720
      TabIndex        =   171
      Top             =   8835
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   21
      X1              =   10080
      X2              =   10080
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label13 
      Caption         =   "Palindrome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   167
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbllines 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   159
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label lblCol 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6840
      TabIndex        =   158
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label lblRow 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8400
      TabIndex        =   157
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   20
      X1              =   1800
      X2              =   -240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sring Functions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   153
      Top             =   960
      Width           =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   19
      X1              =   7800
      X2              =   7800
      Y1              =   1680
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   18
      X1              =   1800
      X2              =   -120
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   17
      X1              =   3720
      X2              =   3720
      Y1              =   825
      Y2              =   1665
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   16
      X1              =   11160
      X2              =   11160
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "No of words Hi lited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   147
      Top             =   6960
      Width           =   675
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Word Lister"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   134
      Top             =   9435
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "TAB cnt"
      Height          =   240
      Index           =   49
      Left            =   10035
      TabIndex        =   133
      Top             =   10095
      Width           =   660
   End
   Begin VB.Label Label3 
      Caption         =   "TAB sel"
      Height          =   240
      Index           =   2
      Left            =   11370
      TabIndex        =   132
      Top             =   10095
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   48
      Left            =   1095
      TabIndex        =   125
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   47
      Left            =   1095
      TabIndex        =   124
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   46
      Left            =   1095
      TabIndex        =   123
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   45
      Left            =   945
      TabIndex        =   122
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   44
      Left            =   945
      TabIndex        =   121
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   43
      Left            =   945
      TabIndex        =   120
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   42
      Left            =   1395
      TabIndex        =   119
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   41
      Left            =   1395
      TabIndex        =   118
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   40
      Left            =   1395
      TabIndex        =   117
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   39
      Left            =   1245
      TabIndex        =   116
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   38
      Left            =   1245
      TabIndex        =   115
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   37
      Left            =   1245
      TabIndex        =   114
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   36
      Left            =   795
      TabIndex        =   113
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   35
      Left            =   795
      TabIndex        =   112
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   34
      Left            =   795
      TabIndex        =   111
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   33
      Left            =   645
      TabIndex        =   110
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   32
      Left            =   645
      TabIndex        =   109
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   31
      Left            =   645
      TabIndex        =   108
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   30
      Left            =   1095
      TabIndex        =   107
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   29
      Left            =   1095
      TabIndex        =   106
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   28
      Left            =   1095
      TabIndex        =   105
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   27
      Left            =   945
      TabIndex        =   104
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   26
      Left            =   945
      TabIndex        =   103
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   25
      Left            =   945
      TabIndex        =   102
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   24
      Left            =   1395
      TabIndex        =   101
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   23
      Left            =   1395
      TabIndex        =   100
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   22
      Left            =   1395
      TabIndex        =   99
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   21
      Left            =   1245
      TabIndex        =   98
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   20
      Left            =   1245
      TabIndex        =   97
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   19
      Left            =   1245
      TabIndex        =   96
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   18
      Left            =   495
      TabIndex        =   95
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   17
      Left            =   495
      TabIndex        =   94
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   16
      Left            =   495
      TabIndex        =   93
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   15
      Left            =   345
      TabIndex        =   92
      Top             =   5910
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   14
      Left            =   345
      TabIndex        =   91
      Top             =   5760
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   13
      Left            =   345
      TabIndex        =   90
      Top             =   5610
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   12
      Left            =   795
      TabIndex        =   89
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   11
      Left            =   795
      TabIndex        =   88
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   10
      Left            =   795
      TabIndex        =   87
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   645
      TabIndex        =   86
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   645
      TabIndex        =   85
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   645
      TabIndex        =   84
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   495
      TabIndex        =   83
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   495
      TabIndex        =   82
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   495
      TabIndex        =   81
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   345
      TabIndex        =   80
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   345
      TabIndex        =   79
      Top             =   5310
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   345
      TabIndex        =   78
      Top             =   5460
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Hi-lite Colour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   4920
      Width           =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   6600
      X2              =   13800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "find / spell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   72
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   -240
      X2              =   1800
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   15
      X1              =   1800
      X2              =   -240
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   14
      X1              =   11160
      X2              =   11160
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1800
      X2              =   1800
      Y1              =   1680
      Y2              =   10680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   9240
      X2              =   9240
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Tick the box to select all text "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   51
      Top             =   555
      Width           =   2415
   End
   Begin VB.Label lblWordCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   36
      Top             =   10080
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   13
      X1              =   3720
      X2              =   3720
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Fonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   1680
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   12
      X1              =   15120
      X2              =   1800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   11
      X1              =   10200
      X2              =   10200
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Label Label3 
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   28
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Font Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   9
      X1              =   -120
      X2              =   1800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   13800
      X2              =   13800
      Y1              =   -360
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   12120
      X2              =   12120
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   13080
      X2              =   13080
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   6600
      X2              =   6600
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cryptography"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   13800
      X2              =   13800
      Y1              =   -240
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   1800
      X2              =   1800
      Y1              =   1680
      Y2              =   10560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      X1              =   12120
      X2              =   12120
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   6
      X1              =   1800
      X2              =   15480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   8
      X1              =   13080
      X2              =   13080
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   9
      X1              =   -240
      X2              =   1800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   12
      X1              =   6600
      X2              =   6600
      Y1              =   -120
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   11
      X1              =   10200
      X2              =   10200
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   10
      X1              =   1800
      X2              =   -120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   13
      X1              =   1800
      X2              =   -240
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   9240
      X2              =   9240
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   14
      X1              =   11160
      X2              =   11160
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   15
      X1              =   3720
      X2              =   3720
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   17
      X1              =   3720
      X2              =   3720
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   16
      X1              =   1800
      X2              =   -240
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   19
      X1              =   7800
      X2              =   7800
      Y1              =   1680
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   18
      X1              =   11160
      X2              =   11160
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   20
      X1              =   1800
      X2              =   -120
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   21
      X1              =   10080
      X2              =   10080
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   22
      X1              =   11760
      X2              =   11760
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      X1              =   13800
      X2              =   6600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   23
      X1              =   12360
      X2              =   12360
      Y1              =   840
      Y2              =   1680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    
Private Const MAX_TAB_STOPS = 32&
Private Const EM_SETPARAFORMAT = &H447
Private Const PFM_LINESPACING = &H100&

Private Type PARAFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    wNumbering As Integer
    wReserved As Integer
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long
    dySpaceBefore As Long          ' Vertical spacing before para
    dySpaceAfter As Long           ' Vertical spacing after para
    dyLineSpacing As Long          ' Line spacing depending on Rule
    sStyle As Integer              ' Style handle
    bLineSpacingRule As Byte       ' Rule for line spacing
    bCRC As Byte                   ' Reserved for CRC for rapid searching
    wShadingWeight As Integer      ' Shading in hundredths of a per cent
    wShadingStyle As Integer       ' Nibble 0: style, 1: cfpat, 2: cbpat
    wNumberingStart As Integer     ' Starting value for numbering
    wNumberingStyle As Integer     ' Alignment, roman/arabic, (), ), .,     etc.
    wNumberingTab As Integer       ' Space between 1st indent and 1st-line text
    wBorderSpace As Integer        ' Space between border and text(twips)
    wBorderWidth As Integer        ' Border pen width (twips)
    wBorders As Integer            ' Byte 0: bits specify which borders; Nibble 2: border style; 3: color                                     index*/
End Type
              
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Const EM_LINEFROMCHAR As Long = &HC9
Const EM_LINEINDEX As Long = &HBB
Const EM_GETLINECOUNT As Long = &HBA
Const EM_LINESCROLL As Long = &HB6


Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_USER = &H400
Const EM_CANUNDO = &HC6
Const EM_UNDO = &HC7

Dim SFile As String
Dim Startupcomplete
Dim mdpath
Dim asdsd As String
Dim nul As String

'Dim strForFind As String
'Dim strText As String

Dim TextHeigth As Long, fTop As Integer  ' Text height - important
Dim LineCount As Integer
Dim LineCountChange As Integer           ' This is used to determin if we need _
                                             to redraw the numbers
Dim FirstLine As Long                    ' Dim the First visible line
Dim FirstLineNow As Long
Dim blues, greens, reds, colours
Const myerrfilepath = 380
Const myerr = 5

Private Declare Function ReleaseCapture Lib "user32" () As Long ' drag n drop
Private CurX As Double
Private CurY As Double
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'for URL Linking
Private Const EM_CHARFROMPOS& = &HD7
Dim XX1 As Single
Dim YY1 As Single
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Const WM_USER = &H400
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
Private Const EM_AUTOURLDETECT = (WM_USER + 91)


Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Index = 0 To 18
Check1(Index).ToolTipText = "TICK THE BOX TO SELECT ALL TEXT"
Check1(Index).MousePointer = 99
Check1(Index).MouseIcon = LoadResPicture(103, vbResCursor)
Next Index
End Sub

Private Sub Check2_Click()

If Check2.Value = 1 Then
Check2.Picture = LoadResPicture(102, vbResBitmap)
Call Rtts
RichTextBox1(Text8.Text).SetFocus
ElseIf Check2.Value = 0 Then
Check2.Picture = LoadResPicture(101, vbResBitmap)
Call Etts
RichTextBox1(Text8.Text).SetFocus
End If
'Call Rtts
'Call Etts
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check2.ToolTipText = "DISABLE CONTROL TABBING WHEN REQUIRED"
Check2.MousePointer = 99
Check2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check3.ToolTipText = "TICK THE BOX TO SELECT ALL TEXT"
Check3.MousePointer = 99
Check3.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check4.ToolTipText = "TICK THE BOX TO SELECT ALL TEXT"
Check4.MousePointer = 99
Check4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub chkAware_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Index = 0 To 1
chkAware(Index).MousePointer = 99
chkAware(Index).MouseIcon = LoadResPicture(103, vbResCursor)
Next Index
End Sub

Private Sub Combo1_Click()
Call tSample
If Check3.Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).Font.Name = Combo1.Text
ElseIf Check3.Value = 0 Then
RichTextBox1(Text8.Text).SelFontName = Combo1.Text
End If
'Call tHoller
'PicLines.Cls
'Call Form_Paint
End Sub

Private Sub Combo2_Click()
If Check4.Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).Font.Size = Combo2.Text
ElseIf Check4.Value = 0 Then
'RichTextBox1.SetFocus
RichTextBox1(Text8.Text).SelFontSize = Val(Combo2.Text)
End If
'Call tHoller
'PicLines.Cls
'Call Form_Paint
End Sub

Private Sub Combo3_Click()
Text4.Text = Combo3.Text
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
frmMain.Caption = "Comprehensive Note Pad - Untitled"
If TabStrip1.Tabs.Count = 10 Then
  MsgBox UCase("Maximum Number of Documents Open" _
  & vbCrLf & "10 DOCUMENTS OPEN!")
  Exit Sub
Else
  TabStrip1.Tabs.Add (TabStrip1.Tabs.Count + 1)
   EnableAutoURLDetection RichTextBox1(Text9.Text)
Text14.Text = ""
End If
Text8.Text = TabStrip1.Tabs.Count
Call sTabStrip
If Text9.Text <> 1 Then
Command1(22).Enabled = True
End If
With RichTextBox1(Text8.Text)
.Text = ""
.SelAlignment = 0
.filename = ""
.SelFontName = "Arial"
.SelBold = False
.SelBullet = False
.SelColor = vbBlack
.SelFontSize = 10
.SelItalic = False
.SelRightIndent = False
.SelStrikeThru = False
.SelUnderline = False
.SelCharOffset = 0
End With
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SendMessage RichTextBox1(Text8.Text).hWnd, WM_CLEAR, 0&, 0&
'RichTextBox1(Text8.Text).SetFocus
Line1.X1 = 1
Line1.X2 = 1
Text9.Text = "1"
Case 1
If SFile = "" Then
Command4_Click
GoTo ok:
End If
On Error GoTo jjj:
RichTextBox1(Text8.Text).SaveFile (SFile)
GoTo ok:
jjj:
asdsd = MsgBox("There was an error in saving the file!", vbExclamation + vbOKOnly, "Error!")
ok:
frmMain.Caption = "Comprehensive Note Pad - " & SFile
Case 2

    With CommonDialog1
        .DialogTitle = "Pick a File to Open"
        .CancelError = False
        .flags = 1
        .Filter = "Txt Files (*.txt)|*.txt|RTF Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
    End With
SFile = CommonDialog1.filename
RichTextBox1(Text8.Text).LoadFile (SFile)
frmMain.Caption = "Comprehensive Note Pad - " & SFile
Case 3
PrintRTF RichTextBox1(Text8.Text), 720, 720, 720, 720 'Call PrintRTF sub
Case 4
frmEncryp.Show
Case 5
frmDecryp.Show
Case 6
Me.WindowState = vbMinimized
Case 7
frmAbout.Show
Case 8
RichTextBox1(Text8.Text).SelBullet = True
Case 9
RichTextBox1(Text8.Text).SelCharOffset = 0
Case 10
Unload Me
Case 11
If Check1(14).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelStrikeThru = True
ElseIf Check1(14).Value = 0 Then
RichTextBox1(Text8.Text).SelStrikeThru = True
End If
Case 12
Form1.Show
Case 13
frmLink.Show
Case 14
RichTextBox1(Text8.Text).SelBullet = False
Case 15
sFind
Case 16
sReplace
Case 17
frmSpellC.Show
Case 18
If Check1(17).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SelLineSpacing RichTextBox1(Text8.Text), Text4.Text
ElseIf Check1(17).Value = 0 Then
SelLineSpacing RichTextBox1(Text8.Text), Text4.Text
End If
Case 19
'remove hi lite
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelColor = vbBlack
RichTextBox1(Text8.Text).SelBold = False
RichTextBox1(Text8.Text).SelUnderline = False
Text5.Text = "word cnt"
Text7.Text = "hi lite word"
Case 20
' hi lite selected word
 HighlightWords RichTextBox1(Text8.Text), Text7.Text, Label1(Text6.Text - 1).BackColor
Case 21
'word list creator

frmWordLister.Show

Case 22
'tab killer
If Text9.Text <> 10 Then
Call rTransFile
End If
Index = Text9.Text
TabStrip1.Tabs.Remove TabStrip1.SelectedItem.Index
RichTextBox1(Text9.Text - 1).Visible = True
RichTextBox1(Text9.Text).Visible = False
TabStrip1.Tabs.Item(Text9.Text - 1).Selected = True

'buffer needed
If Text8.Text = 1 Then
Command1(22).Enabled = False
End If

Case 23 ' copies text to the check for palindrome txt box
Picture2.ZOrder 0
Clipboard.SetText RichTextBox1(Text8.Text).SelText
Text10.Text = ""
Text10.SetFocus
Text10.SelText = Clipboard.GetText()
'call palimdrme check module and display result
Label10.Caption = "Is It Palindrome = " & IsPalindrome(Text10.Text, chkAware(0).Value, chkAware(1).Value)
Case 24
Label10.Caption = "Check if word or phrase is a palindrome"
Text10.Text = "Palindrome"
Picture2.ZOrder 1
Picture2.Left = 1920
Picture2.Top = 2760
Case 25 ' number to roman numeral

Dim Temp
If Not IsNumeric(RichTextBox1(Text8.Text).SelText) Then
MsgBox "Select numeric text only !", vbExclamation
Exit Sub
End If

Temp = MsgBox("Proceed with conversion?" & vbCr & "This action can't be undone!", vbYesNoCancel + vbQuestion)

If RichTextBox1(Text8.Text).SelText > 10000000 Then
MsgBox "Sorry! " & vbCr & "The number is too large!!", vbCritical
Exit Sub
Else
If Temp = vbYes Then
RichTextBox1(Text8.Text).SelText = NumericToRoman(RichTextBox1(Text8.Text).SelText)

Else
Exit Sub
End If
End If

Exit Sub
Case 26 'number in figures to number in letters
Dim Tempa
If Not IsNumeric(RichTextBox1(Text8.Text).SelText) Then
MsgBox "Select numeric text only !", vbExclamation
Exit Sub
End If

If Len(RichTextBox1(Text8.Text).SelText) > 66 Then 'Checks if they pass the 10^66 barrier
    MsgBox "Sorry!! currently conversion upto 9.99 * 10^66 is supported!", vbInformation
Exit Sub
Else
Tempa = MsgBox("Proceed with conversion?" & vbCr & "This action can't be undone!", vbYesNoCancel + vbQuestion)
If Tempa = vbYes Then
Convert RichTextBox1(Text8.Text).SelText
Else
Exit Sub
End If
End If

Exit Sub
Case 27 'insert or copy special symbols
frmSpecSym.Show
Case 28
    Dim RTFformat As CHARFORMAT2
    With RTFformat
        .cbSize = Len(RTFformat)
        .dwMask = CFM_BACKCOLOR
        .crBackColor = vbCyan
    End With
    SendMessage RichTextBox1(Text8.Text).hWnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat

End Select
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
Case 0
Command1(Index).ToolTipText = "START A NEW FILE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 1
Command1(Index).ToolTipText = "SAVE THIS FILE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 2
Command1(Index).ToolTipText = "OPEN A SAVED FILE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 3
Command1(Index).ToolTipText = "PRINT CURRENT FILE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 4
Command1(Index).ToolTipText = "OPEN ENCRYPT(TEXT TO PICTURE)DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 5
Command1(Index).ToolTipText = "OPEN DECRYPT(PICTURE TO TEXT)DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 6
Command1(Index).ToolTipText = "MINIMISE THIS WINDOW"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 7
Command1(Index).ToolTipText = "ABOUT DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 8
Command1(Index).ToolTipText = "ADD BULLETS"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 9
Command1(Index).ToolTipText = "REMOVE SUPER & SUB SCRIPTS"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 10
Command1(Index).ToolTipText = "EXIT THIS APPLICATION, PLEASE USE THE DOOR!"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 11
Command1(Index).ToolTipText = "STRIKE THROUGH TEXT"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 12
Command1(Index).ToolTipText = "OPEN SECURE FILE WIPER DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 13
Command1(Index).ToolTipText = "WEB LINK"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 14
Command1(Index).ToolTipText = "REMOVE BULLETS"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 15
Command1(Index).ToolTipText = "OPEN FIND DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 16
Command1(Index).ToolTipText = "OPEN REPLACE DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 17
Command1(Index).ToolTipText = "OPEN SPELL CHECK DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 18
Command1(Index).ToolTipText = "CHANGE LINE SPACING"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 19
Command1(Index).ToolTipText = "REMOVE HI LITER(S)"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 20
Command1(Index).ToolTipText = "HI LITE YOUR CHOSEN WORD"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 21
Command1(Index).ToolTipText = "CREATE A WORD LIST FROM CURRENT DOCUMENT"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 22
Command1(Index).ToolTipText = "REMOVE SELECTED TAB"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 23
Command1(Index).ToolTipText = "PALINDROME CAPTURE - HI LITE & CLICK "
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 24
Command1(Index).ToolTipText = "CLEAR AND RESET PALINDROME"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 25
Command1(Index).ToolTipText = "NUMBER TO ROMAN NUMERAL"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 26
Command1(Index).ToolTipText = "NUMBER TO WORDS"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 27
Command1(Index).ToolTipText = "SHOW SPECIAL SYMBOLS DIALOGUE"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 28
Command1(Index).ToolTipText = "HIGHLIGHT SELECTED TEXT"
Command1(Index).MousePointer = 99
Command1(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Select
End Sub

Private Sub Command10_Click()
'help files shown from here when written
MsgBox UCase("help files will be here " & vbCrLf & "when written!")
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command10.ToolTipText = "SHOW HELP FILES"
Command10.MousePointer = 99
Command10.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command11_Click(Index As Integer)
Select Case Index
Case 0 ' bold
If Check1(0).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelBold = False
ElseIf Check1(0).Value = 0 Then
RichTextBox1(Text8.Text).SelBold = False
End If
Case 1
If Check1(1).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelItalic = False
ElseIf Check1(1).Value = 0 Then
RichTextBox1(Text8.Text).SelItalic = False
End If
Case 2
If Check1(1).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelUnderline = False
ElseIf Check1(1).Value = 0 Then
RichTextBox1(Text8.Text).SelUnderline = False
End If
Case 3
If Check1(14).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelStrikeThru = False
ElseIf Check1(14).Value = 0 Then
RichTextBox1(Text8.Text).SelStrikeThru = False
End If
End Select
End Sub

Private Sub Command11_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Command11(Index).ToolTipText = "UNDO BOLD TEXT"
Command11(Index).MousePointer = 99
Command11(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 1
Command11(Index).ToolTipText = "UNDO ITALIC TEXT"
Command11(Index).MousePointer = 99
Command11(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 2
Command11(Index).ToolTipText = "UNDO UNDERLINE TEXT"
Command11(Index).MousePointer = 99
Command11(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 3
Command11(Index).ToolTipText = "UNDO STRIKE THROUGH TEXT"
Command11(Index).MousePointer = 99
Command11(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Select
End Sub

Private Sub Command19_Click()
Dim gotoline As Long
            'Get pos of start of the line
            gotoline = SendMessage(RichTextBox1(Text8.Text).hWnd, EM_LINEINDEX, Text15.Text - 1, 0&)
            If gotoline = -1 Then 'Invalid line number
                MsgBox "Line number out of range", 0, "Comprehensive NotePad"
                Exit Sub
            End If
            RichTextBox1(Text8.Text).SelStart = gotoline 'Go To line
            RichTextBox1(Text8.Text).SetFocus
            DoEvents
Picture3.ZOrder 1
Picture3.Left = 1920
Picture3.Top = 3960
Text15.Text = "1"
End Sub

Private Sub Command19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command19.ToolTipText = "ENTER LINE NUMBER TO GO TO THEN CLICK ME"
Command19.MousePointer = 99
Command19.MouseIcon = LoadResPicture(101, vbResCursor)

End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0 ' bold
If Check1(0).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelBold = True
ElseIf Check1(0).Value = 0 Then
RichTextBox1(Text8.Text).SelBold = True
End If
Case 1
If Check1(1).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelItalic = True
ElseIf Check1(1).Value = 0 Then
RichTextBox1(Text8.Text).SelItalic = True
End If
Case 2
If Check1(1).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelUnderline = True
ElseIf Check1(1).Value = 0 Then
RichTextBox1(Text8.Text).SelUnderline = True
End If
Case 3 ' left
If Check1(3).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelAlignment = 0
ElseIf Check1(3).Value = 0 Then
RichTextBox1(Text8.Text).SelAlignment = 0
End If
Case 4 ' centre
If Check1(4).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelAlignment = 2
ElseIf Check1(4).Value = 0 Then
RichTextBox1(Text8.Text).SelAlignment = 2
End If
Case 5 ' right
If Check1(5).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelAlignment = 1
ElseIf Check1(5).Value = 0 Then
RichTextBox1(Text8.Text).SelAlignment = 1
End If
Case 6 ' indent
If Check1(6).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelIndent = RichTextBox1(Text8.Text).SelIndent + 850.5
ElseIf Check1(6).Value = 0 Then
RichTextBox1(Text8.Text).SelIndent = RichTextBox1(Text8.Text).SelIndent + 850.5
End If
Case 7 ' delete
If Check1(7).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SendMessage RichTextBox1(Text8.Text).hWnd, WM_CLEAR, 0&, 0&
Line1.X1 = 1
Line1.X2 = 1
RichTextBox1(Text8.Text).SetFocus
ElseIf Check1(7).Value = 0 Then
RichTextBox1(Text8.Text).SetFocus
SendMessage RichTextBox1(Text8.Text).hWnd, WM_CLEAR, 0&, 0&
'SendKeys "^{DEL}"
End If
Case 8 ' copy
If Check1(8).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SendMessage RichTextBox1(Text8.Text).hWnd, WM_COPY, 0&, 0& 'Copy
RichTextBox1(Text8.Text).SetFocus
ElseIf Check1(8).Value = 0 Then
RichTextBox1(Text8.Text).SetFocus
SendMessage RichTextBox1(Text8.Text).hWnd, WM_COPY, 0&, 0& 'Copy
'Clipboard.SetText RichTextBox1(Text8.Text).SelText
'SendKeys "^{C}"
End If
Case 9 ' paste
SendMessage RichTextBox1(Text8.Text).hWnd, WM_PASTE, 0&, 0& 'Paste
'RichTextBox1(Text8.Text).SetFocus
Case 10
frmFontColour.Show
Case 11
frmColourPicker.Show
Case 12 ' cut
If Check1(11).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SendMessage RichTextBox1(Text8.Text).hWnd, WM_CUT, 0&, 0& 'Cut
Line1.X1 = 1
Line1.X2 = 1
RichTextBox1(Text8.Text).SetFocus
ElseIf Check1(11).Value = 0 Then
RichTextBox1(Text8.Text).SetFocus
SendMessage RichTextBox1(Text8.Text).hWnd, WM_CUT, 0&, 0& 'Cut
'SendKeys "^{x}"
End If
Case 13 ' lose all formatting
'If Check1(10).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelBold = False
RichTextBox1(Text8.Text).SelItalic = False
RichTextBox1(Text8.Text).SelUnderline = False
'ElseIf Check1(10).Value = 0 Then
RichTextBox1(Text8.Text).SelBold = False
RichTextBox1(Text8.Text).SelItalic = False
RichTextBox1(Text8.Text).SelUnderline = False
RichTextBox1(Text8.Text).SelStrikeThru = False
RichTextBox1(Text8.Text).SelColor = vbBlack
RichTextBox1(Text8.Text).BackColor = vbWhite
RichTextBox1(Text8.Text).SelFontSize = 10
RichTextBox1(Text8.Text).SelFontName = "Arial"
Combo1.Text = "Arial"
Combo2.Text = "10"
Text3.ForeColor = vbBlack
Text3.BackColor = vbWhite
Text3.FontSize = 14
Text3.FontName = "Arial"
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text))
Dim RTFformat As CHARFORMAT2
    With RTFformat
        .cbSize = Len(RTFformat)
        .dwMask = CFM_BACKCOLOR
        .crBackColor = vbWhite
    End With
    SendMessage RichTextBox1(Text8.Text).hWnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
Call Command19_Click
    
'End If
Case 14 ' lowercase
If Check1(13).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
Clipboard.SetText RichTextBox1(Text8.Text).SelText
RichTextBox1(Text8.Text).SelText = LCase(Clipboard.GetText)
ElseIf Check1(13).Value = 0 Then
Clipboard.SetText RichTextBox1(Text8.Text).SelText
RichTextBox1(Text8.Text).SelText = LCase(Clipboard.GetText)
End If
Case 15 ' uppercase
If Check1(12).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
Clipboard.SetText RichTextBox1(Text8.Text).SelText
RichTextBox1(Text8.Text).SelText = UCase(Clipboard.GetText)
ElseIf Check1(12).Value = 0 Then
Clipboard.SetText RichTextBox1(Text8.Text).SelText
RichTextBox1(Text8.Text).SelText = UCase(Clipboard.GetText)
End If
Case 16 ' outdent
If Check1(15).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).SelIndent = RichTextBox1(Text8.Text).SelIndent - 850.5
ElseIf Check1(15).Value = 0 Then
RichTextBox1(Text8.Text).SelIndent = RichTextBox1(Text8.Text).SelIndent - 850.5
End If
Case 17 ' subscript
RichTextBox1(Text8.Text).SelCharOffset = -55
Case 18 ' superscript
RichTextBox1(Text8.Text).SelCharOffset = 55
Case 19
If Check1(18).Value = 1 Then
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
RichTextBox1(Text8.Text).Text = StrReverse(RichTextBox1(Text8.Text).Text)
RichTextBox1(Text8.Text).SetFocus
ElseIf Check1(18).Value = 0 Then
RichTextBox1(Text8.Text).SelText = StrReverse(RichTextBox1(Text8.Text).SelText)
RichTextBox1(Text8.Text).SetFocus
End If
Case 20
Call InsertTable(RichTextBox1(Text8.Text), Text13.Text, Text12.Text)
End Select
End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Command2(Index).ToolTipText = "BOLD TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 1
Command2(Index).ToolTipText = "ITALIC TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 2
Command2(Index).ToolTipText = "UNDERLINE TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 3
Command2(Index).ToolTipText = "LEFT ALIGN TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 4
Command2(Index).ToolTipText = "CENTRE ALIGN TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 5
Command2(Index).ToolTipText = "RIGHT ALIGN TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 6
Command2(Index).ToolTipText = "INDENT TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 7
Command2(Index).ToolTipText = "DELETE TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 8
Command2(Index).ToolTipText = "COPY TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 9
Command2(Index).ToolTipText = "PASTE TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 10
Command2(Index).ToolTipText = "CHANGE FONT / TEXT COLOUR"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 11
Command2(Index).ToolTipText = "CHANGE PAGE COLOUR"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 12
Command2(Index).ToolTipText = "CUT TEXT"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 13
Command2(Index).ToolTipText = "REMOVE TEXT AND PAGE FORMATTING"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 14
Command2(Index).ToolTipText = "CONVERT TEXT TO LOWERCASE (NOT CAPITALS)"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 15
Command2(Index).ToolTipText = "CONVERT TEXT TO UPPERCASE (CAPITALS)"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 16
Command2(Index).ToolTipText = "REMOVE TEXT INDENT (OUTDENT)"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 17
Command2(Index).ToolTipText = "CHANGE TEXT TO SUBSCRIPT (BELOW LINE)"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 18
Command2(Index).ToolTipText = "CHANGE TEXT TO SUPERSCRIPT (ABOVE LINE)"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 19
Command2(Index).ToolTipText = "REVERSE TEXT TXET ESREVER"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 20
Command2(Index).ToolTipText = "INSERT TABLE"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
Case 21
Command2(Index).ToolTipText = "CREATE ADOBE PDF FILE"
Command2(Index).MousePointer = 99
Command2(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Select
End Sub
Private Sub Command3_Click()
Dim wc As Long
    wc = WordCount()
    lblWordCount.Caption = " Word Count: " & CStr(wc)
    End Sub

Private Sub Command4_Click()
On Error GoTo ghj:
CommonDialog1.flags = 1
CommonDialog1.Filter = "Txt Files (*.txt)|*.txt|RTF Files (*.rtf)|*.rtf|All Files (*.*)|*.*|Word for Windows 6.0|*.doc"
CommonDialog1.ShowSave
If CommonDialog1.filename = "" Then GoTo h:
RichTextBox1(Text8.Text).SaveFile (CommonDialog1.filename)
SFile = CommonDialog1.filename
h:
frmMain.Caption = "Comprehensive Note Pad - " & SFile
GoTo h2:
ghj:
nul = MsgBox("Error in saving file!", vbExclamation + vbOKOnly, "Error!")
h2:
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.ToolTipText = "SAVE THIS FILE AS ..."
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(101, vbResCursor)

End Sub

Private Sub Command5_Click()
frmDTI.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.ToolTipText = "PLACE A SHORTCUT ON THE DESKTOP"
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command6_Click()
Clipboard.SetText RichTextBox1(Text8.Text).SelText
Text14.Text = ""
'Text14.SetFocus
Text14.SelText = Clipboard.GetText()
If Text14.Text <> "" Then
Call getpage
End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.ToolTipText = "HIGHLIGHT THE WEB LINK THEN CLICK ME"
Command6.MousePointer = 99
Command6.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub
Private Sub Command7_Click()
'start the basic PDF conversion
RichTextBox1(Text8.Text).SelStart = 0 'Set the start pos of the selection
RichTextBox1(Text8.Text).SelLength = Len(RichTextBox1(Text8.Text)) 'Set length of the selection
SendMessage RichTextBox1(Text8.Text).hWnd, WM_COPY, 0&, 0& 'Copy

SendMessage Text16.hWnd, WM_PASTE, 0&, 0& 'Paste into text buffer
Call Command9_Click
formConvertToPDF.Show
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.ToolTipText = "CREATE BASIC PDF FILE FROM CURRENT TAB/FILE"
Command7.MousePointer = 99
Command7.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command8_Click()
'show goto line tab
Picture3.ZOrder 0
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command8.ToolTipText = "HIGHLIGHT THE WEB LINK THEN CLICK ME"
Command8.MousePointer = 99
Command8.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Command9_Click()
' text buffer to temp file for PDF conversion
Dim data As String
Dim Filehandle As Integer
  

  Filehandle = FreeFile
data = Text16.Text
Open App.Path & "\temp.txt" For Output As #Filehandle
Print #Filehandle, data
Close #Filehandle

Text16.Text = ""
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim G As String
   Dim F   ' Declare variable.
   For F = 0 To Printer.FontCount - 1  ' Determine number of fonts.
      Combo1.AddItem Printer.Fonts(F)    ' Put each font into list box.
   Next F
   G = Combo1.ListCount
   Text2.Text = G
   frmMain.Caption = "Comprehensive Note Pad - Untitled"
   For i = 1 To 500
   Combo2.AddItem i
   Next i
   RichTextBox1(Text8.Text).Font = Combo1.Text
   RichTextBox1(Text8.Text).Font.Size = Combo2.Text
   
   EnableAutoURLDetection RichTextBox1(Text8.Text)
Text14.Text = ""
   
   Check2.Picture = LoadResPicture(101, vbResBitmap)
   
 RichTextBox1(Text8.Text).SelBold = False
RichTextBox1(Text8.Text).SelItalic = False
RichTextBox1(Text8.Text).SelUnderline = False
RichTextBox1(Text8.Text).SelStrikeThru = False
RichTextBox1(Text8.Text).SelColor = vbBlack
RichTextBox1(Text8.Text).BackColor = vbWhite

Dim r As Integer
For r = 0 To 2
Combo3.AddItem r
Next r
Text4.Text = Combo3.Text
Set TabStrip1.ImageList = ImageList1

TabStrip1.Tabs.Item(1).Selected = True
TabStrip1.Tabs.Item(1).Caption = "NEW CNP"
TabStrip1.Tabs.Item(1).Image = 1
RichTextBox1(1).Text = ""
RichTextBox1(1).Visible = True
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
     Call tRuler
     
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
Text6.Text = Index + 1
Case 1
Text6.Text = Index + 1
Case 2
Text6.Text = Index + 1
Case 3
Text6.Text = Index + 1
Case 4
Text6.Text = Index + 1
Case 5
Text6.Text = Index + 1
Case 6
Text6.Text = Index + 1
Case 7
Text6.Text = Index + 1
Case 8
Text6.Text = Index + 1
Case 9
Text6.Text = Index + 1
Case 10
Text6.Text = Index + 1
Case 11
Text6.Text = Index + 1
Case 12
Text6.Text = Index + 1
Case 13
Text6.Text = Index + 1
Case 14
Text6.Text = Index + 1
Case 15
Text6.Text = Index + 1
Case 16
Text6.Text = Index + 1
Case 17
Text6.Text = Index + 1
Case 18
Text6.Text = Index + 1
Case 19
Text6.Text = Index + 1
Case 20
Text6.Text = Index + 1
Case 21
Text6.Text = Index + 1
Case 22
Text6.Text = Index + 1
Case 23
Text6.Text = Index + 1
Case 24
Text6.Text = Index + 1
Case 25
Text6.Text = Index + 1
Case 26
Text6.Text = Index + 1
Case 27
Text6.Text = Index + 1
Case 28
Text6.Text = Index + 1
Case 29
Text6.Text = Index + 1
Case 30
Text6.Text = Index + 1
Case 31
Text6.Text = Index + 1
Case 32
Text6.Text = Index + 1
Case 33
Text6.Text = Index + 1
Case 34
Text6.Text = Index + 1
Case 35
Text6.Text = Index + 1
Case 36
Text6.Text = Index + 1
Case 37
Text6.Text = Index + 1
Case 38
Text6.Text = Index + 1
Case 39
Text6.Text = Index + 1
Case 40
Text6.Text = Index + 1
Case 41
Text6.Text = Index + 1
Case 42
Text6.Text = Index + 1
Case 43
Text6.Text = Index + 1
Case 44
Text6.Text = Index + 1
Case 45
Text6.Text = Index + 1
Case 46
Text6.Text = Index + 1
Case 47
Text6.Text = Index + 1
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Index = 0 To 48
Label1(Index).MousePointer = 99
Label1(Index).MouseIcon = LoadResPicture(102, vbResCursor)
colours = Label1(Index).BackColor
Call sVColour
Label1(Index).ToolTipText = "No: " & Index + 1 & "/ " & "R:" & reds & " G:" & greens & " B:" & blues
Next Index
End Sub
Private Sub sVColour()
blues = Int(colours / 65536)
greens = Int((colours - (65536 * blues)) / 256)
reds = colours - (blues * 65536) - (greens * 256)
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ToolTipText = "GOTO LINE TAB DISPLAYS ONLY WHEN NEEDED"
Label20.MousePointer = 99
Label20.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ToolTipText = "PALINDROME TAB DISPLAYS ONLY WHEN NEEDED"
Label9.MousePointer = 99
Label9.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.MousePointer = 0
End Sub

Private Sub RichTextBox1_Click(Index As Integer)
Text8.Text = Index
If RichTextBox1(Text8.Text).Text <> "" Then
Dim XPos As Long
    Dim YPos As Long
    
    XPos = GetTCursX
    YPos = GetTCursY
    Text1.Text = XPos
  Call tRuler
End If
Me.lbllines.Caption = " Lines: " & _
CStr(GetLineCount(Me.RichTextBox1(Text8.Text)))
Me.lblRow.Caption = " Row: " & _
CStr(GetLineNum(Me.RichTextBox1(Text8.Text)) + 1)
Me.lblCol.Caption = " Col: " & _
CStr(GetColPos(Me.RichTextBox1(Text8.Text)) + 1)
End Sub


Public Function GetTCursX() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursX = pt.X
End Function

Public Function GetTCursY() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursY = pt.Y
End Function

Private Sub tRuler()
Line1.X1 = Val(Text1.Text)
Line1.X2 = Val(Text1.Text)
End Sub

Private Sub RichTextBox1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim XPos As Long
    Dim YPos As Long
    
    XPos = GetTCursX
    YPos = GetTCursY
    '=================================
    'However, as vbTwips are the default unit of measure, you
    'can obtain
    'the position in vbTwips using the following call
    'XPos = ScaleX(GetTCursX, vbPixels, vbTwips)
    'YPos = ScaleY(GetTCursY, vbPixels, vbTwips)
    '=================================
'     Me.Caption = "X: " & XPos & " Y: " & YPos
'     Label2.Caption = "X: " & XPos & " Y: " & YPos
    Text1.Text = XPos
  Call tRuler
Me.lbllines.Caption = " Lines: " & _
CStr(GetLineCount(Me.RichTextBox1(Text8.Text)))
Me.lblRow.Caption = " Row: " & _
CStr(GetLineNum(Me.RichTextBox1(Text8.Text)) + 1)
Me.lblCol.Caption = " Col: " & _
CStr(GetColPos(Me.RichTextBox1(Text8.Text)) + 1)
End Sub
Private Sub tSample()
Text3.Font = Combo1.Text
End Sub

Private Function WordCount() As Long
    Dim txt As String
    Dim txtlen As Long
    Dim chrASCII As Integer
    Dim prevChr As Integer
    Dim ctn As Long
    Dim i As Long
    
    txt = RichTextBox1(Text8.Text).Text
    txtlen = Len(RichTextBox1(Text8.Text).Text)
    If txt = "" Then
        WordCount = 0
        Exit Function
    End If
    
    ctn = 0
    prevChr = 0
    For i = 1 To txtlen
        chrASCII = Asc(Mid(txt, i, 1))
        If chrASCII = 32 Or chrASCII = 10 Or chrASCII = 13 Then
            If prevChr <> 32 And prevChr <> 10 And prevChr <> 13 Then
                 ctn = ctn + 1
                 prevChr = chrASCII
            End If
        Else
            prevChr = chrASCII
        End If
    Next i
    If prevChr <> 32 And prevChr <> 10 And prevChr <> 13 Then
        ctn = ctn + 1       ' Add last
    End If
    WordCount = ctn
End Function
Private Sub Text10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.MousePointer = 0
End Sub

Private Sub Timer2_Timer()
If Timer2.Interval = 250 Then
Call Command3_Click
Timer2.Enabled = True
End If
End Sub
Private Function SelLineSpacing(rtbTarget As RichTextBox, SpacingRule As Long, Optional LineSpacing As Long = 20)
    ' SpacingRule
    ' Type of line spacing. To use this member, set the PFM_SPACEAFTER flag in the dwMask member. This member can be one of the following values.
    ' 0 - Single spacing. The dyLineSpacing member is ignored.
    ' 1 - One-and-a-half spacing. The dyLineSpacing member is ignored.
    ' 2 - Double spacing. The dyLineSpacing member is ignored.
    ' 3 - The dyLineSpacing member specifies the spacingfrom one line to the next, in twips. However, if dyLineSpacing specifies a value that is less than single spacing, the control displays single-spaced text.
    ' 4 - The dyLineSpacing member specifies the spacing from one line to the next, in twips. The control uses the exact spacing specified, even if dyLineSpacing specifies a value that is less than single spacing.
    ' 5 - The value of dyLineSpacing / 20 is the spacing, in lines, from one line to the next. Thus, setting dyLineSpacing to 20 produces single-spaced text, 40 is double spaced, 60 is triple spaced, and so on.

    Dim Para As PARAFORMAT2
    With Para
        .cbSize = Len(Para)
        .dwMask = PFM_LINESPACING
        .bLineSpacingRule = SpacingRule
        .dyLineSpacing = LineSpacing
    End With
    
    SendMessage rtbTarget.hWnd, EM_SETPARAFORMAT, ByVal 0&, Para
End Function
'Private Sub RichTextBox1_Change(Index As Integer)
'SelLineSpacing RichTextBox1(Text8.Text), Text4.Text
'End Sub
Private Sub sTabStrip()
If Text8.Text = 2 Then
TabStrip1.Tabs.Item(2).Caption = "NEW CNP"
TabStrip1.Tabs.Item(2).Selected = True
TabStrip1.Tabs.Item(2).Image = 1
RichTextBox1(2).Text = "New CNP Document 2"
 EnableAutoURLDetection RichTextBox1(2)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = True
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 3 Then
TabStrip1.Tabs.Item(3).Caption = "NEW CNP"
TabStrip1.Tabs.Item(3).Selected = True
TabStrip1.Tabs.Item(3).Image = 1
RichTextBox1(3).Text = "New CNP Document 3"
 EnableAutoURLDetection RichTextBox1(3)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = True
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 4 Then
TabStrip1.Tabs.Item(4).Caption = "NEW CNP"
TabStrip1.Tabs.Item(4).Selected = True
TabStrip1.Tabs.Item(4).Image = 1
RichTextBox1(4).Text = "New CNP Document 4"
 EnableAutoURLDetection RichTextBox1(4)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = True
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 5 Then
TabStrip1.Tabs.Item(5).Caption = "NEW CNP"
TabStrip1.Tabs.Item(5).Selected = True
TabStrip1.Tabs.Item(5).Image = 1
RichTextBox1(5).Text = "New CNP Document 5"
 EnableAutoURLDetection RichTextBox1(5)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = True
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 6 Then
TabStrip1.Tabs.Item(6).Caption = "NEW CNP"
TabStrip1.Tabs.Item(6).Selected = True
TabStrip1.Tabs.Item(6).Image = 1
RichTextBox1(6).Text = "New CNP Document 6"
 EnableAutoURLDetection RichTextBox1(6)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = True
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 7 Then
TabStrip1.Tabs.Item(7).Caption = "NEW CNP"
TabStrip1.Tabs.Item(7).Selected = True
TabStrip1.Tabs.Item(7).Image = 1
RichTextBox1(7).Text = "New CNP Document 7"
 EnableAutoURLDetection RichTextBox1(7)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = True
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 8 Then
TabStrip1.Tabs.Item(8).Caption = "NEW CNP"
TabStrip1.Tabs.Item(8).Selected = True
TabStrip1.Tabs.Item(8).Image = 1
RichTextBox1(8).Text = "New CNP Document 8"
 EnableAutoURLDetection RichTextBox1(8)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = True
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
End If
If Text8.Text = 9 Then
TabStrip1.Tabs.Item(9).Caption = "NEW CNP"
TabStrip1.Tabs.Item(9).Selected = True
TabStrip1.Tabs.Item(9).Image = 1
RichTextBox1(9).Text = "New CNP Document 9"
 EnableAutoURLDetection RichTextBox1(9)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = True
RichTextBox1(10).Visible = False
End If
If Text8.Text = 10 Then
TabStrip1.Tabs.Item(10).Caption = "NEW CNP"
TabStrip1.Tabs.Item(10).Selected = True
TabStrip1.Tabs.Item(10).Image = 1
RichTextBox1(10).Text = "New CNP Document 10"
 EnableAutoURLDetection RichTextBox1(10)
Text14.Text = ""
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = True
End If
End Sub

Private Sub TabStrip1_Click()

Text9.Text = TabStrip1.SelectedItem.Index
If TabStrip1.SelectedItem.Index = 1 Then
Command1(22).Enabled = False
Else
Command1(22).Enabled = True
End If

If TabStrip1.SelectedItem.Index = 1 Then
RichTextBox1(1).Visible = True
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 2 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = True
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 3 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = True
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 4 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = True
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 5 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = True
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 6 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = True
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 7 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = True
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 8 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = True
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 9 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = True
RichTextBox1(10).Visible = False
Call getline
End If
If TabStrip1.SelectedItem.Index = 10 Then
RichTextBox1(1).Visible = False
RichTextBox1(2).Visible = False
RichTextBox1(3).Visible = False
RichTextBox1(4).Visible = False
RichTextBox1(5).Visible = False
RichTextBox1(6).Visible = False
RichTextBox1(7).Visible = False
RichTextBox1(8).Visible = False
RichTextBox1(9).Visible = False
RichTextBox1(10).Visible = True
Call getline
End If
End Sub
Private Function HighlightWords(RTB As RichTextBox, sFindString As String, lColor As Long) As Integer
                            
Dim lFoundPos As Long           'Position of first character of match
Dim lFindLength As Long         'Length of string to find
Dim lOriginalSelStart As Long
Dim lOriginalSelLength As Long
Dim iMatchCount As Integer      'Number of matches
 
           'Save the insertion points current location and length
          lOriginalSelStart = RTB.SelStart
          lOriginalSelLength = RTB.SelLength
 
           'Cache the length of the string to find
          lFindLength = Len(sFindString)
 
           'Attempt to find the first match
          lFoundPos = RTB.Find(sFindString, 0, , rtfNoHighlight)
          While lFoundPos > 0
            iMatchCount = iMatchCount + 1
 
             RTB.SelStart = lFoundPos
            'The SelLength property is set to 0 as
            'soon as you change SelStart
            RTB.SelLength = lFindLength
            RTB.SelColor = lColor
            RTB.SelBold = True
            RTB.SelUnderline = True
             'Attempt to find the next match
            lFoundPos = RTB.Find(sFindString, _
              lFoundPos + lFindLength, , rtfNoHighlight)
          Wend
 
           'Restore the insertion point to its original
          'location and length
          RTB.SelStart = lOriginalSelStart
          RTB.SelLength = lOriginalSelLength
 
           'Return the number of matches
          HighlightWords = iMatchCount
  Text5.Text = iMatchCount
End Function
Sub Convert(sStr As String)
'Convert numbers to words upto "9.99 * 10^66"
Dim X As Integer
Dim sText As String
Dim T1 As Integer
Dim Bot, Top As Integer
Dim Neg, Dol As String
Dim TempChars As String
Dim LenChars As Integer
Dim Lenght As Integer

If Left(sStr, 1) = "$" Then 'Checks if it is in dollars
sStr = Right(sStr, Len(sStr) - 1)   'Removes the dollar sign
Dol = " Dollars"                    'Adds the 'Dollars' Flag
End If

If Int(sStr) < 0 Then   'Checks if the number is negative
sStr = Right(sStr, Len(sStr) - 1) 'Turns number positive
Neg = "Negative "   'Adds the 'Negative' Flag
End If

TempChars = Flip(sStr) 'Takes the number and flips it so that the ones come first

LenChars = Len(sStr) 'Finds how long the Number is
Lenght = Int(LenChars / 3 + 2 / 3) 'Calulates how many places (powers of 10)
'Returns 1 if < 1000 2 if less than 1000000...

For X = 1 To Lenght
Bot = 3 * X - 2 'Sets the bottom barrier
Top = 3 * X     'Sets the top barrier
If Top > LenChars Then Top = LenChars 'Checks that the top does not exceed the amount of charachters
'Cuts the 3 numbers,flips then so that tehy are in correct order, convers them to decimal
T1 = Int(Flip(Mid(TempChars, Bot, Top - Bot + 1))) 'Derives numbers
sText = Places(Int(T1), X) & " " & sText 'Calls function to turn nums to text
Next
sText = Trim(sText) 'Removes unnessesary spaces
sText = Neg & sText 'If negative flag, then adde 'Negative'
sText = sText & Dol 'If dollars flag, then adds 'Dollars'
On Error GoTo Hell
RichTextBox1(Text8.Text).SelText = sText 'Cuts of unneccesary spaces
If Int(sStr) = 0 Then RichTextBox1(Text8.Text).SelText = "Zero" 'If number was 0

Exit Sub
Hell:

End Sub
Function Flip(St As String) As String 'Funtion flips string 'hello' into 'olleh'
Dim X As Long
For X = Len(St) To 1 Step -1
Flip = Flip & Mid(St, X, 1)
Next
End Function
Function Places(Num As Integer, Ln As Integer) As String
Dim D1 As Integer
Dim D2 As Integer
Dim D3 As Integer
Dim Labels(23) As String
Labels(1) = ""              'Declare place identifiers
Labels(2) = "Thousand"
Labels(3) = "Million"
Labels(4) = "Billion"
Labels(5) = "Trillion"
Labels(6) = "Quadrillion"  'From this one down, I found these at a website
Labels(7) = "Quintillion"
Labels(8) = "Sextillion"
Labels(9) = "Septillion"
Labels(10) = "Octillion"
Labels(11) = "Nonillion"
Labels(12) = "Decillion"
Labels(13) = "Undecillion"
Labels(14) = "Duodecillion"
Labels(15) = "Tredecillion"
Labels(16) = "Quatuordecillion"
Labels(17) = "Quindecillion"
Labels(18) = "Sexdecillion"
Labels(19) = "Septendecillion"
Labels(20) = "Octodecillion"
Labels(21) = "Novemdecillion"
Labels(22) = "Vigintillion"
D3 = Num Mod 10                 'Calculates the ones place
D2 = ((Num Mod 100) - D3) / 10  'Calculates the tens place
D1 = (Num - D2 * 10 - D3) / 100 'Calculates the hundreds place

If D1 <> 0 Then Places = Places & TT(D1) & " Hundred"  'Convers hudreds place to text
If D2 <> 0 Then
If D2 = 1 Then  'If the number is between 10-19
Places = Places & " " & TT(D2 * 10 + D3) 'So that it is 'Nineteen' instead of 'Ten Nine'
D3 = 0  'Turns ones place into zero so its doesnt print twice 'Nineteen Nine'
Else  'If the number is not 10-19
Places = Places & " " & TT(D2 * 10) 'Does tens seperately 'Twenty'
Places = Places & " " & TT(D3)      'Does ones seperately 'Nine'
D3 = 0
End If
End If
If D3 <> 0 Then Places = Places & " " & TT(D3) 'If Tens were 0 it prints the ones place
Places = Places & " " & Labels(Ln) 'Add the label (Thousand, Million, ect..)
End Function
Function TT(Num As Integer) As String
Select Case Num 'A string for every special number (1-19, and 20, 30, ..., 90 but not 0)
Case 1
TT = "One"
Case 2
TT = "Two"
Case 3
TT = "Three"
Case 4
TT = "Four"
Case 5
TT = "Five"
Case 6
TT = "Six"
Case 7
TT = "Seven"
Case 8
TT = "Eight"
Case 9
TT = "Nine"
Case 10
TT = "Ten"
Case 11
TT = "Eleven"
Case 12
TT = "Twelve"
Case 13
TT = "Thirteen"
Case 14
TT = "Fourteen"
Case 15
TT = "Fifteen"
Case 16
TT = "Sixteen"
Case 17
TT = "Seventeen"
Case 18
TT = "Eighteen"
Case 19
TT = "Nineteen"
Case 20
TT = "Twenty"
Case 30
TT = "Thirty"
Case 40
TT = "Forty"
Case 50
TT = "Fifty"
Case 60
TT = "Sixty"
Case 70
TT = "Seventy"
Case 80
TT = "Eighty"
Case 90
TT = "Ninety"
End Select

End Function
Private Sub rTransFile()
rtbBuffer.Text = RichTextBox1(Text9.Text + 1).Text
RichTextBox1(Text9.Text).Text = ""
RichTextBox1(Text9.Text).Text = rtbBuffer.Text
rtbBuffer.Text = ""
End Sub
'*******************
' GetLineCount
' Get the line count
'*******************
 Function GetLineCount(tBox As Object) As Long
GetLineCount = SendMessageByNum(tBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function
'*******************
' GetLineNum
' Get current line number
'*******************
Function GetLineNum(tBox As Object) As Long
GetLineNum = SendMessageByNum(tBox.hWnd, EM_LINEFROMCHAR, tBox.SelStart, 0&)
End Function
'*******************
' GetColPos
' Get current Column
'*******************
 Function GetColPos(tBox As Object) As Long
GetColPos = tBox.SelStart - SendMessageByNum(tBox.hWnd, EM_LINEINDEX, -1&, 0&)
End Function
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Picture2.Move Picture2.Left + (X - CurX), Picture2.Top + (Y - CurY)
Picture2.MousePointer = 99
Picture2.MouseIcon = LoadResPicture(105, vbResCursor)
End If
Picture2.ToolTipText = "RIGHT CLICK AND HOLD TO DRAG"
End Sub
Public Sub InsertTable(RTB As RichTextBox, vRows As Integer, vCols As Integer)
Dim A As String, i As Integer, j As Integer

A = "{\rtf1\ansi\ansicpg1252\deff0" & _
"{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}}" & _
"\viewkind4\uc1\trowd\trqc\trgaph108\trleft-8" & _
"\trbrdrt\brdrs\brdrw10" & _
"\trbrdrl\brdrs\brdrw10" & _
"\trbrdrb\brdrs\brdrw10" & _
"\trbrdrr\brdrs\brdrw10"
For i = 1 To vCols
A = A & "\clbrdrt\brdrw15\brdrs" & _
"\clbrdrl\brdrw15\brdrs" & _
"\clbrdrb\brdrw15\brdrs" & _
"\clbrdrr\brdrw15\brdrs" & _
"\cellx" & _
CStr((ScaleX(RichTextBox1(Text8.Text).Width, RichTextBox1(Text8.Text).Parent.ScaleMode, vbTwips) \ vCols * Val(Text11.Text)) * i) & _
"\clbrdrt"
Next
A = A & "\pard\intbl\lang3082\f0\fs24"
For i = 1 To vRows
A = A & "\intbl\clmrg"
For j = 1 To vCols
A = A & "\cell"
Next
A = A & "\row"
Next
A = A & "}"
RichTextBox1(Text8.Text).SelText = A

End Sub
Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
Dim pt As POINTAPI
Dim pos As Integer
Dim start_pos As Integer
Dim end_pos As Integer
Dim ch As String
Dim txt As String
Dim txtlen As Integer

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
    
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "!" And ch <= "z") Or _
            ch = "_" _
        ) Then Exit For
        
    Next start_pos
    
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
    
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "!" And ch <= "z") Or _
            ch = "_" _
        ) Then Exit For
        
    Next end_pos
    
    end_pos = end_pos - 1

    If start_pos <= end_pos Then _
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
        
End Function

Private Sub getpage()

Dim lWindow As Long
    Call ShellExecute(lWindow, "open", Text14.Text, vbNullString, vbNullString, 5)
    Text14.Text = ""
End Sub
Public Sub EnableAutoURLDetection(RTB As RichTextBox)

    'enable auto URL detection
    
    SendMessage RTB.hWnd, EM_AUTOURLDETECT, 1&, ByVal 0&
  
End Sub
Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Picture3.Move Picture3.Left + (X - CurX), Picture3.Top + (Y - CurY)
Picture3.MousePointer = 99
Picture3.MouseIcon = LoadResPicture(105, vbResCursor)
End If
Picture3.ToolTipText = "RIGHT CLICK AND HOLD TO DRAG"
End Sub
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.MousePointer = 0
End Sub
Private Sub getline()
Dim CurrentLine As Long
    CurrentLine = SendMessage(RichTextBox1(Text8.Text).hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    Text15.Text = Trim(Str$(CurrentLine))
    Text15.SelStart = 0
    Text15.SelLength = Len(Text15.Text)
End Sub
Private Sub Rtts()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
' remove the tabstops
 For i = 1 To 10
 RichTextBox1(i).TabStop = False
 Next i

 For j = 0 To 28
 Command1(j).TabStop = False
 Next j

 For k = 0 To 21
 Command2(k).TabStop = False
 Next k

 For l = 0 To 18
 Check1(l).TabStop = False
 Next l

For m = 0 To 3
 Command11(m).TabStop = False
 Next m
 
 Picture1.TabStop = False
  chkAware(0).TabStop = False: Picture2.TabStop = False
   Command3.TabStop = False: chkAware(1).TabStop = False: Picture3.TabStop = False
    Command4.TabStop = False: Text1.TabStop = False: Text12.TabStop = False
     Command5.TabStop = False: Text2.TabStop = False: Text13.TabStop = False
      Command6.TabStop = False: Text3.TabStop = False: Text14.TabStop = False
       Command7.TabStop = False: Text4.TabStop = False: Text15.TabStop = False
        Command8.TabStop = False: Text5.TabStop = False: Text16.TabStop = False
         Command9.TabStop = False: Text6.TabStop = False
           Text7.TabStop = False: Check2.TabStop = False
           Text8.TabStop = False: Command10.TabStop = False
             Text9.TabStop = False: Check3.TabStop = False
              Text10.TabStop = False: Check4.TabStop = False
               Text11.TabStop = False
                  rtbBuffer.TabStop = False
                   TabStrip1.TabStop = False
                   Command19.TabStop = False: Combo1.TabStop = False
                     Combo2.TabStop = False
                      Combo3.TabStop = False
End Sub
Private Sub Etts()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
' remove the tabstops
 For i = 1 To 10
 RichTextBox1(i).TabStop = True
 Next i

 For j = 0 To 28
 Command1(j).TabStop = True
 Next j

 For k = 0 To 21
 Command2(k).TabStop = True
 Next k

 For l = 0 To 18
 Check1(l).TabStop = True
 Next l

 For m = 0 To 3
 Command11(m).TabStop = True
 Next m
 
 Picture1.TabStop = True
  chkAware(0).TabStop = True: Picture2.TabStop = True
   Command3.TabStop = True: chkAware(1).TabStop = True: Picture3.TabStop = True
    Command4.TabStop = True: Text1.TabStop = True: Text12.TabStop = True
     Command5.TabStop = True: Text2.TabStop = True: Text13.TabStop = True
      Command6.TabStop = True: Text3.TabStop = True: Text14.TabStop = True
       Command7.TabStop = True: Text4.TabStop = True: Text15.TabStop = True
        Command8.TabStop = True: Text5.TabStop = True: Text16.TabStop = True
         Command9.TabStop = True: Text6.TabStop = True
           Text7.TabStop = True: Check2.TabStop = True
           Text8.TabStop = True: Command10.TabStop = True
             Text9.TabStop = True: Check3.TabStop = True
              Text10.TabStop = True: Check4.TabStop = True
               Text11.TabStop = True
                  rtbBuffer.TabStop = True
                   TabStrip1.TabStop = True
                   Command19.TabStop = True: Combo1.TabStop = True
                     Combo2.TabStop = True
                      Combo3.TabStop = True
End Sub

