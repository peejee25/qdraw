VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "Custom Filter"
   ClientHeight    =   3030
   ClientLeft      =   3945
   ClientTop       =   5925
   ClientWidth     =   4110
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4110
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -180
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17485e-38
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      TabIndex        =   33
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Now"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      TabIndex        =   32
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      TabIndex        =   29
      Text            =   "0"
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Text            =   "9"
      Top             =   2505
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   1560
      TabIndex        =   24
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   1200
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   840
      TabIndex        =   22
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   480
      TabIndex        =   21
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   1560
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   1200
      TabIndex        =   18
      Text            =   "1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   840
      TabIndex        =   17
      Text            =   "1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   480
      TabIndex        =   16
      Text            =   "1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1560
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1200
      TabIndex        =   13
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   840
      TabIndex        =   12
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   840
      TabIndex        =   7
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2220
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   480
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "5 X 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 X 3 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Filter Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Bias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   31
      Top             =   2235
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Divide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   30
      Top             =   2235
      Width           =   780
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileSave 
         Caption         =   "Save Filter"
      End
      Begin VB.Menu FileLoad 
         Caption         =   "Load Filter"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FileLoad_Click()
On Error GoTo ExitSub
    CommonDialog1.Filter = "Custom Filters|*.FLR"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    FFileNum = FreeFile
    Me.Caption = CommonDialog1.FileTitle
    Open CommonDialog1.FileName For Input As FFileNum
    Input #FFileNum, FSize
    If FSize = 3 Then Option1.Value = True Else Option2.Value = True
    Input #FFileNum, div
    Input #FFileNum, bias
    Text2.Text = div
    Text3.Text = bias
    For i = 0 To 4
        For j = 0 To 4
            Input #FFileNum, fweight
            Text1(i * 5 + j).Text = fweight
        Next
    Next
    Close #FFileNum
    Exit Sub
    
ExitSub:
    Exit Sub
End Sub

Private Sub FileSave_Click()
On Error GoTo ExitSub
    CommonDialog1.Filter = "Custom Filters|*.FLR"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowSave
    Me.Caption = CommonDialog1.FileTitle
    FFileNum = FreeFile
    Open CommonDialog1.FileName For Output As FFileNum
    If Option1.Value Then Write #FFileNum, "3" Else Write #FFileNum, "5"
    Write #FFileNum, Text2.Text
    Write #FFileNum, Text3.Text
    For i = 0 To 4
        For j = 0 To 4
            Write #FFileNum, Text1(i * 5 + j).Text
        Next
    Next
    Close #FFileNum
    Exit Sub
    
ExitSub:
    Exit Sub
End Sub

Private Sub Form_Load()
    Option1_Click
End Sub

Private Sub Option1_Click()
Dim i As Integer
    For i = 0 To 4
        Text1(i).Visible = False
        Text1(i + 20).Visible = False
    Next
    For i = 1 To 3
        Text1(i * 5).Visible = False
        Text1(i * 5 + 4).Visible = False
    Next
End Sub

Private Sub Option2_Click()
Dim i As Integer
    For i = 0 To 4
        Text1(i).Visible = True
        Text1(i + 20).Visible = True
    Next
    For i = 1 To 3
        Text1(i * 5).Visible = True
        Text1(i * 5 + 4).Visible = True
    Next
End Sub


Private Sub Command1_Click()
Dim i, j
    FilterCancel = False
    For i = 0 To 4
        For j = 0 To 4
            CustomFilter(i, j) = Val(Text1(i * 5 + j).Text)
        Next
    Next
   
    FilterNorm = Val(Text2.Text)
    FilterBias = Val(Text3.Text)
    Form5.Hide

End Sub

Private Sub Command2_Click()
    FilterCancel = True
    Form5.Hide
End Sub

