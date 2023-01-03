VERSION 5.00
Begin VB.Form Form_con_param 
   Caption         =   "Specify Parameters"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
   LinkTopic       =   "Form7"
   ScaleHeight     =   7170
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Text            =   "0.5"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Text            =   "1"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Text            =   "0.5"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Text            =   "160"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Text            =   "80"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Slope "
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Line3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Slope "
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Line2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Slope "
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Cut Off Point X1 - value (X0< X1< =255)"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Cut Off Point X0 - value (0<=X0<=255)"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Line1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "Form_con_param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   
   cx0 = CInt(Text1.Text)
   cx1 = CInt(Text2.Text)
   cm1 = CDbl(Text3.Text)
   cm2 = CDbl(Text4.Text)
   cm3 = CDbl(Text5.Text)
   cc2 = cm1 * cx0 - cm2 * cx0
   cc3 = (cm2 * cx1 + cc2) - (cm3 * cx1)
   
     
   If cx0 < 0 Or cx0 > 255 Then
      MsgBox ("Error: Value should be 0 <= X0 <= 255")
   End If
   
   If cx1 > 255 Or cx1 < x0 Then
        MsgBox ("Error: Value should be X0 <= X1 <= 255")
   End If
   
  ' If (cm1 * cx0) > 255 Or (cm1 * cx0) < 0 Then
   '     MsgBox ("Error: Please change slope of line1, Intensity value after transformation out of range")
 '  End If
   
 '  If (cm2 * cx1 + cc2) > 255 Or (cm2 * cx1 + cc2) < 0 Then
  '      MsgBox ("Error: Please change slope of Line2, Intensity value after transformation out of range")
  ' End If
    
  ' If (cm3 * 255 + cc3) > 255 Or (cm3 * 255 + cc3) < 0 Then
  '      MsgBox ("Error: Please change the slope of line3, Intensity value after transformation out of range")
  ' End If
   
   Me.Hide
   
   
   Call Form6.contrast_image
   
   
   
   
End Sub
