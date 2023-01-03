VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PG Draw   Flip/Rotate Image"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   480
      ScaleHeight     =   1935
      ScaleWidth      =   2055
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3428
      TabIndex        =   6
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   1268
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flip/Rotate"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "Both"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Vertical"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Horizontal"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   5400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   3120
      MousePointer    =   1  'Arrow
      Stretch         =   -1  'True
      ToolTipText     =   "Preview"
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public flipdirection As String

Private Sub Command1_Click()
    If Option1.Value = True Then
        Picture1.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.Width, Picture1.Height, Picture1.Width, _
        0, -Picture1.Width, Picture1.Height, &HCC0020
        
    ElseIf Option2.Value = True Then
        Picture1.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.Width, Picture1.Height, 0, _
        Picture1.Height, Picture1.Width, -Picture1.Height, &HCC0020
        
    ElseIf Option3.Value = True Then
        Picture1.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.Width, Picture1.Height, Picture1.Width, _
        Picture1.Height, -Picture1.Width, -Picture1.Height, &HCC0020
        
    End If
    Image1.Picture = Picture1.Image
    Picture1.Picture = Picture1.Image
        
End Sub

Private Sub Command2_Click()
    Form1.Picture2.Picture = Picture1.Image
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Image1.Picture = Form1.Picture2.Image
    Picture1.Picture = Image1.Picture
End Sub
