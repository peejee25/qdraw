VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PG Draw 1.0. -  Attributes"
   ClientHeight    =   3525
   ClientLeft      =   4425
   ClientTop       =   2175
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3885
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Units"
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "Twips"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pixel"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Size"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox text2 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "400"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Text            =   "400"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Height: "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Width: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   495
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
   Dim ht As Integer
   Dim wdt As Integer
   Dim mult As Double
   If Option1.Value = True Then
        mult = 1
    Else
        mult = 1 / 15
   End If
   
   If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
    If Text1.Text > 0 And Text2.Text > 0 Then
       ht = CInt(Text2.Text) * mult
       wdt = CInt(Text1.Text) * mult
       If Form1.Picture2.Width <> wdt Or Form1.Picture2.Height <> ht Then
        Form1.Picture2.Width = wdt
        Form1.Picture2.Height = ht
       
        Form1.Picture3.Height = ht
        Form1.Picture3.Width = wdt
        
        Form1.Picture1.Width = wdt
        Form1.Picture1.Height = ht
       
        Form1.Picture3.PaintPicture Form1.Picture2.Image, 0, 0, Form1.Picture3.Width, Form1.Picture3.Height, 0, 0, Form1.Picture2.Width, Form1.Picture2.Height
        Form1.Picture2.Picture = LoadPicture()
        Form1.Picture2.PaintPicture Form1.Picture3.Image, 0, 0, Form1.Picture2.Width, Form1.Picture2.Height, 0, 0, Form1.Picture3.Width, Form1.Picture3.Height
        Form1.Picture3.Picture = LoadPicture()
        
        Form1.Picture1.Width = wdt
        Form1.Picture1.Height = ht
       End If
    End If
  End If
  Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Text1.Text = Form1.Picture2.Width
    Text2.Text = Form1.Picture2.Height
End Sub


Private Sub Option1_Click()
Text1.Text = CInt(Text1.Text) / 15
Text2.Text = CInt(Text2.Text) / 15
End Sub

Private Sub Option2_Click()
Text1.Text = CInt(Text1.Text) * 15
Text2.Text = CInt(Text2.Text) * 15
End Sub
