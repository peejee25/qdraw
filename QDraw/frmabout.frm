VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About PG Draw 1.0. "
   ClientHeight    =   3525
   ClientLeft      =   4215
   ClientTop       =   2280
   ClientWidth     =   4680
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3525
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1973
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label5 
      Caption         =   "Developed by:  Prashant Ganesh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   533
      TabIndex        =   4
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Distribution:  Freeware"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Application Type:  Paint && Image Processing"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Version:  1.0.0."
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "PG Draw 1.0."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub
