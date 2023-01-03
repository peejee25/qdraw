VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "PG Draw 1.0."
   ClientHeight    =   4395
   ClientLeft      =   2820
   ClientTop       =   2535
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   PaletteMode     =   2  'Custom
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar3 
      Align           =   3  'Align Left
      Height          =   3300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   5821
      ButtonWidth     =   741
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "ImageList3"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Freehand"
            Description     =   "Freehand Drawing "
            Object.ToolTipText     =   "Freehand"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Line"
            Description     =   "Draws a Straight Line"
            Object.ToolTipText     =   "Line"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Draws a Circle"
            Object.ToolTipText     =   "Circle"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Draws a Rectangle"
            Object.ToolTipText     =   "Box"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Draws Polygon"
            Object.ToolTipText     =   "Polygon"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Eraser"
            Object.ToolTipText     =   "Eraser"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Inserts Text"
            Object.ToolTipText     =   "Text"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Fills a Region"
            Object.ToolTipText     =   "FloodFill"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Draws a Bezier Curve"
            Object.ToolTipText     =   "Curve"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Description     =   "Spray Paints over specified region"
            Object.ToolTipText     =   "AirBrush"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
   End
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   3300
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Black"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "White"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Silver"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Gray"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Maroon"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Red"
            Object.Tag             =   ""
            ImageIndex      =   16
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Olive"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Yellow"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Green"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Lime"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Teal"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Aqua"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Navy"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Blue"
            Object.Tag             =   ""
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Purple"
            Object.Tag             =   ""
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Fulchusia"
            Object.Tag             =   ""
            ImageIndex      =   15
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pix1"
            Object.ToolTipText     =   "1 Pixel"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pix2"
            Object.ToolTipText     =   "2 Pixel"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "pix3"
            Object.ToolTipText     =   "3 Pixel"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4140
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2716
            MinWidth        =   2716
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   315
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   3.62669e-37
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   600
      MouseIcon       =   "qdraw.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3600
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   60
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":06E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":10E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":15EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":1B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":2176
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":2678
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":2C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":314C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":364E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":39A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":3CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":4044
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":4396
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":46E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":4A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":4D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":50DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":5430
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":5782
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":5AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":5E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":6178
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":64CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":681C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":6B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":6EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "qdraw.frx":7212
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu FileNew 
         Caption         =   "New"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu FilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu fileseperator 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      NegotiatePosition=   3  'Right
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu toolbars 
         Caption         =   "&Toolbars"
         Begin VB.Menu toolcolor 
            Caption         =   "&Color Box"
         End
         Begin VB.Menu toolwidth 
            Caption         =   "&Width Box"
         End
         Begin VB.Menu tooldrawbox 
            Caption         =   "&Draw Box"
         End
      End
   End
   Begin VB.Menu Toolmenu 
      Caption         =   "Tools"
      Begin VB.Menu ShapeMenu 
         Caption         =   "Shape"
         Begin VB.Menu DrawFreehand 
            Caption         =   "Freehand"
         End
         Begin VB.Menu DrawLine 
            Caption         =   "Line"
         End
         Begin VB.Menu DrawCircle 
            Caption         =   "Circle"
         End
         Begin VB.Menu DrawBox 
            Caption         =   "Box"
         End
         Begin VB.Menu DrawText 
            Caption         =   "Text"
         End
         Begin VB.Menu DrawPoly 
            Caption         =   "Polygon"
         End
         Begin VB.Menu DrawCurve 
            Caption         =   "Curve"
         End
         Begin VB.Menu DrawAirBrush 
            Caption         =   "AirBrush"
         End
      End
      Begin VB.Menu WidthMenu 
         Caption         =   "Width"
         Begin VB.Menu width1 
            Caption         =   "1 pixel"
            Checked         =   -1  'True
         End
         Begin VB.Menu Width2 
            Caption         =   "2 pixels"
         End
         Begin VB.Menu Width3 
            Caption         =   "3 pixels"
         End
      End
      Begin VB.Menu StyleMenu 
         Caption         =   "DrawStyle"
         Begin VB.Menu StyleSolid 
            Caption         =   "Solid"
            Checked         =   -1  'True
         End
         Begin VB.Menu StyleDash 
            Caption         =   "Dash"
         End
         Begin VB.Menu StyleDot 
            Caption         =   "Dot"
         End
      End
      Begin VB.Menu ColorMenu 
         Caption         =   "More Colors..."
      End
   End
   Begin VB.Menu imagemenu 
      Caption         =   "Image"
      Begin VB.Menu ImageAttrib 
         Caption         =   "Attributes"
      End
      Begin VB.Menu imageflip 
         Caption         =   "Flip/Rotate"
      End
      Begin VB.Menu ProcessMenu 
         Caption         =   "Process"
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "Help"
      Begin VB.Menu helpabout 
         Caption         =   "About PG Draw..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shape As String
Dim drawstring As String
Dim XStart, YStart, XPrevious, YPrevious As Single
Dim CopyBMP, PasteBMP, CutBMP, PrintText As Integer
Dim PDrawWidth, PDrawStyle, PFillStyle As Integer
Dim PPencolor As Long
Dim reset As Boolean
Dim CopyWidth, CopyHeight As Integer
Dim XLabel, YLabel As Integer
Dim OpenFile As String
Dim xold As Single
Dim yold As Single
Dim polydrawn As Boolean
Dim dragmode As Boolean
Dim dragdirection As Integer
Dim bezier_stepthru As Integer
Dim p0 As POINTAPI, p1 As POINTAPI, p2 As POINTAPI, p3 As POINTAPI
Dim u As Double
Dim curvex As Double, curvey As Double
Dim airbrushx As Integer, airbrushy As Integer
Dim counti As Integer

   
'Uncheck all the items in Style menu
Private Sub UnCheckStyles()
    
    StyleSolid.Checked = False
    StyleDash.Checked = False
    StyleDot.Checked = False
End Sub

'Uncheck all the items in Draw menu
Private Sub uncheckdraw()
    DrawBox.Checked = False
    DrawLine.Checked = False
    DrawCircle.Checked = False
    DrawText.Checked = False
    DrawFreehand.Checked = False
    DrawPoly.Checked = False
    DrawCurve.Checked = False
    DrawAirBrush.Checked = False
End Sub

'Uncheck all the items in Width menu
Private Sub uncheckwidth()
    width1.Checked = False
    Width2.Checked = False
    Width3.Checked = False
End Sub

Private Sub ColorMenu_Click()
    On Error GoTo cancel_operation
    CommonDialog1.CancelError = True
    CommonDialog1.Color = Picture2.ForeColor
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Picture2.ForeColor = CommonDialog1.Color
cancel_operation:
End Sub

Private Sub DrawAirBrush_Click()
    set_drawmode ("AirBrush")
End Sub

Private Sub DrawCurve_Click()
    set_drawmode ("Curve")
End Sub

Private Sub DrawFreehand_Click()
    set_drawmode ("Freehand")
End Sub

Private Sub DrawPoly_Click()
    set_drawmode ("Polygon")
End Sub

'Print the Canvas Image if printer is present else
'report the error

Private Sub FilePrint_Click()
    On Error GoTo NoPrinter
    
    'Take a new page
    Printer.NewPage
    CommonDialog1.CancelError = True
    
    'Show the printer common dialog box to set
    'properties like no. of copies etc.
    CommonDialog1.ShowPrinter
    Printer.Copies = CommonDialog1.Copies
    'Printer.Orientation = CommonDialog1.Orientation
    
    'Start printing
    Printer.PaintPicture Picture2.Image, 0, 0, Picture2.Width, Picture2.Height, _
    0, 0, Picture2.Width, Picture2.Height
    Exit Sub
NoPrinter:
If Err.Number = 32755 Then
MsgBox ("Printing Canceled.")
Exit Sub
End If
MsgBox ("Error: No Printer installed or Printer not connected. Operation Aborted.")
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not dragmode Then
    If Button = 1 Then
        If Abs(x - (Picture2.Width + Picture2.Left)) <= 20 And Abs(y - (Picture2.Height + Picture2.Top)) <= 20 Then
            Form1.MousePointer = 8
            dragdirection = 1
            Picture2.BorderStyle = 1
            dragmode = True
        
        ElseIf Abs(x - (Picture2.Width + Picture2.Left)) <= 15 Then
            Form1.MousePointer = 9
            dragdirection = 2
            Picture2.BorderStyle = 1
            dragmode = True
        
        ElseIf Abs(y - (Picture2.Height + Picture2.Top)) <= 15 Then
            Form1.MousePointer = 7
            dragdirection = 3
            Picture2.BorderStyle = 1
            dragmode = True
        
        End If
        
    End If
    
End If

    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo last1
    If Not drawmode Then
    If Abs(x - (Picture2.Width + Picture2.Left)) <= 20 And Abs(y - (Picture2.Height + Picture2.Top)) <= 20 Then
            Form1.MousePointer = 8
            
        ElseIf Abs(x - (Picture2.Width + Picture2.Left)) <= 15 Then
            Form1.MousePointer = 9
            
        ElseIf Abs(y - (Picture2.Height + Picture2.Top)) <= 15 Then
            Form1.MousePointer = 7
            
        Else
            Form1.MousePointer = 1
        End If
    End If
    
    If Button = 1 And dragmode = True Then
        Select Case dragdirection
            Case 1:
                Picture2.Width = x - Picture2.Left
                Picture2.Height = y - Picture2.Top
                
                Picture3.Width = Picture2.Width
                Picture3.Height = Picture2.Height
                
                Picture1.Width = Picture2.Width
                Picture1.Height = Picture2.Height
                
            Case 2:
                Picture2.Width = x - Picture2.Left
                Picture3.Width = Picture2.Width
                Picture1.Width = Picture2.Width
                
            Case 3:
                Picture2.Height = y - Picture2.Top
                Picture3.Height = Picture2.Height
                Picture1.Height = Picture2.Height
                
        End Select
        
        
    End If
last1:
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If dragmode Then
    Form1.MousePointer = 1
    Picture2.BorderStyle = 0
    dragmode = False
    
    Picture3.PaintPicture Form1.Picture2.Image, 0, 0, Form1.Picture3.Width, Form1.Picture3.Height, 0, 0, Form1.Picture2.Width, Form1.Picture2.Height
    Picture2.Picture = LoadPicture()
    Picture2.PaintPicture Form1.Picture3.Image, 0, 0, Form1.Picture2.Width, Form1.Picture2.Height, 0, 0, Form1.Picture3.Width, Form1.Picture3.Height
    Picture3.Picture = LoadPicture()
    
    
    End If
End Sub

Private Sub helpabout_Click()
    frmabout.Show 1
End Sub

Private Sub ImageAttrib_Click()
    Form2.Show
End Sub

Private Sub imageflip_Click()
    Form3.Show 1
End Sub

Private Sub picture2_DblClick()
    If shape = "POLYGON" And polydrawn = False Then
        Picture2.Refresh
        Picture2.AutoRedraw = True
        Picture2.Line (xold, yold)-(XStart, YStart)
        polydrawn = True
    End If
End Sub


Private Sub ProcessMenu_Click()
    Form6.Show
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.ToolTipText
        Case "Black":
            Picture2.ForeColor = RGB(0, 0, 0)
        Case "White":
            Picture2.ForeColor = RGB(255, 255, 255)
        Case "Silver":
            Picture2.ForeColor = RGB(192, 192, 192)
        Case "Gray":
            Picture2.ForeColor = RGB(128, 128, 128)
        Case "Aqua":
            Picture2.ForeColor = RGB(0, 255, 255)
        Case "Fulchusia":
            Picture2.ForeColor = RGB(255, 0, 255)
        Case "Lime":
            Picture2.ForeColor = RGB(0, 255, 0)
        Case "Teal":
            Picture2.ForeColor = RGB(0, 128, 128)
        Case "Purple":
            Picture2.ForeColor = RGB(128, 0, 128)
        Case "Navy":
            Picture2.ForeColor = RGB(0, 0, 128)
        Case "Blue":
            Picture2.ForeColor = RGB(0, 0, 255)
        Case "Red"
            Picture2.ForeColor = RGB(255, 0, 0)
        Case "Maroon"
            Picture2.ForeColor = RGB(128, 0, 0)
        Case "Green"
            Picture2.ForeColor = RGB(0, 128, 0)
        Case "Yellow"
            Picture2.ForeColor = RGB(255, 255, 0)
        Case "Olive"
            Picture2.ForeColor = RGB(128, 128, 0)
            
    End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As ComctlLib.Button)
    set_drawmode (Button.ToolTipText)
End Sub



Private Sub toolcolor_Click()
    toolcolor.Checked = Not toolcolor.Checked
    If toolcolor.Checked Then
        Toolbar2.Visible = True
    Else
        Toolbar2.Visible = False
    End If
End Sub

Private Sub tooldrawbox_Click()
    tooldrawbox.Checked = Not tooldrawbox.Checked
    If tooldrawbox.Checked Then
        Toolbar3.Visible = True
    Else
        Toolbar3.Visible = False
    End If
End Sub

Private Sub toolwidth_Click()
    toolwidth.Checked = Not toolwidth.Checked
    If toolwidth.Checked Then
        Toolbar1.Visible = True
    Else
        Toolbar1.Visible = False
    End If
End Sub


Private Sub DrawBox_Click()
    set_drawmode ("Box")
End Sub

Private Sub DrawCircle_Click()
    set_drawmode ("Circle")
End Sub

Private Sub DrawLine_Click()
    set_drawmode ("Line")
End Sub

Private Sub DrawText_Click()
    set_drawmode ("Text")
End Sub

Private Sub EditClear_Click()
    Picture2.Picture = LoadPicture()
End Sub

Private Sub EditCopy_Click()
    CopyBMP = True
End Sub

Private Sub EditCut_Click()
    CutBMP = True
End Sub

Private Sub EditPaste_Click()
    PasteBMP = True
    If Clipboard.GetFormat(vbCFBitmap) Then
          Picture1.AutoSize = True
          Picture1.Picture = Clipboard.GetData()
          Picture1.AutoSize = False
          CopyWidth = Picture1.Width
          CopyHeight = Picture1.Height
    End If
End Sub

Private Sub FileExit_Click()
    End
End Sub

Private Sub FileNew_Click()
    Picture2.Picture = LoadPicture()
    OpenFile = ""
End Sub

Private Sub FileOpen_Click()
    CommonDialog1.Filter = "Images|*.bmp;*.gif;*.jpg"
    CommonDialog1.DefaultExt = "BMP"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    Picture2.AutoSize = True
    Picture2.Picture = LoadPicture(CommonDialog1.FileName, , , 2000, 2000)
    Picture2.AutoSize = False
    OpenFile = CommonDialog1.FileName
    Picture1.AutoSize = True
    Picture1.Picture = Picture2.Picture
    Picture1.AutoSize = False
End Sub

Private Sub FileSave_Click()
    If OpenFile <> "" Then
        SavePicture Image, OpenFile
    End If
End Sub

'If cancel is clicked file should not be saved
'correct this error
'Status: Not corrected yet !
Private Sub FileSaveAs_Click()
    On Error GoTo cancel_operation
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Images|*.bmp"
    CommonDialog1.DefaultExt = "BMP"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    SavePicture Picture2.Image, CommonDialog1.FileName
    OpenFile = CommonDialog1.FileName
    Exit Sub
cancel_operation:
End Sub



Private Sub Form_Load()
    CopyBMP = False
    PasteBMP = False
    PrintText = False
    Toolbar1.Visible = True
    Toolbar2.Visible = True
    toolcolor.Checked = True
    toolwidth.Checked = True
    tooldrawbox.Checked = True
    DrawFreehand.Checked = True
    shape = "FREEHAND"
    If Clipboard.GetFormat(vbCFBitmap) Then
    EditPaste.Enabled = True
    Else
    EditPaste.Enabled = False
    End If
    dragmode = False
    Picture3.Width = Picture2.Width
    Picture3.Height = Picture2.Height
    Picture1.Width = Picture2.Width
    Picture1.Height = Picture2.Height
    Load Form2
    Load Form3
    
End Sub

Private Sub Editmenu_click()
If Clipboard.GetFormat(vbCFBitmap) Then
        EditPaste.Enabled = True
        Else
        EditPaste.Enabled = False
        End If

End Sub
Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo op_abort
    If Button = 2 Then
        If Clipboard.GetFormat(vbCFBitmap) Then
        EditPaste.Enabled = True
        Else
        EditPaste.Enabled = False
        End If

        Form1.PopupMenu EditMenu
        
    End If
    
    If Button = 1 Then
        
        If shape <> "POLYGON" Or polydrawn <> False Then
        xold = x
        yold = y
        
        XStart = x
        YStart = y
        polydrawn = False
        End If
        
        If shape = "CURVE" And bezier_stepthru = 0 Then
        p0.x = XStart
        p0.y = YStart
        End If
        
        If shape = "ERASER" Then
            PDrawWidth = Picture2.DrawWidth
            PDrawStyle = Picture2.DrawStyle
            PPencolor = Picture2.ForeColor
            Picture2.DrawWidth = 10
            Picture2.DrawStyle = 0
            Picture2.ForeColor = RGB(255, 255, 255)
        End If
        
        If shape = "AIRBRUSH" Then
        Picture2.AutoRedraw = True
        For counti = 0 To 30
        airbrushx = x + CInt(Rnd() * 10 - 5)
        airbrushy = y + CInt(Rnd() * 10 - 5)
        If airbrushx < 0 Then airbrushx = 0
        If airbrushy < 0 Then airbrushy = 0
        SetPixelV Picture2.hdc, airbrushx, airbrushy, Picture2.ForeColor
        Next
        Picture2.Refresh
        End If
        
        
        
        XPrevious = XStart
        YPrevious = YStart
        Picture2.AutoRedraw = False
        
    

    If CopyBMP Or CutBMP Then
        PDrawWidth = Picture2.DrawWidth
        PDrawStyle = Picture2.DrawStyle
        PFillStyle = Picture2.FillStyle
        PPencolor = Picture2.ForeColor
        Picture2.DrawWidth = 1
        Picture2.DrawStyle = 2
        Picture2.FillStyle = 1
        Picture2.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    If PasteBMP Then
        Picture2.PaintPicture Picture1.Image, x, y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
        XPrevious = x
        YPrevious = y
     Exit Sub
    End If
    
If PrintText Then
    'Label1.Visible = True
    'Label1.Left = x
    'Label1.Top = y
    Picture2.CurrentX = x
    Picture2.CurrentY = y
    Picture2.Print drawstring
    Exit Sub
   End If
   If shape = "FLOODFILL" Then
    Picture2.AutoRedraw = True
    Call MyFloodfill(x, y, Picture2.ForeColor)
End If
End If
Exit Sub
op_abort:
End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo op_abort
    StatusBar1.Panels(2).Text = "X: " & x & "   Y: " & y
    If Button <> 1 Then Exit Sub
    If CopyBMP Or CutBMP Then
        'Picture2.Line (XStart, YStart)-(XPrevious, YPrevious), , B
        Picture2.Refresh
        Debug.Print "x=" & x & " y=" & y
        If x > Picture2.Width Then
        x = Picture2.Width
        ElseIf x < 0 Then
        x = 0
        End If
        
        If y > Picture2.Height Then
        y = Picture2.Height
        ElseIf y < 0 Then
        y = 0
        End If
        Debug.Print "cx=" & x & " cy=" & y
        Picture2.Line (XStart, YStart)-(x, y), , B
        XPrevious = x
        YPrevious = y
        Exit Sub
    End If
    If PasteBMP Then
        'picture2.PaintPicture Picture1.Image, XPrevious, YPrevious, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &H660046
        Picture2.Refresh
        Picture2.PaintPicture Picture1.Image, x, y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
    Exit Sub
    End If
If PrintText Then
    Picture2.Refresh
    Picture2.CurrentX = x
    Picture2.CurrentY = y
    Picture2.Print drawstring
    Exit Sub
End If
    
    Select Case shape
        Case "LINE":
            'picture2.Line (XStart, YStart)-(XPrevious, YPrevious)
            Picture2.Refresh
            Picture2.Line (XStart, YStart)-(x, y)
        Case "CIRCLE":
            'picture2.Circle (XStart, YStart), Sqr((XPrevious - XStart) ^ 2 + (YPrevious - YStart) ^ 2)
            Picture2.Refresh
            Picture2.Circle (XStart, YStart), Sqr((x - XStart) ^ 2 + (y - YStart) ^ 2)

        Case "BOX":
            'picture2.Line (XStart, YStart)-(XPrevious, YPrevious), , B
            Picture2.Refresh
            Picture2.Line (XStart, YStart)-(x, y), , B
        Case "FREEHAND":
            Picture2.AutoRedraw = True
            Picture2.Line (xold, yold)-(x, y)
            xold = x
            yold = y
        Case "POLYGON":
            Picture2.Refresh
            Picture2.Line (xold, yold)-(x, y)
        Case "ERASER":
            Picture2.AutoRedraw = True
            Picture2.Line (xold, yold)-(x, y)
            xold = x
            yold = y
        Case "CURVE":
            Picture2.Refresh
            If bezier_stepthru = 0 Then
            Picture2.Line (xold, yold)-(x, y)
            ElseIf bezier_stepthru = 1 Then
            p1.x = x - 5
            p1.y = y
            
            p2.x = x + 5
            p2.y = y
            
            
            curvex = p0.x
            curvey = p0.y
            MoveToEx Picture2.hdc, curvex, curvey, imgmod.point
            
            For u = 0.01 To 1 Step 0.01
            curvex = (1 - u) * (1 - u) * (1 - u) * p0.x
            curvex = curvex + 3 * (1 - u) * (1 - u) * u * p1.x
            curvex = curvex + 3 * (1 - u) * u * u * p2.x
            curvex = curvex + u * u * u * p3.x
            
            curvey = (1 - u) * (1 - u) * (1 - u) * p0.y
            curvey = curvey + 3 * (1 - u) * (1 - u) * u * p1.y
            curvey = curvey + 3 * (1 - u) * u * u * p2.y
            curvey = curvey + u * u * u * p3.y
            
            
            LineTo Picture2.hdc, curvex, curvey
            Next
            End If
            Case "AIRBRUSH":
                Picture2.AutoRedraw = True
                For counti = 0 To 30
                airbrushx = x + CInt(Rnd() * 10 - 5)
                airbrushy = y + CInt(Rnd() * 10 - 5)
                If airbrushx < 0 Then airbrushx = 0
                If airbrushy < 0 Then airbrushy = 0
                SetPixelV Picture2.hdc, airbrushx, airbrushy, Picture2.ForeColor
                Next
                Picture2.Refresh
        
        End Select
    Exit Sub
op_abort:
End Sub

Private Sub picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo op_abort
Dim X1 As Single, Y1 As Single

    If CopyBMP Or CutBMP Then
    If x > Picture2.Width Then
        x = Picture2.Width
        ElseIf x < 0 Then
        x = 0
        End If
        
        If y > Picture2.Height Then
        y = Picture2.Height
        ElseIf y < 0 Then
        y = 0
        End If
    End If
        
    If Button = 1 Then
        If CopyBMP Then
        
        'Picture2.Line (XStart, YStart)-(XPrevious, YPrevious), , B
        Picture2.Refresh
        If x > XStart Then X1 = XStart Else X1 = x
        If y > YStart Then Y1 = YStart Else Y1 = y
        Picture1.Picture = LoadPicture()
        Picture1.Width = Abs(x - XStart)
        Picture1.Height = Abs(y - YStart)
        Picture1.PaintPicture Picture2.Image, 0, 0, Abs(x - XStart), Abs(y - YStart), X1, Y1, Abs(x - XStart), Abs(y - YStart), &HCC0020
        Clipboard.Clear
        Clipboard.SetData Picture1.Image, vbCFBitmap
        CopyBMP = False
        Picture2.DrawWidth = PDrawWidth
        Picture2.DrawStyle = PDrawStyle
        Picture2.FillStyle = PFillStyle
        Picture2.ForeColor = PPencolor
        CopyWidth = Abs(x - XStart)
        CopyHeight = Abs(y - YStart)
        EditPaste.Enabled = True
        Exit Sub
    End If
    If CutBMP Then
        Picture2.AutoRedraw = True
        CopyWidth = XStart - x
        CopyHeight = YStart - y
        If x > XStart Then X1 = XStart Else X1 = x
        If y > YStart Then Y1 = YStart Else Y1 = y
        Picture1.Picture = LoadPicture()
        Picture1.Width = Abs(x - XStart)
        Picture1.Height = Abs(y - YStart)
        Picture1.PaintPicture Picture2.Image, 0, 0, Abs(x - XStart), Abs(y - YStart), X1, Y1, Abs(x - XStart), Abs(y - YStart), &HCC0020
        Clipboard.Clear
        Clipboard.SetData Picture1.Image, vbCFBitmap
        Picture2.DrawStyle = 0
        Picture2.Line (x, y)-Step(CopyWidth, CopyHeight), Picture2.BackColor, BF
        CutBMP = False
        Picture2.DrawWidth = PDrawWidth
        Picture2.DrawStyle = PDrawStyle
        Picture2.FillStyle = PFillStyle
        Picture2.ForeColor = PPencolor
        CopyWidth = Abs(x - XStart)
        CopyHeight = Abs(y - YStart)
        EditPaste.Enabled = True
        Exit Sub
    End If
    
    If PasteBMP Then
          Picture2.AutoRedraw = True
          Picture2.PaintPicture Picture1.Image, x, y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
          PasteBMP = False
        
        Exit Sub
    End If
    
    If PrintText Then
        Picture2.AutoRedraw = True
        Picture2.CurrentX = x
        Picture2.CurrentY = y
        Picture2.Print drawstring
        'Label1.Visible = False
        PrintText = False
        Toolbar3.Buttons.Item(7).value = tbrUnpressed
        
        Exit Sub
    End If

    'picture2.DrawMode = 13
    Picture2.Refresh
    Picture2.AutoRedraw = True
    Select Case shape
        Case "LINE":
            Picture2.Line (XStart, YStart)-(x, y)
        Case "CIRCLE":
            Picture2.Circle (XStart, YStart), Sqr((x - XStart) ^ 2 + (y - YStart) ^ 2)
            'Ellipse Picture2.hdc, XStart, YStart, x, y
        Case "BOX":
            Picture2.Line (XStart, YStart)-(x, y), , B
        Case "POLYGON":
            Picture2.Line (xold, yold)-(x, y)
            xold = x
            yold = y
        Case "ERASER":
            Picture2.DrawWidth = PDrawWidth
            Picture2.DrawStyle = PDrawStyle
            Picture2.ForeColor = PPencolor
        Case "CURVE":
            If bezier_stepthru = 0 Then
            Picture2.AutoRedraw = False
            Picture2.Line (XStart, YStart)-(x, y)
            p3.x = x
            p3.y = y
            bezier_stepthru = 1
            ElseIf bezier_stepthru = 1 Then
            Picture2.AutoRedraw = True
            p1.x = x - 5
            p1.y = y
            
            p2.x = x + 5
            p2.y = y
            
            
            curvex = p0.x
            curvey = p0.y
            MoveToEx Picture2.hdc, curvex, curvey, imgmod.point
            
            For u = 0.01 To 1 Step 0.01
            curvex = (1 - u) * (1 - u) * (1 - u) * p0.x
            curvex = curvex + 3 * (1 - u) * (1 - u) * u * p1.x
            curvex = curvex + 3 * (1 - u) * u * u * p2.x
            curvex = curvex + u * u * u * p3.x
            
            curvey = (1 - u) * (1 - u) * (1 - u) * p0.y
            curvey = curvey + 3 * (1 - u) * (1 - u) * u * p1.y
            curvey = curvey + 3 * (1 - u) * u * u * p2.y
            curvey = curvey + u * u * u * p3.y
            
            
            LineTo Picture2.hdc, curvex, curvey
            Next
            bezier_stepthru = 0
            Picture2.Refresh
            End If
    End Select
       
    End If
    
    Exit Sub
op_abort:
    
End Sub

Private Sub Form_Resize()
    On Error GoTo pos1
    If Form1.Width < 6060 Then
        Form1.Width = 6060
    End If
    
    If Form1.Height < 5940 Then
        Form1.Height = 5940
    End If
    
    
pos1:
   ' Picture1.Width = Picture2.Width
    'Picture1.Height = Picture2.Height
End Sub

Private Sub StyleDash_Click()
    UnCheckStyles
    StyleDash.Checked = True
    Picture2.DrawStyle = 1
End Sub

Private Sub StyleDot_Click()
    UnCheckStyles
    StyleDot.Checked = True
    Picture2.DrawStyle = 2
End Sub


Private Sub StyleSolid_Click()
    UnCheckStyles
    StyleSolid.Checked = True
    Picture2.DrawStyle = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    uncheckwidth
    If Button.Index = 1 Then
        Picture2.DrawWidth = 1
        width1.Checked = True
    ElseIf Button.Index = 2 Then
        Picture2.DrawWidth = 2
        Width2.Checked = True
    Else
        Picture2.DrawWidth = 3
        Width3.Checked = True
    End If
    
End Sub

Private Sub width1_Click()
    uncheckwidth
    width1.Checked = True
    Picture2.DrawWidth = 1
End Sub

Private Sub Width2_Click()
    uncheckwidth
    Width2.Checked = True
    Picture2.DrawWidth = 2
End Sub

Private Sub Width3_Click()
    uncheckwidth
    Width3.Checked = True
    Picture2.DrawWidth = 3
End Sub

Private Sub set_drawmode(ByVal drawmode As String)
    Select Case drawmode
        Case "Freehand":
            
            shape = "FREEHAND"
            uncheckdraw
            DrawFreehand.Checked = True
            Picture2.MouseIcon = LoadPicture("common\cursors\cursor2.cur")
            Picture2.MousePointer = 99
            Toolbar3.Buttons.Item(1).value = tbrPressed
            
        Case "Line":
            
            shape = "LINE"
            Picture2.MousePointer = 2
            uncheckdraw
            DrawLine.Checked = True
            Toolbar3.Buttons.Item(2).value = tbrPressed
        
        Case "Circle":
            
            shape = "CIRCLE"
            uncheckdraw
            DrawCircle.Checked = True
            Picture2.MousePointer = 2
            Toolbar3.Buttons.Item(3).value = tbrPressed
        Case "Box":
            
            shape = "BOX"
            uncheckdraw
            DrawBox.Checked = True
            Picture2.MousePointer = 2
            Toolbar3.Buttons.Item(4).value = tbrPressed
        Case "Text":
            Toolbar3.Buttons.Item(7).value = tbrPressed
            Picture2.MousePointer = 2
            uncheckdraw
            DrawText.Checked = True
            drawstring = InputBox("Enter string")
            If drawstring = "" Then
               Toolbar3.Buttons.Item(7).value = tbrUnpressed
               PrintText = False
               Exit Sub
            End If
            'Label1.Caption = drawstring
            'Label1.ForeColor = Picture2.ForeColor
            PrintText = True
        Case "Polygon":
            shape = "POLYGON"
            uncheckdraw
            DrawPoly.Checked = True
            Picture2.MousePointer = 2
            polydrawn = True
            Toolbar3.Buttons.Item(5).value = tbrPressed
            'Dim i As Integer
        Case "FloodFill":
            Picture2.MousePointer = 2
            shape = "FLOODFILL"
            Toolbar3.Buttons.Item(8).value = tbrPressed
        Case "Eraser":
            shape = "ERASER"
            Picture2.MouseIcon = LoadPicture("common\cursors\cursor3.cur")
            Picture2.MousePointer = 99
            
        Case "Curve":
            shape = "CURVE"
            uncheckdraw
            DrawCurve.Checked = True
            Picture2.MousePointer = 2
            bezier_stepthru = 0
            Toolbar3.Buttons.Item(9).value = tbrPressed
        Case "AirBrush":
            shape = "AIRBRUSH"
            uncheckdraw
            DrawAirBrush.Checked = True
            Picture2.MousePointer = 2
            Toolbar3.Buttons.Item(10).value = tbrPressed
    End Select
End Sub


Sub MyFloodfill(ByVal x As Integer, ByVal y As Integer, ByVal fill_color As Long)
    
    Dim brush
    'MsgBox ("I am called")
    ' create solid brush object
    brush = CreateSolidBrush(fill_color)
    ' select it into the current device context
    SelectObject Picture2.hdc, brush
    ' fill area with selected brush
    ExtFloodFill Picture2.hdc, x, y, Picture2.point(x, y), 1
    ' brush object no longer needed, delete it
    DeleteObject brush

    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Remove everything related to the application from Memory
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload frmabout

End Sub

