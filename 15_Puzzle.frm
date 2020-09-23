VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "The 15 Puzzle"
   ClientHeight    =   2910
   ClientLeft      =   3375
   ClientTop       =   2865
   ClientWidth     =   5490
   Icon            =   "15_Puzzle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5490
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FF0000&
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF0000&
      Caption         =   "E&xit"
      DownPicture     =   "15_Puzzle.frx":044A
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   1200
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   2880
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H001E35DB&
      X1              =   240
      X2              =   2640
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   2640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H001E35DB&
      X1              =   2640
      X2              =   2640
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Moves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblMoves 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblClock 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Your Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape shpBacking 
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   120
      Shape           =   1  'Square
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   2880
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu separtor 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHow 
         Caption         =   "How To &Play"
         Shortcut        =   ^H
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicNumber(1 To 16) As Integer, I As Integer, GridPos As Integer, Row As Integer, Column As Integer, Pos As Integer

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNew_Click()
   Random
   tmrClock.Enabled = True
   lblClock.Caption = 0
   lblMoves.Caption = 0
    
End Sub



Private Sub Form_Load()
Random
tmrClock.Enabled = True
End Sub

Private Sub lblTile_Click(Index As Integer)
Dim Hi As Integer, Side As Integer
If lblTile(Index).Top = lblTile(16).Top + 615 And lblTile(Index).Left = lblTile(16).Left Or lblTile(Index).Top = lblTile(16).Top - 615 And lblTile(Index).Left = lblTile(16).Left Or lblTile(Index).Top = lblTile(16).Top And lblTile(Index).Left = lblTile(16).Left - 615 Or lblTile(Index).Top = lblTile(16).Top And lblTile(Index).Left = lblTile(16).Left + 615 Then
Hi = lblTile(Index).Top
Side = lblTile(Index).Left
lblTile(Index).Top = lblTile(16).Top
lblTile(Index).Left = lblTile(16).Left
lblTile(16).Top = Hi
lblTile(16).Left = Side
lblMoves.Caption = lblMoves.Caption + 1
Else
Beep
End If
End Sub

Private Sub Random()
Static Occupied(1 To 16) As Boolean
    Dim I As Integer, Row As Integer, Column As Integer
    Randomize
    For I = 1 To 16
        Occupied(I) = False
    Next I
    For I = 1 To 16
        Do
            Row = Int(Rnd * 4) + 1
            Column = Int(Rnd * 4) + 1
            GridPos = (Row - 1) * 4 + Column
        Loop While Occupied(GridPos)
        Occupied(GridPos) = True
        PicNumber(GridPos) = I
        lblTile(I).Move (210 + (Column - 1) * 615), (210 + (Row - 1) * 615)
        Next I
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Program created and designed in Visual Basic 6 by Ken McKeon And Mathew Ramsey. 2001", 64, "The 15 Puzzle"
End Sub

Private Sub mnuHow_Click()
    Help.Show
End Sub

Private Sub mnuNew_Click()
     Random
   tmrClock.Enabled = True
   lblClock.Caption = 0
   lblMoves.Caption = 0
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub tmrClock_Timer()
lblClock.Caption = lblClock.Caption + (tmrClock.Interval / 1000)
End Sub
