VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper"
   ClientHeight    =   5535
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6375
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFace 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   2040
      Picture         =   "Game.frx":0A8A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grdGame 
      Height          =   4410
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7779
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin VB.Image imgFace 
      Height          =   780
      Index           =   3
      Left            =   5520
      Picture         =   "Game.frx":10BC
      Top             =   4320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgFalseFlag 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":16EE
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFace 
      Height          =   780
      Index           =   1
      Left            =   5520
      Picture         =   "Game.frx":1970
      Top             =   2640
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgFace 
      Height          =   780
      Index           =   2
      Left            =   5520
      Picture         =   "Game.frx":1FA2
      Top             =   3480
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":25D4
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFace 
      Height          =   780
      Index           =   0
      Left            =   5520
      Picture         =   "Game.frx":2856
      Top             =   1800
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":2E88
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMines 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   945
   End
   Begin VB.Image imgMine 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":310A
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgExplode 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":338C
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTileClick 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":360E
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   480
      Left            =   4920
      Picture         =   "Game.frx":3890
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   945
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuStats 
         Caption         =   "&High Scores"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplash 
         Caption         =   "&Splash"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Culminating Activity - Minesweeper
'Author: Yun Jie (Jeffrey) Li
'Date: May 9, 2016
'Files: Minesweeper.vbp, Game.bas, Game.frm, Game.frx, Splash.frm. Splash.frx
'Purpose: The purpose of this program is the recreate the classical computer puzzle game
'         of Minesweeper.

'Declare the form variables.

Dim NewScore As ScoreRec
Dim Time As Integer
Dim GameOver As Boolean
Dim FirstClick As Boolean

Option Explicit

Private Sub About_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub Form_Load()
    
    InitializeScores NewScore
    
    If TotalMines = 0 Or Rows = 0 Or Columns = 0 Then
        TotalMines = EASY_MINES
        Rows = 9
        Columns = 9
    End If
    
    GameOver = False
    tmrGame.Enabled = False
    cmdFace.Picture = imgFace(0).Picture
    NumMines = TotalMines
    lblMines.Caption = TotalMines
    Time = 0
    lblTime.Caption = "000"
    
    With grdGame
    
        .Visible = False
        
        InitializeBoard grdGame, imgTile
        SetFormSize Me, grdGame, Rows, Columns, cmdFace, lblMines, lblTime
        CentreForm Me
        SetMines grdGame, TotalMines
        CheckRadius grdGame
        ReadCellsIntoArray grdGame
        InitializeBoard grdGame, imgTile
        
        .Visible = True
        FirstClick = True
        
    End With
    
End Sub

Private Sub grdGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FirstClick = False
    
    With grdGame
    
    .Visible = False
    
    .Row = .MouseRow
    .Col = .MouseCol
    
    If Not GameOver And .CellPicture <> imgTileClick.Picture Then
    
        If Button = vbLeftButton Then
            
            If .CellPicture <> imgFlag.Picture Then
                If Not GameOver And Not FirstClick Then
                    tmrGame.Enabled = True
                End If
                CellLeftClick grdGame, tmrGame, GameOver, imgTile, imgTileClick, imgMine, _
                    imgExplode, imgFlag, imgFalseFlag
                If GameOver Then
                    EndGame grdGame, LOST, cmdFace, imgFace, TotalMines, Time, NewScore
                End If
            End If
            
        ElseIf Button = vbRightButton Then
            
            CellRightClick grdGame, imgTile, imgFlag, imgQuestion, NumMines
            lblMines.Caption = NumMines
            
            If NumMines = 0 Then
            
                If IsGameWin(grdGame, imgFlag) Then
                
                    tmrGame.Enabled = False
                    GameOver = True
                    EndGame grdGame, WIN, cmdFace, imgFace, TotalMines, Time, NewScore
                    
                End If
                
            End If
            
        End If 'Left or Right button statement.
    
    End If 'Not GameOver statement.
    
    .Visible = True
    
    End With
    
End Sub

Private Sub cmdFace_Click()
    
    mnuNewGame_Click
    
End Sub

Private Sub mnuExit_Click()
           
    ExitProgram
    
End Sub

Private Sub mnuNewGame_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "New Game?"
    DMsg = "Game is currently in progress."
    DMsg = DMsg & vbCrLf & "Are you sure you want to start a new game?"
    
    If Not FirstClick And Not GameOver Then
        Response = MsgBox(DMsg, DType, DTitle)
        If Response = vbYes Then
            Form_Load
        End If
    Else
        Form_Load
    End If
    
End Sub

Private Sub mnuOptions_Click()
        
    frmOptions.Show vbModal
    
End Sub

Private Sub mnuSplash_Click()
        
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "New Game?"
    DMsg = "Game is currently in progress."
    DMsg = DMsg & vbCrLf & "Are you sure you want to exit the current game"
    DMsg = DMsg & vbCrLf & "in order to display the splash screen?"
    If Not FirstClick And Not GameOver Then
        Response = MsgBox(DMsg, vbQuestion + vbYesNo, "Exit and display splash?")
        If Response = vbYes Then
            frmSplash.Show
            Unload Me
        End If
    Else
        frmSplash.Show
        Unload Me
    End If
    
End Sub

Private Sub mnuStats_Click()
    
    frmStats.Show vbModal
    
End Sub

Private Sub tmrGame_Timer()
    
    Time = Time + 1
    lblTime.Caption = VBA.Format$(Time, "000")
    
    If Time = MAX_TIME Then
        tmrGame.Enabled = False
        GameOver = True
        cmdFace.Picture = imgFace(3).Picture
        MsgBox "Times Up!"
    End If
    
End Sub
