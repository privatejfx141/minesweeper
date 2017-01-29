VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmStats 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   Icon            =   "Stats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraModes 
      Caption         =   "Select Difficulty"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optMode 
         Caption         =   "Expert"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Intermediate"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Beginner"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Re&set Scores"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton frmReturn 
      Caption         =   "&Return"
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid grdScores 
      Height          =   1815
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   6
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Culminating Activity - Minesweeper
'Author: Yun Jie (Jeffrey) Li
'Date: June 2, 2016
'Files: Minesweeper.vbp, Game.bas, Game.frm, Game.frx, Splash.frm. Splash.frx, Options.frm,
'       Stats.frm, About.frm
'Purpose: The purpose of this program is the recreate the classical computer puzzle game
'         of Minesweeper.

Dim SelectedMode As Integer
Dim ScoreLen As Integer

Option Explicit

Private Sub cmdReset_Click()
    
    Dim FileName As String
    Dim DMsg As String
    Dim Response As Integer
    Dim CMode As Integer
    
    DMsg = "Are you sure you want to reset all high scores?"
    Response = MsgBox(DMsg, vbYesNo + vbQuestion, "Reset Scores")
    
    If Response = vbYes Then
        ClearScoreRecords
        For CMode = EASY To HARD
            FileName = App.Path & FILENAME_START & CMode & FILENAME_EXT
            SaveFile CMode, FileName, MAX_DISPLAY_RECS, RecordLen
        Next CMode
        Form_Load
    End If
    
End Sub

Private Sub Form_Load()
        
    Dim CRow, CCol As Integer
    
    SetClearScoreBoard grdScores
    DisplayScores grdScores, EASY
    CentreForm Me
    
End Sub

Private Sub frmReturn_Click()
    
    Unload Me
    
End Sub

Private Sub optMode_Click(Index As Integer)
    
    SetClearScoreBoard grdScores
    DisplayScores grdScores, Index
    
End Sub
