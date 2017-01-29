VERSION 5.00
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton frmReturn 
      Caption         =   "&Return"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Frame fraDifficulty 
      Caption         =   "Difficulty"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtMines 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Text            =   "10"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Text            =   "9"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHeight 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         Text            =   "9"
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Custom"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Expert"
         Height          =   735
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Intermediate"
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Beginner"
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Mines (10-500):"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Width (9-30):"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height (9-18):"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Dim Selected As Integer

Option Explicit

Private Sub cmdConfirm_Click()
    
    Dim Valid As Boolean
    Dim DialogMsg As String
    Valid = True
    
    SetDifficulty Selected, txtHeight, txtWidth, txtMines
    If optMode(3).Value = True Then
        Valid = IsCustomSettingsValid(DialogMsg)
    End If
    
    If Valid Then
        TotalMines = NumMines
        Unload Me
        Unload frmGame
        frmGame.Show
    Else
        MsgBox DialogMsg, vbCritical, "Error! Invalid Settings!"
    End If
    
End Sub

Private Sub Form_Load()
    
    Selected = LastSelectedMode
    optMode(Selected).Value = True
    
    optMode(0).Caption = "Beginner" & vbCrLf & "10 mines" & vbCrLf & "9 x 9 tile grid"
    optMode(1).Caption = "Intermediate" & vbCrLf & "40 mines" & vbCrLf & "16 x 16 tile grid"
    optMode(2).Caption = "Expert" & vbCrLf & "99 mines" & vbCrLf & "16 x 30 tile grid"
    
    If LastSelectedMode = 3 Then
        txtHeight.Text = Rows
        txtWidth.Text = Columns
        txtMines.Text = NumMines
    End If
    
    CentreForm Me
    
End Sub

Private Sub frmReturn_Click()
    
    Unload Me
    
End Sub

Private Sub optMode_Click(Index As Integer)
    
    Selected = Index
    LastSelectedMode = Selected
    
    If Index = CUSTOM Then
        txtHeight.Enabled = True
        txtWidth.Enabled = True
        txtMines.Enabled = True
    Else
        txtHeight.Enabled = False
        txtWidth.Enabled = False
        txtMines.Enabled = False
    End If
    
End Sub
