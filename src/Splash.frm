VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Minesweeper"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":0A8A
   ScaleHeight     =   10412.97
   ScaleMode       =   0  'User
   ScaleWidth      =   5451.104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Interval        =   3000
      Left            =   9000
      Top             =   120
   End
End
Attribute VB_Name = "frmSplash"
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

Option Explicit

Private Sub Form_Load()
    
    CentreForm Me
    
End Sub


Private Sub tmrTimer_Timer()
    
    frmGame.Show
    Unload Me
    
End Sub
