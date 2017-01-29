VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&How to Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdScores 
      BackColor       =   &H00E0E0E0&
      Caption         =   "High &Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "MINESWEEPER"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSplash 
         Caption         =   "&Splash"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuInstructions 
         Caption         =   "H&ow to Play"
      End
   End
End
Attribute VB_Name = "frmMenu"
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

Option Explicit

Private Sub cmdExit_Click()
    
    mnuExit_Click
    
End Sub

Private Sub cmdPlay_Click()

    frmGame.Show
    Unload Me
    
End Sub

Private Sub Form_Load()

    CentreForm Me
    Me.Width = 11610
    Me.Height = 6240
    
End Sub

Private Sub mnuGame_Click()

End Sub

Private Sub mnuExit_Click()
    
    ExitProgram
    
End Sub

Private Sub mnuInstructions_Click()
    
    MsgBox "Coming soon!"
    
End Sub

Private Sub mnuSplash_Click()
    
    frmSplash.Show
    Unload Me
    
End Sub

Private Sub tmrTimer_Timer()

    Time = Time + 1
    lblTime.Caption = Format$(Time \ 60, "00") & ":" & Format$(Time Mod 60, "00")

End Sub
