VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Return"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblDesc 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INSERT DESCRIPTION HERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MINESWEEPER"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdReturn_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim Msg As String
    
    Msg = "This program allows you to play the classic single-player puzzle game of Minesweeper. "
    Msg = Msg & "For instructions on how to play the game, please view the user manual."
    Msg = Msg & vbCrLf & vbCrLf & "Programmed by Jeffrey Li 12H,"
    Msg = Msg & vbCrLf & "Riverdale Collegiate Institute, 2016."
    
    lblDesc.Caption = Msg
    
    CentreForm Me
    
End Sub
