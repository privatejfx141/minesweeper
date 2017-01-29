Attribute VB_Name = "modGame"
'Name: Culminating Activity - Minesweeper
'Author: Yun Jie (Jeffrey) Li
'Date: June 2, 2016
'Files: Minesweeper.vbp, Game.bas, Game.frm, Game.frx, Splash.frm. Splash.frx, Options.frm,
'       Stats.frm, About.frm
'Purpose: The purpose of this program is the recreate the classical computer puzzle game
'         of Minesweeper.

'Declare the global constants.

Global Const FILENAME_START = "\Scores_", FILENAME_EXT = ".rec"
Global Const EASY = 0, MEDIUM = 1, HARD = 2, CUSTOM = 3
Global Const LOST = 0, WIN = 1

Global Const MAX_LENGTH = 30
Global Const MINE = -1, BLANK = 0
Global Const CELLWIDTH = 480 'Twips

Global Const EASY_MINES = 10, MEDIUM_MINES = 40, HARD_MINES = 99

Global Const BLANK_REC = "@#$%IN-VAL1D0!" 'Arbitrary name that no one would probably use.

Global Const MAX_TIME = 999
Global Const MAX_DISPLAY_RECS = 5
Global Const MAX_RECS = 6 '6th record is always discarded.

'Declare the global two-dimensional array.

Global CellValue(0 To MAX_LENGTH, 0 To MAX_LENGTH) As Integer 'Value at (Column, Row).

Global TotalMines As Integer
Global NumMines As Integer
Global Rows As Integer
Global Columns As Integer
Global LastSelectedMode As Integer
Global NumRecords(0 To 2) As Integer
Global RecordLen As Integer

'Declare the user-defined data type.

Type ScoreRec
    Name As String * 20
    Time As Integer
End Type

Global Score(EASY To HARD, 1 To MAX_RECS) As ScoreRec

Option Explicit

'This GP asks for the user's request to exit the program.

Public Sub ExitProgram()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMsg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "Exit program?"
    DMsg = "Are you sure you want to exit Minesweeper?"
    
    Response = MsgBox(DMsg, DType, DTitle)
    If Response = vbYes Then
        End
    End If
    
End Sub

Public Sub SwapScores(Item1 As ScoreRec, Item2 As ScoreRec)

    Dim Temp As ScoreRec
    
    Temp = Item1
    Item1 = Item2
    Item2 = Temp

End Sub

'This GP determines if the opened file is a text file or a record file, and then it
'reads the data into the variables.

Public Sub ReadFile(ByVal Mode As Integer, ByVal FileName As String, NumRecs As Integer, _
    ByVal RecordLen As Integer)

    Dim K As Integer
    
    On Error GoTo ErrorHandler
    
    K = 0
    If VBA.Right$(FileName, 3) = "rec" Then
        
        'Read the file using this procedure, if the opened file is a record file.
        
        Open FileName For Random As #1 Len = RecordLen
            Do While Not EOF(1)
                K = K + 1
                Get #1, K, Score(Mode, K)
            Loop
        Close #1

        NumRecs = K - 1

    Else
    
        'Read the file using this procedure, if the opened file is a text file.
        
        Open FileName For Input As #1
            Do While Not EOF(1)
                K = K + 1
                With Score(Mode, K)
                    Input #1, .Name, .Time
                End With
            Loop
        Close #1
        
        NumRecs = K
        
    End If
    
Exit Sub
    
ErrorHandler:
    Resume Next
    
End Sub

'This GP saves a record file, by inputing data into a newly-created record file.

Public Sub SaveFile(GameMode As Integer, ByVal FileName As String, ByVal NumRecs As Integer, _
    ByVal RecordLen As Integer)

    Dim K As Integer
    On Error GoTo ErrorHandler
    Kill FileName
    
    Open FileName For Random As #1 Len = RecordLen
        For K = 1 To NumRecs
            Put #1, K, Score(GameMode, K)
        Next K
    Close #1
    
    Exit Sub
    
ErrorHandler:
    Resume Next
    
End Sub

'This GP centres the form on the screen.

Public Sub CentreForm(CurrentForm As Form)
    
    With CurrentForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
End Sub

'This GP clears and initializes the game board.

Public Sub InitializeBoard(Grid As Control, imgTile As Control)
        
    Dim CCol, CRow As Integer
    
    With Grid
    
    For CRow = 0 To .Rows - 1
        .Row = CRow
        .RowHeight(CRow) = 480
        For CCol = 0 To .Cols - 1
            .Col = CCol
            .ColWidth(CCol) = 480
            .Text = ""
            .CellAlignment = 4
            .CellBackColor = vbWhite
            .CellForeColor = vbBlack
            .CellFontBold = True
            .CellFontSize = 12
            .CellPicture = imgTile.Picture
        Next CCol
    Next CRow
    
    End With
End Sub

'This GP places mines on the game board.

Public Sub SetMines(Grid As Control, ByVal NumMines As Integer)
    
    Const LOW = 0, HIGH = 10
    Dim CCol, CRow As Integer
    Dim RndValue As Integer
    Dim DeployedMines As Integer
    
    If NumMines = 0 Then
        NumMines = 10
    End If
    
    Randomize
    
    With Grid
    
    Do
        For CRow = 0 To .Rows - 1
            .Row = CRow
            For CCol = 0 To .Cols - 1
                .Col = CCol
                RndValue = Int((HIGH - LOW + 1) * Rnd) + LOW
                
                '7 is an arbitrary value.
                If (RndValue = 7) And (DeployedMines <> NumMines) And _
                    (.Text <> MINE) Then
                    .Text = MINE
                    DeployedMines = DeployedMines + 1
                End If
                
            Next CCol
        Next CRow
    Loop Until (DeployedMines = NumMines)
    
    End With
    
End Sub

'This GP reads the values of the cells on the game board into a two-dimensional array.

Public Sub ReadCellsIntoArray(Grid As Control)

    Dim CCol, CRow As Integer
    
    With Grid
    
    For CRow = 0 To .Rows - 1
        .Row = CRow
        For CCol = 0 To .Cols - 1
            .Col = CCol
            CellValue(CCol, CRow) = Val(.Text)
        Next CCol
    Next CRow
    
    End With
    
End Sub

'If the player clicks on the mine, this GP will display the locations of all mines
'and falsely flagged cells on the game board.

Public Sub ShowMines(Grid As Control, ByVal CurrentCol As Integer, ByVal CurrentRow As Integer, _
    imgMine As Control, imgExplode As Control, imgFalseFlag As Control, imgFlag As Control)

    Dim CCol, CRow As Integer
    
    With Grid
    
    For CRow = 0 To .Rows - 1
        .Row = CRow
        For CCol = 0 To .Cols - 1
            .Col = CCol
            
            If CellValue(.Col, .Row) = MINE And .CellPicture <> imgFlag.Picture Then
                .CellPicture = imgMine.Picture
            ElseIf CellValue(.Col, .Row) <> MINE And .CellPicture = imgFlag.Picture Then
                .CellPicture = imgFalseFlag.Picture
            End If
            
        Next CCol
    Next CRow
    
    .Col = CurrentCol
    .Row = CurrentRow
    .CellPicture = imgExplode.Picture
    
    End With
    
End Sub

'This GP gives each non-mine cell its number by checking around that cell for any mines.

Public Sub CheckRadius(Grid As Control)

    Dim CCol As Integer
    Dim CRow As Integer
    Dim NumMines As Integer
    
    With Grid
    
    For CRow = 0 To .Rows - 1
        .Row = CRow
        For CCol = 0 To .Cols - 1
            .Col = CCol
            
            If .Text <> MINE Then
                
                NumMines = 0
                
                If (CCol - 1 >= 0) And (CRow - 1 >= 0) Then 'Top-left (-1, -1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, -1, -1
                End If
                If (CRow - 1 >= 0) Then 'Top (0, -1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, -1, 0
                End If
                If (CCol + 1 <= .Cols - 1) And (CRow - 1 >= 0) Then 'Top-right (1, -1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, -1, 1
                End If
                If (CCol + 1 <= .Cols - 1) Then 'Right (1, 0).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, 0, 1
                End If
                If (CCol + 1 <= .Cols - 1) And (CRow + 1 <= .Rows - 1) Then 'Bottom-right (1, 1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, 1, 1
                End If
                If (CRow + 1 <= .Rows - 1) Then 'Bottom (0, 1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, 1, 0
                End If
                If (CCol - 1 >= 0) And (CRow + 1 <= .Rows - 1) Then 'Bottom-left (-1, 1).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, 1, -1
                End If
                If (CCol - 1 >= 0) Then 'Left (-1, 0).
                    CheckRadius_CountMinesResetPos Grid, NumMines, CRow, CCol, 0, -1
                End If
                
                If NumMines > 0 Then
                    .Text = NumMines
                End If
                
            End If
            
        Next CCol
    Next CRow
    
    End With
    
End Sub

'This GP is a component of CheckRadius. It sets the offset position, counts any mines found,
'and resets the position back to the original cell.

Private Sub CheckRadius_CountMinesResetPos(Grid As Control, NumMines As Integer, _
    CurrentRow As Integer, CurrentCol As Integer, RowOffset As Integer, ColOffset As Integer)
    
    With Grid
        
        'Set the offset position.
        
        .Col = CurrentCol + ColOffset
        .Row = CurrentRow + RowOffset
        
        If .Text = MINE Then
            NumMines = NumMines + 1
        End If
                    
        .Col = CurrentCol
        .Row = CurrentRow
    
    End With
    
End Sub

'This GP displays the value of the cell, if the player clicks on a non-mine or non-blank cell.

Public Sub ClickShowCellValue(Grid As Control, imgTileClick As Image)
    
    With Grid
        .CellPicture = imgTileClick.Picture
        If CellValue(.Col, .Row) > 0 Then
            .Text = CellValue(.Col, .Row)
            .CellForeColor = ColourNumber(CellValue(.Col, .Row))
        End If
    End With
    
End Sub

'This GP determines what actions are committed when the player left-clicks on the grid.

Public Sub CellLeftClick(Grid As Control, tmrGame As Timer, GameOver As Boolean, _
    imgTile As Image, imgTileClick As Image, imgMine As Image, imgExplode As Image, _
    imgFlag As Image, imgFalseFlag As Image)

    With Grid
    
    .Visible = False
    
    'If the player selects a mine.
    
    If CellValue(.Col, .Row) = MINE Then

        ShowMines Grid, .Col, .Row, imgMine, imgExplode, imgFalseFlag, imgFlag
        tmrGame.Enabled = False
        GameOver = True
        
    ElseIf CellValue(.Col, .Row) = BLANK Then
    
        BlankReveal Grid, .Row, .Col, imgTileClick, imgFlag
    
    Else
    
        ClickShowCellValue Grid, imgTileClick

    End If
    
    .Visible = True
    
    End With

End Sub

'This GP determines what object should be placed on the cell when the player right-clicks.

Public Sub CellRightClick(Grid As Control, imgTile As Image, imgFlag As Image, _
    imgQuestion As Image, NumMines As Integer)
    
    With Grid
        Select Case .CellPicture
            Case imgTile 'If the cell is just a covered tile, add a flag.
                .CellPicture = imgFlag.Picture
                NumMines = NumMines - 1
            Case imgFlag 'If the cell is flagged, replace it with a question.
                .CellPicture = imgQuestion.Picture
                NumMines = NumMines + 1
            Case imgQuestion 'If the cell is questioned, reset it to a covered tile.
                .CellPicture = imgTile.Picture
        End Select
    End With

End Sub

'This GP reveals any blank and numbered cells adjacent to the blank cell that the player has clicked on.

Public Sub BlankReveal(Grid As Control, ByVal CurrentRow As Integer, ByVal CurrentCol As Integer, _
    ImgRevealed As Control, imgFlag As Control)

    With Grid
        
        'If the cell does not have a flag on, reveal it.
        
        If .CellPicture <> imgFlag.Picture Then
            .CellPicture = ImgRevealed.Picture
        End If
        
        If (CurrentCol - 1 >= 0) And (CurrentRow - 1 >= 0) Then 'Top-left (-1, -1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, -1, -1, _
                ImgRevealed, imgFlag
        End If
        If (CurrentRow - 1 >= 0) Then 'Top (0, -1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, -1, 0, _
                ImgRevealed, imgFlag
        End If
        If (CurrentCol + 1 <= .Cols - 1) And (CurrentRow - 1 >= 0) Then 'Top-right (1, -1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, -1, 1, _
                ImgRevealed, imgFlag
        End If
        If (CurrentCol + 1 <= .Cols - 1) Then 'Right (1, 0).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, 0, 1, _
                ImgRevealed, imgFlag
        End If
        If (CurrentCol + 1 <= .Cols - 1) And (CurrentRow + 1 <= .Rows - 1) Then 'Bottom-right (1, 1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, 1, 1, _
                ImgRevealed, imgFlag
        End If
        If (CurrentRow + 1 <= .Rows - 1) Then 'Bottom (0, 1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, 1, 0, _
                ImgRevealed, imgFlag
        End If
        If (CurrentCol - 1 >= 0) And (CurrentRow + 1 <= .Rows - 1) Then 'Bottom-left (-1, 1).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, 1, -1, _
                ImgRevealed, imgFlag
        End If
        If (CurrentCol - 1 >= 0) Then 'Left (-1, 0).
            BlankReveal_SetTileResetPos Grid, CurrentRow, CurrentCol, 0, -1, _
                ImgRevealed, imgFlag
        End If
    
    End With
       
End Sub

'This GP is a component of BlankReveal. It sets the offset position, checks for and reveals any
'adjacent blank or numbered cells, and resets the position back to the original cell.
'If a blank cell happens to be found, recursion is used.

Private Sub BlankReveal_SetTileResetPos(Grid As Control, CurrentRow As Integer, CurrentCol As Integer, _
    RowOffset As Integer, ColOffset As Integer, ImgRevealed As Control, imgFlag As Control)

    With Grid
        
        'Set the offset position.
        
        .Col = CurrentCol + ColOffset
        .Row = CurrentRow + RowOffset
        
        'If the cell does not have a flag on, reveal it.
        
        If .CellPicture <> imgFlag.Picture Then
        
            'If the cell is a non-blank, reveal the cell's value.
        
            If CellValue(.Col, .Row) > BLANK Then
                .CellPicture = ImgRevealed.Picture
                .Text = CellValue(.Col, .Row)
                .CellForeColor = ColourNumber(CellValue(.Col, .Row))
            End If
            
            'If the cell is a blank, use recursion.
            
            If CellValue(.Col, .Row) = BLANK And .CellPicture <> ImgRevealed.Picture Then
                BlankReveal Grid, CurrentRow + RowOffset, CurrentCol + ColOffset, _
                    ImgRevealed, imgFlag
            End If
            
        End If
        
        'Reset all offsets back to the original current cell.
        
        .Col = CurrentCol
        .Row = CurrentRow
        
    End With
    
End Sub

'This GP determines if the players has won the game.

Public Function IsGameWin(Grid As Control, imgFlag As Control) As Boolean

    Dim CCol, CRow As Integer
    Dim FalseFlags As Integer
    Dim IsWin As Boolean
    
    IsWin = True
    
    With Grid
    
    FalseFlags = 0
    For CRow = 0 To .Rows - 1
        .Row = CRow
        For CCol = 0 To .Cols - 1
            .Col = CCol
            If .CellPicture = imgFlag.Picture And CellValue(CCol, CRow) <> MINE Then
                FalseFlags = FalseFlags + 1
            End If
        Next CCol
    Next CRow
    
    End With

    If FalseFlags > 0 Then
        IsWin = False
    End If
    
    IsGameWin = IsWin

End Function

'This GP ends the game, based on if the end is a victory or a defeat.

Public Sub EndGame(Grid As Control, ByVal EndType As Integer, FaceButton As CommandButton, _
    imgFace As Variant, NumMines As Integer, GameTime As Integer, NewScore As ScoreRec)
    
    Dim FileName As String
    Dim Msg, InMsg As String
    Dim Name As String
    Dim FirstScore As Boolean
    
    FirstScore = False
    Grid.Visible = True
    
    Select Case EndType
    
        Case LOST
        
            FaceButton.Picture = imgFace(2).Picture
            
        Case WIN
            
            FaceButton.Picture = imgFace(1).Picture
            
            Msg = "You have cleared all the mines!"
            Msg = Msg & vbCrLf & "Mines cleared: " & NumMines & "."
            Msg = Msg & vbCrLf & "Time: " & GameTime & " seconds."
            MsgBox Msg, vbInformation, "Winner!"
            
            If NumRecords(LastSelectedMode) = 0 Then
                NumRecords(LastSelectedMode) = 1
                FirstScore = True
            End If
            If LastSelectedMode <> CUSTOM And _
                GameTime <= Score(LastSelectedMode, NumRecords(LastSelectedMode)).Time Then
                InMsg = "You have gotten a new high score! Enter your name [20 chars]: "
                Name = VBA.Trim$(InputBox$(InMsg, "Winner!"))
                If Name = "" Then
                    Name = "Anonymous"
                End If
                
                With NewScore
                    .Name = Name
                    .Time = GameTime
                End With
                
                If FirstScore Then
                    NumRecords(LastSelectedMode) = 0
                    FirstScore = False
                End If
                
                AddNewScoreToArray NewScore, LastSelectedMode, NumRecords(LastSelectedMode)
                FileName = App.Path & FILENAME_START & LastSelectedMode & FILENAME_EXT
                SaveFile LastSelectedMode, FileName, NumRecords(LastSelectedMode), RecordLen
            End If
            
    End Select
    
End Sub

'This function returns the respective RGB value for a non-mine and non-blank cell's font colour.

Public Function ColourNumber(Number As Integer) As Long
    
    Dim RGBValue As Long
    
    Select Case Number
        Case 1
            RGBValue = RGB(0, 0, 255)
        Case 2
            RGBValue = RGB(0, 128, 0)
        Case 3
            RGBValue = RGB(255, 0, 0)
        Case 4
            RGBValue = RGB(0, 0, 128)
        Case 5
            RGBValue = RGB(128, 0, 0)
        Case 6
            RGBValue = RGB(0, 128, 128)
        Case 7
            RGBValue = RGB(128, 0, 128)
        Case 8
            RGBValue = RGB(128, 128, 0)
        Case Else
            RGBValue = RGB(0, 0, 0)
    End Select
    
    ColourNumber = RGBValue

End Function

'This function returns the mode name given its respective integer.

Function GetModeName(Number As Integer) As String
    
    Dim Mode As String
    
    Select Case Number
        Case EASY
            Mode = "Beginner"
        Case MEDIUM
            Mode = "Intermediate"
        Case HARD
            Mode = "Expert"
        Case CUSTOM
            Mode = "Custom"
    End Select
    
    GetModeName = Mode
    
End Function

'This GP dynamically sets the form and grid size based on the number of rows and columns.

Public Sub SetFormSize(GameForm As Form, Grid As Control, Rows As Integer, Columns As Integer, _
    Face As CommandButton, Mines As Label, Time As Label)
    
    Const MARGINS = 420
    
    'Set the size of the form.
    
    With GameForm
        .Width = CELLWIDTH * Columns + MARGINS
        .Height = CELLWIDTH * Rows + MARGINS + 1500
    End With
    
    'Set the size of the gameboard.
    
    With Grid
        .Cols = Columns
        .Width = CELLWIDTH * Columns + 97
        .Rows = Rows
        .Height = CELLWIDTH * Rows + 97
    End With
    
    'Set the positions of the top controls.
    
    Face.Left = (GameForm.Width - Face.Width) / 2
    Mines.Left = (Face.Left - Mines.Width) / 2
    Time.Left = GameForm.Width - (GameForm.Width - Face.Left) / 1.8
    
End Sub

'This GP sets the number of mines, rows, and columns based on the difficulty mode selected.

Public Sub SetDifficulty(Selected As Integer, txtHeight As TextBox, txtWidth As TextBox, _
    txtMines As TextBox)

    Select Case Selected
        Case EASY
            NumMines = EASY_MINES
            Rows = 9
            Columns = 9
        Case MEDIUM
            NumMines = MEDIUM_MINES
            Rows = 16
            Columns = 16
        Case HARD
            NumMines = HARD_MINES
            Rows = 16
            Columns = 30
        Case CUSTOM
            NumMines = Val(txtMines.Text)
            Rows = Val(txtHeight.Text)
            Columns = Val(txtWidth.Text)
    End Select
    
    TotalMines = NumMines

End Sub

'This function checks if the custom settings inputs are valid.

Public Function IsCustomSettingsValid(Msg As String) As Boolean
    
    Dim NumErrors As Integer
    Dim Valid As Boolean
    
    NumErrors = 0
    Msg = "Error! Invalid settings!"
    Valid = True
    If Rows < 9 Or Rows > 18 Then
        NumErrors = NumErrors + 1
        Msg = Msg & vbCrLf & "Invalid number of rows! [9, 18]"
    End If
    If Columns < 9 Or Columns > 30 Then
        NumErrors = NumErrors + 1
        Msg = Msg & vbCrLf & "Invalid number of columns! [9, 30]"
    End If
    If NumMines < 10 Or NumMines > 500 Then
        NumErrors = NumErrors + 1
        Msg = Msg & vbCrLf & "Invalid number of mines! [10, 500]"
    End If
    If NumMines >= (Rows * Columns) Then
        NumErrors = NumErrors + 1
        Msg = Msg & vbCrLf & "More mines than empty cells!"
    End If
    
    If NumErrors > 0 Then
        Valid = False
    End If
    IsCustomSettingsValid = Valid

End Function

'This GP initializes all score records and files.

Public Sub InitializeScores(NewScore As ScoreRec)

    Dim CMode As Integer
    Dim FileName As String
    
    ClearScoreRecords
    CheckForAndCreateScoreFiles
    
    RecordLen = Len(NewScore)
    
    With NewScore
        .Name = BLANK_REC
        .Time = MAX_TIME
    End With
    
    For CMode = EASY To HARD
        FileName = App.Path & FILENAME_START & CMode & FILENAME_EXT
        ReadFile CMode, FileName, NumRecords(CMode), RecordLen
    Next CMode

End Sub

'This GP displays the scores on the high scores (frmStats) form.

Public Sub DisplayScores(Grid As Control, GameMode As Integer)

    Dim CRow As Integer
    
    With Grid
        For CRow = 1 To NumRecords(GameMode)
            .Row = CRow
            .Col = 0
            If Score(GameMode, CRow).Name <> BLANK_REC And _
                    Score(GameMode, CRow).Time <> MAX_TIME Then
                .Text = Score(GameMode, CRow).Name
                .Col = 1
                .Text = Score(GameMode, CRow).Time
            End If
        Next CRow
    End With
    
End Sub

'This GP clears and initializes all score records.

Public Sub ClearScoreRecords()

    Dim CMode As Integer
    Dim CRec As Integer
    
    For CMode = EASY To HARD
        For CRec = 1 To MAX_RECS
            With Score(CMode, CRec)
                .Name = BLANK_REC
                .Time = MAX_TIME
            End With
        Next CRec
        NumRecords(CMode) = 0
    Next CMode
    
End Sub

'This GP clears and resets the score board on the high scores (frmStats) form.

Public Sub SetClearScoreBoard(Grid As Control)
    
    Dim CRow, CCol As Integer
    
    With Grid
    
        .Clear
        
        .Row = 0
        .Col = 0
        .CellAlignment = 4
        .Text = "Name"
        .Col = 1
        .CellAlignment = 4
        .Text = "Time"
        
        For CRow = 0 To .Rows - 1
            .RowHeight(CRow) = 300
        Next CRow
        
        For CCol = 0 To .Cols - 1
            .ColWidth(CCol) = 1700
        Next CCol
        
        .Width = .Cols * .ColWidth(0) + 90
        .Height = .Rows * 308
        
    End With
    
End Sub

'This GP sorts the top 5 high scores from the fastest time to the slowest time.

Public Sub SortScores(GameMode As Integer, NumRecords As Integer)
    
    Dim J, K As Integer
    
    For K = 1 To NumRecords - 1
        For J = 1 To NumRecords - K
            If Score(GameMode, J).Time > Score(GameMode, J + 1).Time Then
                SwapScores Score(GameMode, J), Score(GameMode, J + 1)
            End If
        Next J
    Next K
    
    
End Sub

'This GP adds a new high score to its respective top 5 scores.

Public Sub AddNewScoreToArray(NewScore As ScoreRec, ByVal GameMode As Integer, _
    ByRef NumRecs As Integer)
    
    NumRecs = NumRecs + 1
    Score(GameMode, NumRecs) = NewScore
    If NumRecs > 1 Then
        SortScores GameMode, NumRecs
    End If
    
    With NewScore
        .Name = ""
        .Time = MAX_TIME
    End With
    If NumRecs = MAX_RECS Then
        Score(GameMode, NumRecs) = NewScore
        NumRecs = NumRecs - 1
    End If
    
End Sub

'This GP checks for the location of any preexisting score record files.
'If no score files exists, this GP creates these files.

Public Sub CheckForAndCreateScoreFiles()

    Dim FileName As String
    
    Dim CMode As Integer
    
    For CMode = EASY To HARD
        FileName = App.Path & FILENAME_START & CMode & FILENAME_EXT
        If Dir(FileName) = "" Then
            SaveFile CMode, FileName, MAX_DISPLAY_RECS, RecordLen
        End If
    Next CMode

End Sub
