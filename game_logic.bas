Attribute VB_Name = "Module1"
Public Const GridRange As String = "B3:D5"
Public Const CurrentPlayerDisplayCell As String = "G12:H12"
Dim Combinations(7, 2) As String 'Parts of the grid that need to be checked
Public Player As Boolean

Function InitCombinations()
    'Note: it's still hardcoded, but at least it's in one place
    'Main diagonal
    Combinations(0, 0) = "B3"
    Combinations(0, 1) = "C4"
    Combinations(0, 2) = "D5"
    'Anti diagonal
    Combinations(1, 0) = "D3"
    Combinations(1, 1) = "C4"
    Combinations(1, 2) = "B5"
    'Rows
    Combinations(2, 0) = "B3"
    Combinations(2, 1) = "C3"
    Combinations(2, 2) = "D3"
    
    Combinations(3, 0) = "B4"
    Combinations(3, 1) = "C4"
    Combinations(3, 2) = "D4"
    
    Combinations(4, 0) = "B5"
    Combinations(4, 1) = "C5"
    Combinations(4, 2) = "D5"
    'Cols
    Combinations(5, 0) = "B3"
    Combinations(5, 1) = "B4"
    Combinations(5, 2) = "B5"
    
    Combinations(6, 0) = "C3"
    Combinations(6, 1) = "C4"
    Combinations(6, 2) = "C5"
    
    Combinations(7, 0) = "D3"
    Combinations(7, 1) = "D4"
    Combinations(7, 2) = "D5"
    
End Function

Sub SetDefaultPlayer()
    Player = True
End Sub

Sub MoveUp()
    Selection.Offset(-1, 0).Select
End Sub

Sub MoveDown()
    Selection.Offset(1, 0).Select
End Sub

Sub MoveRight()
    Selection.Offset(0, 1).Select
End Sub

Sub MoveLeft()
    Selection.Offset(0, -1).Select
End Sub

Function CreateGrid()
    Set OldSelection = ActiveCell
    'Game grid setup
    Range(GridRange).Select
    With Selection
        .Interior.Color = vbGreen
        .ColumnWidth = 10
        .RowHeight = 40
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = vbRed
        .Font.Size = 48
        .Value = ""
    End With
    'Add borders
    Dim iRange As Range
    Dim iCells As Range
    Set iRange = Range(GridRange)
    For Each iCells In iRange
        iCells.BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
    Next iCells
    'Revert the selection
    OldSelection.Select
End Function

Sub MarkField()
    'Check if the selected cell is in the game's grid
    Set IsInGrid = Application.Intersect(Range(GridRange), ActiveCell)
    If IsInGrid Is Nothing Then
        MsgBox "Selected cell is not on the game board!"
    Else
        If CheckCellAvailability Then
            Call SetPlayerCell
            If GameEndCondition Then
                'Restart the game if it's over
                Call ResetBoardState
            Else
                Player = Not Player
                Call DisplayCurrentPlayer
            End If
        End If
    End If
End Sub

Function CheckCellAvailability() As Boolean
    If ActiveCell.Value = "" Then
        CheckCellAvailability = True
    Else
        MsgBox "Selected cell is already occupied!"
        CheckCellAvailability = False
    End If
End Function

Function SetPlayerCell()
    If Player Then
        ActiveCell.Value = "X"
    Else
        ActiveCell.Value = "O"
    End If
End Function

Function DisplayCurrentPlayer()
    If Player Then
        Range(CurrentPlayerDisplayCell).Value = "X"
    Else
        Range(CurrentPlayerDisplayCell).Value = "O"
    End If
End Function

Sub ResetBoardState()
    Call SetDefaultPlayer
    Call DisplayCurrentPlayer
    Call CreateGrid
End Sub

Function GameEndCondition() As Boolean
    Call InitCombinations
    FreeSpace = False
    GameEndCondition = False
    For i = LBound(Combinations, 1) To UBound(Combinations, 1)
        XCount = 0
        OCount = 0
        For j = LBound(Combinations, 2) To UBound(Combinations, 2)
            If Range(Combinations(i, j)).Value = "X" Then
                XCount = XCount + 1
            ElseIf Range(Combinations(i, j)).Value = "O" Then
                OCount = OCount + 1
            Else
                FreeSpace = True
            End If
        Next j
        'Check if one of the players won
        If XCount = 3 Then
            MsgBox "X wins!"
            GameEndCondition = True
            Exit For
        ElseIf OCount = 3 Then
            MsgBox "O wins!"
            GameEndCondition = True
            Exit For
        End If
    Next i
    'Check if we ran out of space on the grid
    If Not FreeSpace And Not GameEndCondition Then
        MsgBox "Tie!"
        GameEndCondition = True
    End If
End Function
