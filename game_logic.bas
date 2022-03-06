Attribute VB_Name = "Module1"
Public Player As Boolean
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
Sub CreateGrid()
    Set OldSelection = ActiveCell
    'Game grid setup
    Range("B3:D5").Select
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
    Set iRange = Range("B3:D5")
    For Each iCells In iRange
        iCells.BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
    Next iCells

    Call SetDefaultPlayer
    'Revert the selection
    OldSelection.Select
End Sub
Sub MarkField()
    'Check if the selected cell is in the game's grid
    Set IsInGrid = Application.Intersect(Range("B3:D5"), ActiveCell)
    If IsInGrid Is Nothing Then
        MsgBox "Selected cell is not on the game board!"
    Else
        If CheckCellAvailability Then
            Call SetPlayerCell
            If GameEndCondition Then
                'Restart the game if it's over
                Call CreateGrid
            End If
            Player = Not Player
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
Function GameEndCondition() As Boolean
    'Parts of the grid that need to be checked
    Dim Combinations(5) As Range
    Set Combinations(0) = Range("B3:B5")
    Set Combinations(1) = Range("C3:C5")
    Set Combinations(2) = Range("D3:D5")
    Set Combinations(3) = Range("B3:D3")
    Set Combinations(4) = Range("B4:D4")
    Set Combinations(5) = Range("B5:D5")
    'TODO: add diagonal checks

    FreeSpace = False
    GameEndCondition = False
    For Each combination In Combinations
        XCount = 0
        OCount = 0
        For Each cell In combination.Cells
            If cell.Value = "X" Then
                XCount = XCount + 1
            ElseIf cell.Value = "O" Then
                OCount = OCount + 1
            Else
                FreeSpace = True
            End If
        Next cell
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
    Next combination
    'Check if we ran out of space on the grid
    If Not FreeSpace And Not GameEndCondition Then
        MsgBox "Tie!"
        GameEndCondition = True
    End If
End Function
