Private Function ToNumber(strCol As String) As Integer
	' Translates the column letter to the column number.
	
	Dim StringHasOneLetter As Boolean

    ToNumber = 0
    StringHasOneLetter = False
    If Len(strCol) = 1 Then
        strCol = "0" & strCol
        StringHasOneLetter = True
    End If
    ToNumber = ToNumber + (Asc(Right(strCol, 1)) - 64)
    ToNumber = ToNumber + IIf(Left(strCol, 1) <> "0", ((Asc(Left(strCol, 1)) - 64) * 26), 0)
    If StringHasOneLetter = True Then
        strCol = Right(strCol, 1)
    End If
End Function


Private Function ToLetter(intCol As Integer) As String
	' Translates the column number to the column letter
	
	Dim intTwentySixes As Integer, intUnits As Integer
	Dim strFirst As String, strSecond As String

    If intCol > 26 Then
        DoEvents
    End If
    intTwentySixes = (intCol - 1) \ 26
    intUnits = intCol - (26 * intTwentySixes)
    strFirst = IIf(intTwentySixes <> 0, Chr(intTwentySixes + 64), "")
    strSecond = Chr(intUnits + 64)
    ToLetter = strFirst & strSecond
End Function


Private Sub CellMoveDisplay(Direction As Integer)
	' Moves the active cell to a given direction.
	
	Dim strTempCell As String
	Dim strRangeSplit() As String

    strTempCell = ActiveCell.AddressLocal
    strRangeSplit() = Split(strTempCell, "$")
    Select Case Direction
        Case 1  ' Up
            If RangeSplit(2) > 1 Then
                NewRange = "$" & RangeSplit(1) & "$" & (RangeSplit(2) - 1)
            Else
                Exit Sub
            End If
        Case 2  ' Down
            NewRange = "$" & RangeSplit(1) & "$" & (RangeSplit(2) + 1)
        Case 3  ' Left
            If Len(RangeSplit(1)) = 1 Then
                NewRange = "$" & Chr(Asc(RangeSplit(1)) + 1) & "$" & RangeSplit(2)
            Else
                Exit Sub
            End If
        Case 4  ' Right
            If (Len(RangeSplit(1)) = 1) And (RangeSplit(1) <> "A") Then
                NewRange = "$" & Chr(Asc(RangeSplit(1)) - 1) & "$" & RangeSplit(2)
            Else
                Exit Sub
            End If
    End Select

    ActiveSheet.Range(NewRange).Select
End Sub


Private Function CellMove(OldCell As String, intDirection As Integer) As String

	'Moves to the next cell to a direction given.
	Dim RangeSplit() As String

    RangeSplit = Split(OldCell, "$")
    Select Case intDirection
        Case 1  ' Up
            If RangeSplit(2) > 1 Then
                RangeSplit(2) = RangeSplit(2) - 1
            End If
        Case 2  ' Down
            RangeSplit(2) = RangeSplit(2) + 1
        Case 3  ' Left
            If ToNumber(RangeSplit(1)) > 1 Then
                RangeSplit(1) = ToLetter(ToNumber(RangeSplit(1)) - 1)
            End If
        Case 4  ' Right
            RangeSplit(1) = ToLetter(ToNumber(RangeSplit(1)) + 1)
    End Select
    CellMove = "$" & RangeSplit(1) & "$" & RangeSplit(2)
End Function


Private Function CountFullRows(strCol As String, intFirstFullRow As Integer) As Integer
	'Counts the number of the full rows starting from a row on a specified column.
	
	Dim intRow As Integer
    CountFullRows = 0
    intRow = 0
    Do While ActiveSheet.Range(strCol & (intFirstFullRow + intRow)).Value <> ""
        intRow = intRow + 1
    Loop
    CountFullRows = intRow
End Function


Private Function CountMaxLen(strFromCol As String, strToCol As String, strFromRow As Integer, strToRow As Integer) As Integer
	'Returns the maximun length of the values within a range
	
	Dim inColTemp As Integer, intRowTemp As Integer
	Dim intFromCol As Integer, intToCol As Integer

    intFromCol = ToNumber(strFromCol)
    intToCol = ToNumber(strToCol)
    
    CountMaxLen = 0
    For inColTemp = intFromCol To intToCol
        For intRowTemp = strFromRow To strToRow
            If Len(ActiveSheet.Range(ToLetter(inColTemp) & intRowTemp).Value) > CountMaxLen Then
                CountMaxLen = Len(ActiveSheet.Range(ToLetter(inColTemp) & intRowTemp).Value)
            End If
        Next intRowTemp
    Next inColTemp
End Function


Private Sub CopySheet(strName As String)
	' Copies a sheet giving a specific name
	
	Application.ScreenUpdating = False
	ActiveWorkbook.Sheets.Add.Name = strName
	Cells.Copy
	ActiveWorkbook.Sheets(strName).Range("A1").PasteSpecial
	Application.CutCopyMode = False
	Application.ScreenUpdating = True
End Sub


Private Sub DeleteSheet(strSheetName As String)
	'Deletes a sheet with a specific name

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(strSheetName).Delete
    Application.DisplayAlerts = True
End Sub
