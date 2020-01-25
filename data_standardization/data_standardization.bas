' Dependcies: 
'   lib_general.bas

Option Explicit


Public Sub ToUCase()
	' Changes the selected cells to Upper case
	
	Const strInputTitle = "Upper Case"
	Dim intFromCol As Integer, intToCol As Integer
	Dim intFromRow As Integer, intToRow As Integer
	Dim intTempCol As Integer, intTempRow As Integer

    ' Gets inputs from the user
    intFromCol = UCase(ToNumber(InputBox("Type the starting column", strInputTitle, "A")))
    intToCol = UCase(ToNumber(InputBox("Type the ending column", strInputTitle, "A")))
    intFromRow = InputBox("Type the starting row", strInputTitle, 1)
    intToRow = InputBox("Type the ending row", strInputTitle, 100)
    
	' Converts to upper case
    For intTempCol = intFromCol To intToCol
        For intTempRow = intFromRow To intToRow
            ActiveSheet.Range(ToLetter(intTempCol) & intTempRow).Value = UCase(ActiveSheet.Range(ToLetter(intTempCol) & intTempRow).Value)
        Next intTempRow
    Next intTempCol

End Sub


Public Sub StandartLength()

	' Converts data to have same length, filling whith a character specified
	
	Const strInputTitle = "Standart Length"
	Const strDataLossSheetName = "DataLoss"
	Dim strDataMainCol As String, strDataLossPasteCol As String, strDataLossAAPasteCol As String, intStartingRow As Integer
	Dim intLenAchieve As Integer, strFillingChar As String

	Dim strDataMainSheetName As String
	Dim intRows As Integer, intRowTemp As Integer
	Dim intMaxLen As Integer

	Dim intCharsBalance As Integer
	Dim blnAllowDataLoss As Boolean, blnAllowDataLossSpecified As Boolean, msgRespAllowDataLoss As VbMsgBoxResult
	Dim blnCollectDataLoss As Boolean, blnCollectDataLossSpecified As Boolean, msgRespCollectDataLoss As VbMsgBoxResult
	Dim intDataLossRowToPaste As Integer
    
    ' Inputs from the user
    strDataMainCol = UCase(InputBox("Type the main data column", strInputTitle, "A"))
    intStartingRow = CInt(InputBox("Type starting row", strInputTitle, 1))
    intLenAchieve = CInt(InputBox("Type new data length", strInputTitle, 5))
    strFillingChar = InputBox("Type the filling character. (1 character)", strInputTitle, "0")
    
    ' Data validations
    ' Checks if the new length is numeric.
    If Not IsNumeric(intLenAchieve) Then
        Call ReInitializeAndClose("The length must be numeric")
    End If
    ' Checks if the new length is greater than zero.
    If Not intLenAchieve > 0 Then
        Call ReInitializeAndClose("The length must be greter than zero")
    End If
    ' Checks if the filling character is 1 character only.
    If Not Len(strFillingChar) = 1 Then
        Call ReInitializeAndClose("Filling character must be 1 character only.")
    End If
    
    ' Variable initialization
    blnAllowDataLoss = False
    blnAllowDataLossSpecified = False
    blnCollectDataLoss = False
    blnCollectDataLossSpecified = False
    intDataLossRowToPaste = 0
    Call DeleteSheet(strDataLossSheetName)
    
    ' Gets the name of the main sheet.
    strDataMainSheetName = ActiveSheet.Name
    
    ' Counts the full cells.
    intRows = CountFullRows(strDataMainCol, intStartingRow)
    If Not intRows > 0 Then
        Call ReInitializeAndClose("intRows is not greater than zero. Data not found on column " & strDataMainCol)
    End If
        
    ' Converts the main data column to text format.
    Worksheets(strDataMainSheetName).Columns(strDataMainCol & ":" & strDataMainCol).NumberFormat = "@"
   
    ' Fills the value with the filling character or trims if from the left.
    For intRowTemp = intStartingRow To (intStartingRow + intRows - 1)
        
		' Compares the value of the cell with the length specified by the user.
        intCharsBalance = intLenAchieve - Len(Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value)
        
        ' Cases for the positive or negative balance. In case of zero nothing should be done.
        If intCharsBalance > 0 Then
            Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value = String(intCharsBalance, strFillingChar) & _
                                                                                    Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value
        ElseIf intCharsBalance < 0 Then
            
			' If the user has not decided yet if he wants to lose data bu trimming to the length specified.
            If blnAllowDataLossSpecified = False Then
                msgRespAllowDataLoss = MsgBox("Some data have greater length than the length specified." & vbNewLine & _
                                                "Do you allow data loss?", vbQuestion + vbYesNo, "Entry")
                blnAllowDataLoss = IIf(msgRespAllowDataLoss = vbYes, True, False)
                blnAllowDataLossSpecified = True
            End If
            
            ' If the user has decided to allow data loss or not.
            If blnAllowDataLossSpecified = True Then
                
				' If the user has not decided yet if he wants to Collect data will be trimmed.
                If blnCollectDataLossSpecified = False Then
                   
				   ' User must specify if he wants to collect data loss or not.
                    blnCollectDataLossSpecified = True
                    msgRespCollectDataLoss = MsgBox("Do you want to collect the rows containing values with bigger length?", vbQuestion + vbYesNo, "Entry")
                    blnCollectDataLoss = IIf(msgRespCollectDataLoss = vbYes, True, False)
                    
                    If blnCollectDataLoss = True Then
                        
						' Inputs from the user
                        strDataLossAAPasteCol = UCase(InputBox("Type autonumbering column for data loss", strInputTitle, "A"))
                        strDataLossPasteCol = UCase(InputBox("Type data loss column", strInputTitle, "B"))
                        
						' Adds the data-loss worksheet
                        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = strDataLossSheetName
                        Worksheets(strDataMainSheetName).Activate
                    End If
                End If
                
                ' Copy-paste the values from the cells with greater length to the sheet DataLosss. A column with the AA of the main data row is added.
                If blnCollectDataLoss = True Then
                    intDataLossRowToPaste = intDataLossRowToPaste + 1
                    Worksheets(strDataLossSheetName).Columns(strDataLossPasteCol & ":" & strDataLossPasteCol).NumberFormat = "@"
                    Worksheets(strDataLossSheetName).Range(strDataLossPasteCol & intDataLossRowToPaste).Value = _
                        Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value
                    Worksheets(strDataLossSheetName).Range(strDataLossAAPasteCol & intDataLossRowToPaste).Value = intRowTemp
                End If
                
                ' If the user has accepted the data loss.
                If blnAllowDataLoss = True Then
                    Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value = Right(Worksheets(strDataMainSheetName).Range(strDataMainCol & intRowTemp).Value, intLenAchieve)
                End If
                
            End If
        End If
    Next intRowTemp
End Sub


Public Sub MergeCells()
	' Merges the values of two cells
	
	Const strInputTitle = "Merge Cells"
	Dim msgResponseActiveCell As VbMsgBoxResult
	Dim strMergeChar As String, blnMergeFromRight As Boolean, blnClearForeignCells As Boolean, strStartCell As String
	Dim strRangeSplit() As String
	Dim strCol(1 To 3) As String '  1=Left, 2=Main, 3=Right
	Dim intFirstRow As Integer, intLastRow As Integer, intRows As Integer
	Dim strForeignCell As String
	Dim intTempRow As Integer, strTempValue As String
    
    ' Gets user' s selected cell
    strStartCell = ActiveCell.AddressLocal

    ' Column-variables preparation 1=Left, 2=Main, 3=Right
    strRangeSplit() = Split(strStartCell, "$")
    strCol(2) = strRangeSplit(1)
    strCol(1) = ToLetter(ToNumber(strCol(2)) - 1)
    strCol(3) = ToLetter(ToNumber(strCol(2)) + 1)
    
    ' Row-variables preparation
    intFirstRow = CInt(strRangeSplit(2))
    
	' Counts the full rows
    intRows = CountFullRows(strCol(2), intFirstRow)
    intLastRow = intFirstRow + intRows - 1
    
    ' Asks the user if the selected cell is correct
    msgResponseActiveCell = MsgBox("Active cell is " & strCol(2) & intFirstRow & " " & vbNewLine & _
                                    strCol(2) & " is the merging column" & vbNewLine & _
                                    intFirstRow & " is the first row with data" & vbNewLine & _
                                    "Are all theese correct?", vbQuestion + vbYesNo, strInputTitle)
    If msgResponseActiveCell = vbNo Then Exit Sub
    
    ' Gets the inputs from the user
    blnMergeFromRight = CBool(InputBox("Which column to Merge?" & vbNewLine & _
                                "Type 0 for " & strCol(1) & " column" & vbNewLine & _
                                "Type 1 for " & strCol(3) & " column", strInputTitle, 1))
    strMergeChar = InputBox("Tyre the text that will be between the merged data", strInputTitle, " - ")
    blnClearForeignCells = CBool(InputBox("Should data on " & IIf(blnMergeFromRight = True, strCol(3), strCol(1)) & " be cleared?" & vbNewLine & _
                                "Type 1 for Yes" & vbNewLine & _
                                "Type 0 for No", strInputTitle, 0))
    
    ' Checks if the merge is from left column and the user selected A column...
    If blnMergeFromRight = True And (strStartCell = CellMove(strStartCell, 3)) Then
        Call ReInitializeAndClose("MergeCells. Cannot select A column with left merge")
    End If
    
    ' Gets the value from the foreign cell and merge it to the native.
    For intTempRow = intFirstRow To intLastRow
        If blnMergeFromRight = False Then
            strForeignCell = CellMove("$" & strCol(2) & "$" & intTempRow, 3) ' 3 means left cell movement direction
            strTempValue = ActiveSheet.Range(strForeignCell).Value
            ActiveSheet.Range(strCol(2) & intTempRow).Value = strTempValue & _
                                                                    strMergeChar & _
                                                                    ActiveSheet.Range(strCol(2) & intTempRow).Value
        ElseIf blnMergeFromRight = True Then
            strForeignCell = CellMove("$" & strCol(2) & "$" & intTempRow, 4) ' 4 means right cell movement direction
            strTempValue = ActiveSheet.Range(strForeignCell).Value
            ActiveSheet.Range(strCol(2) & intTempRow).Value = ActiveSheet.Range(strCol(2) & intTempRow).Value & _
                                                                    strMergeChar & _
                                                                    strTempValue
        End If
        
        ' Clears the data from the foreign cell depending on user' s decision.
        If blnClearForeignCells = True Then
            ActiveSheet.Range(strForeignCell).Value = ""
        End If
    Next intTempRow
End Sub


Public Sub SplitCells()
	' Splits the values from a cell

	Const strInputTitle = "Split Cells"
	Dim intRows As Integer, intFirstRow As Integer, intLastRow As Integer
	Dim msgResponseActiveCell As VbMsgBoxResult
	Dim strStartCell As String, strCol As String, strRangeSplit() As String
	Dim msgResponseFromRightToLeft As VbMsgBoxResult, blnFromRightToLeft As Boolean
	Dim msgResponseKeepSplitString As VbMsgBoxResult, blnKeepSplitString As Boolean
	Dim strSplitString As String

	Dim intTempRow As Integer, strTempCellValue As String
	Dim intTempChar As Integer, strTempString As String
	Dim blnSplitStringFound As Boolean
	Dim intSplitPosition As Integer
	Dim strStringTable(1 To 3) As String
    
    ' Gets the address from the user' s active cell
    strStartCell = ActiveCell.AddressLocal
        
    ' Column-variable preparation
    strRangeSplit = Split(strStartCell, "$")
    strCol = strRangeSplit(1)
    
    ' Row-variable preparation
    intFirstRow = strRangeSplit(2)
    
	' Counts the full rows.
    intRows = CountFullRows(strCol, intFirstRow)
    intLastRow = intFirstRow + intRows - 1
    
   ' Asks the user if the selected cell is correct
    msgResponseActiveCell = MsgBox("Active cell is " & strCol & intFirstRow & vbNewLine & _
                                    strCol & " is the spliting column" & vbNewLine & _
                                    intFirstRow & " is the first row with data" & vbNewLine & _
                                    "Are all theese correct?", vbQuestion + vbYesNo, strInputTitle)
    If msgResponseActiveCell = vbNo Then Exit Sub
    
    ' Gets the searching direction from the user
    msgResponseFromRightToLeft = MsgBox("Give the split search direction. " & vbNewLine & _
                                         "Press Yes for Right to Left" & vbNewLine & _
                                         "Press No for Left to Right", vbQuestion + vbYesNo, strInputTitle)
    blnFromRightToLeft = IIf(msgResponseFromRightToLeft = vbYes, True, False)
    
	' The string to be searche that will define the split point
    strSplitString = InputBox("Type the split string", strInputTitle, "-")
    
	' Asking if the original column will be affected or not
    msgResponseKeepSplitString = MsgBox("Do you want to include the splitting string to the original column?", _
                                        vbQuestion + vbYesNo, strInputTitle)
    blnKeepSplitString = IIf(msgResponseKeepSplitString = vbYes, True, False)
    
        
    ' For all the rows
    For intTempRow = intFirstRow To intLastRow
        ' Gets each row' s value
        strTempCellValue = ActiveSheet.Range(strCol & intTempRow)
		
        ' Initialization for the results
        blnSplitStringFound = False
        intSplitPosition = 0
        
        ' Searching the temporary value from right to left
        ' Loops all the characters into the value
        For intTempChar = 1 To Len(strTempCellValue)
            
			' Waits until the temp string get the same length as the string to be searched
            If (intTempChar >= Len(strSplitString)) Then
                If blnFromRightToLeft = True Then
                    If Left(Right(strTempCellValue, intTempChar), Len(strSplitString)) = strSplitString Then
                        blnSplitStringFound = True
                        Exit For
                    End If
                ElseIf blnFromRightToLeft = False Then
                    If Right(Left(strTempCellValue, intTempChar), Len(strSplitString)) = strSplitString Then
                        blnSplitStringFound = True
                        Exit For
                    End If
                End If
            End If
        Next intTempChar
        
        If blnFromRightToLeft = True Then
            
			' The result is the split position
            intSplitPosition = Len(strTempCellValue) - intTempChar + 1
        ElseIf blnFromRightToLeft = False Then
            
			' The result is the split position
            intSplitPosition = intTempChar - Len(strSplitString) + 1
        End If
        
        ' Splits the values if the string was found into the value
        If blnSplitStringFound = True Then
            
			' Splits the original value to the left piece the searchString and the right piece
            strStringTable(1) = Left(strTempCellValue, (intSplitPosition - 1))
            strStringTable(2) = Right(Left(strTempCellValue, (intSplitPosition + Len(strSplitString) - 1)), Len(strSplitString))
            strStringTable(3) = Right(strTempCellValue, (Len(strTempCellValue) - intSplitPosition - Len(strSplitString) + 1))
            
            ' Writes into the original cell and the right of it the specific vales, 1&2 - 3  or  1 - 2&3, depending on the keep-split user' s decision
            If blnKeepSplitString = True Then
                ActiveSheet.Range(strCol & intTempRow) = strStringTable(1) & strStringTable(2)
                ActiveSheet.Range(ToLetter(ToNumber(strCol) + 1) & intTempRow) = strStringTable(3)
            Else
                ActiveSheet.Range(strCol & intTempRow) = strStringTable(1)
                ActiveSheet.Range(ToLetter(ToNumber(strCol) + 1) & intTempRow) = strStringTable(2) & strStringTable(3)
            End If
        End If
    Next intTempRow
End Sub


Private Sub ReInitializeAndClose(strMessage As String, Optional strSheetToActivate As String, Optional strCellToActivate As String)
	' Shows a message, Reinitializes the excel and ends which is needed after an error occurs
	
    MsgBox strMessage, vbCritical, "Error"
    On Error Resume Next
    Call DeleteSheet("SheetNameTemp")
    
    Worksheets(strSheetToActivate).Activate
    ActiveSheet.Range(strCellToActivate).Activate
    End
End Sub


