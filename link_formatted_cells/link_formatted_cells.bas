' Dependcies: 
'   lib_performance.bas

Const MAX_ROWS = 1000000

Private Sub CopyFormattedButton_Click()

    Call CopyFormatted
    
End Sub

Private Sub CopyFormatted()
    ' Looks-up the destination id into the source look-up column to retrieve the row number
    ' Then it copies the source cell into the destination cell
    ' This is done to copy the format and the within cell new lines

    ' Layout dependent for the Destination
    dstWsName = "sheet1"
    dstFirstRow = 2
    dstIdCol = "A"
    dstWriteCol = "B"
    
    ' Layout dependent for the Source
    srcWsName = "sheet1"
    srcFirstRow = 2
    srcLookupCol = "D"
    srcReadCol = "E"

    Call performancePre
    
    Call lookUpCell(dstWsName, dstFirstRow, dstIdCol, dstWriteCol, _
                    srcWsName, srcFirstRow, srcLookupCol, srcReadCol)
                    
    Call performancePost
    
End Sub

Private Sub lookUpCell(dstWsName, dstFirstRow, dstIdCol, dstWriteCol, _
                       srcWsName, srcFirstRow, srcLookupCol, srcReadCol)
    ' Reads a value in

    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    
    Set srcWs = ActiveWorkbook.Sheets(srcWsName)
    Set dstWs = ActiveWorkbook.Sheets(dstWsName)

    Dim sourceIdsVector As Range
    Set sourceIdsVector = srcWs.Range(srcLookupCol & srcFirstRow & ":" & srcLookupCol & MAX_ROWS)
    
    ' Initialization
    dstWriteRow = dstFirstRow
    Do
        srcRow = Empty
        searchId = dstWs.Range(dstIdCol & dstWriteRow).Value
        
        ' Make sure the id is not empty
        If searchId = vbNullString Then Exit Do
        
        ' Lookup the id to find the row number
        For Each cell In sourceIdsVector.Cells
            If cell.Value = "" Then Exit For
            
            If cell.Value = searchId Then
                srcRow = cell.Row
                Exit For
            End If
        Next cell
            
        ' If the search succeeds id does the copy paste of the cells.
        If srcRow <> Empty Then

            Dim srcCell As Range
            Set srcCell = srcWs.Range(srcReadCol & srcRow)
            
            Dim dstCell As Range
            Set dstCell = dstWs.Range(dstWriteCol & dstWriteRow)
            
            Call CopyPasteRange(srcWs, srcCell, dstWs, dstCell)
        
        End If
        
        ' Update
        dstWriteRow = dstWriteRow + 1
    Loop

End Sub


Private Sub CopyPasteRange(srcWs As Worksheet, srcRange As Range, dstWs As Worksheet, dstRange As Range)
    ' Copy a ranges and pastes it to another
    srcWs.Select
    srcRange.Select
    Selection.Copy
    
    dstWs.Select
    dstRange.Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False

End Sub
