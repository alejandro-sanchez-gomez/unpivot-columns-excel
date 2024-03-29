Sub unpivot_columns()
Attribute unpivot_columns.VB_Description = "For more information, check Options...\r\n\r\nDocumentation and usage:\r\nhttps://github.com/Levantino-Engineering/unpivot-columns-excel/tree/main"
Attribute unpivot_columns.VB_ProcData.VB_Invoke_Func = "P\n14"

    Dim width, height As Integer
    width = WorksheetFunction.CountA(ActiveSheet.Range("1:1"))
    height = WorksheetFunction.CountA(ActiveSheet.Range("A:A"))
    
    Dim ColLetter, limitCell As String
    ColLetter = Split(Cells(1, width + 1).Address, "$")(1)
    limitCell = "B2:" & ColLetter & (height + 1)
    
    Dim leftCells, innerCells, headerCells As Variant
    leftCells = ActiveSheet.Range("A:A")
    innerCells = ActiveSheet.Range(limitCell)
    headerCells = ActiveSheet.Range("1:1")

    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)

    Dim nLeft, nInner, nHeader, rowIndex As Integer
    nLeft = 2
    nInner = 1
    nHeader = 2
    rowIndex = 1
    
    Dim loopCellLeft, loopCellInner, loopCellHeader As String
    loopCellLeft = ""
    loopCellInner = ""
    loopCellHeader = ""
    
    For i = 1 To height
        For j = 1 To width
            loopCellLeft = "A" & rowIndex
            loopCellInner = "B" & rowIndex
            loopCellHeader = "C" & rowIndex
            
            ActiveSheet.Range(loopCellLeft) = leftCells(nLeft, 1)
            ActiveSheet.Range(loopCellInner) = innerCells(i, nHeader - 1)
            ActiveSheet.Range(loopCellHeader) = headerCells(1, nHeader)
            
            nInner = nInner + 1
            nHeader = nHeader + 1
            rowIndex = rowIndex + 1
        Next
        nLeft = nLeft + 1
        nInner = 1
        nHeader = 2
    Next
End Sub
