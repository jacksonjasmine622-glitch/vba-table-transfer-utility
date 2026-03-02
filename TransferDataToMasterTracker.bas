Option Explicit

' ============================================================
' Portfolio Example (Generic):
' Transfer multi-section line items from an "InputForm" sheet
' into a master ListObject table with generic headers.
'
' Table headers expected (exact text):
' QuoteDate, SalesRep, ReferenceNumber, ContactName, CustomerName,
' MaterialType, ShippingCost, AdditionalCharge, QualitySpec1, QualitySpec2,
' Quantity, Shape, Surface, ItemType, Thickness, OuterDiameter, Width, Length,
' Finish, EdgeType, Direction, Protection, UnitPrice, TotalAmount
' ============================================================

Public Sub TransferFormToMasterTable_Generic()
    ' ---- Demo workbook identifiers (edit if needed) ----
    Const TARGET_WORKBOOK_NAME As String = "MasterTracker.xlsm"
    Const SOURCE_SHEET_NAME As String = "InputForm"
    Const TARGET_SHEET_NAME As String = "MasterData"
    Const TARGET_TABLE_NAME As String = "tblMaster"

    Dim sourceWb As Workbook, targetWb As Workbook
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim targetTable As ListObject
    Dim nextRow As ListRow

    ' Header-level values (generic)
    Dim referenceNumber As Variant, contactName As Variant, customerName As Variant
    Dim totalAmount As Variant
    Dim quoteDate As Variant, salesRep As Variant, materialType As Variant
    Dim shippingCost As Variant, additionalCharge As Variant
    Dim qualitySpec1 As Variant, qualitySpec2 As Variant

    Dim totalRowsTransferred As Long
    Dim section1Rows As Long, section2Rows As Long, section3Rows As Long, section4Rows As Long
    Dim totalDataRows As Long

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Source workbook/sheet
    Set sourceWb = ThisWorkbook
    Set sourceWs = sourceWb.Worksheets(SOURCE_SHEET_NAME)

    ' Target workbook must already be open
    Set targetWb = GetOpenWorkbook(TARGET_WORKBOOK_NAME)
    If targetWb Is Nothing Then
        MsgBox "Master workbook not open. Please open: " & TARGET_WORKBOOK_NAME, vbExclamation
        GoTo CleanExit
    End If

    Set targetWs = targetWb.Worksheets(TARGET_SHEET_NAME)

    ' Identify the master table
    Set targetTable = targetWs.ListObjects(TARGET_TABLE_NAME)
    If targetTable Is Nothing Then
        MsgBox "Table not found: " & TARGET_TABLE_NAME & " on sheet " & TARGET_SHEET_NAME, vbExclamation
        GoTo CleanExit
    End If

    If targetTable.ShowTotals Then targetTable.ShowTotals = False
    On Error Resume Next
    targetWs.Unprotect
    On Error GoTo CleanFail

    ' ------------------------------------------------------------
    ' DEMO CELL MAP (made up on purpose)
    ' Edit these ranges to match your demo InputForm layout.
    ' ------------------------------------------------------------
    quoteDate = GetMergedCellValue(sourceWs.Range("B2"))          ' QuoteDate
    salesRep = GetMergedCellValue(sourceWs.Range("B3"))           ' SalesRep
    referenceNumber = GetMergedCellValue(sourceWs.Range("B4"))    ' ReferenceNumber
    contactName = GetMergedCellValue(sourceWs.Range("B5"))        ' ContactName
    customerName = GetMergedCellValue(sourceWs.Range("B6"))       ' CustomerName
    materialType = GetMergedCellValue(sourceWs.Range("B7"))       ' MaterialType

    shippingCost = GetMergedCellValue(sourceWs.Range("E2"))       ' ShippingCost
    additionalCharge = GetMergedCellValue(sourceWs.Range("E3"))   ' AdditionalCharge
    qualitySpec1 = GetMergedCellValue(sourceWs.Range("E4"))       ' QualitySpec1
    qualitySpec2 = GetMergedCellValue(sourceWs.Range("E5"))       ' QualitySpec2

    totalAmount = GetMergedCellValue(sourceWs.Range("E7"))        ' TotalAmount (overall)

    ' ------------------------------------------------------------
    ' Line-item sections (same structure idea as your original)
    ' These are also "demo" ranges. Adjust as needed.
    '
    ' Section 1: start row 17, 9 rows (17–25)
    ' Section 2: start row 28, 3 rows (28–30)
    ' Section 3: start row 33, 5 rows (33–37)
    ' Section 4: start row 40, 5 rows (40–44)
    ' ------------------------------------------------------------

    ' Count total rows with data across all four sections
    totalDataRows = 0
    totalDataRows = totalDataRows + CountDataRows(sourceWs, 17, 8, Array("B", "C", "D", "E", "F", "H", "I", "J", "M", "N"))
    totalDataRows = totalDataRows + CountDataRows(sourceWs, 28, 2, Array("B", "C", "D", "E", "F", "I", "J", "M", "N"))
    totalDataRows = totalDataRows + CountDataRows(sourceWs, 33, 4, Array("B", "C", "D", "E", "F", "H", "I", "J", "M", "N"))
    totalDataRows = totalDataRows + CountDataRows(sourceWs, 40, 4, Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "M", "N"))

    totalRowsTransferred = 0
    section1Rows = 0: section2Rows = 0: section3Rows = 0: section4Rows = 0

    Dim rowOffset As Long, rowHasData As Boolean

    ' =========================
    ' SECTION 1 (rows 17–25)
    ' =========================
    For rowOffset = 0 To 8
        rowHasData = RowHasAnyData(sourceWs, 17, rowOffset, Array("B", "C", "D", "E", "F", "H", "I", "J", "M", "N"))
        If Not rowHasData Then GoTo NextRow1

        Set nextRow = targetTable.ListRows.Add(AlwaysInsert:=True)
        totalRowsTransferred = totalRowsTransferred + 1
        section1Rows = section1Rows + 1

        WriteHeader_Generic nextRow, targetTable, quoteDate, salesRep, referenceNumber, contactName, customerName, _
                            materialType, shippingCost, additionalCharge, qualitySpec1, qualitySpec2, _
                            totalRowsTransferred, totalDataRows, totalAmount

        ' Demo mappings based on your original structure:
        TransferValue sourceWs.Range("B17").Offset(rowOffset, 0), nextRow, targetTable, "Quantity"
        TransferValue sourceWs.Range("C17").Offset(rowOffset, 0), nextRow, targetTable, "Surface"
        TransferValue sourceWs.Range("D17").Offset(rowOffset, 0), nextRow, targetTable, "ItemType"
        TransferValue sourceWs.Range("E17").Offset(rowOffset, 0), nextRow, targetTable, "Thickness"
        TransferValue sourceWs.Range("F17").Offset(rowOffset, 0), nextRow, targetTable, "Width"
        TransferValue sourceWs.Range("H17").Offset(rowOffset, 0), nextRow, targetTable, "Length"
        TransferValue sourceWs.Range("I17").Offset(rowOffset, 0), nextRow, targetTable, "Finish"
        TransferValue sourceWs.Range("J17").Offset(rowOffset, 0), nextRow, targetTable, "EdgeType"
        TransferValue sourceWs.Range("M17").Offset(rowOffset, 0), nextRow, targetTable, "Direction"
        TransferValue sourceWs.Range("N17").Offset(rowOffset, 0), nextRow, targetTable, "Protection"
        TransferValue sourceWs.Range("U17").Offset(rowOffset, 0), nextRow, targetTable, "UnitPrice"

NextRow1:
    Next rowOffset

    ' =========================
    ' SECTION 2 (rows 28–30)
    ' =========================
    For rowOffset = 0 To 2
        rowHasData = RowHasAnyData(sourceWs, 28, rowOffset, Array("B", "C", "D", "E", "F", "I", "J", "M", "N"))
        If Not rowHasData Then GoTo NextRow2

        Set nextRow = targetTable.ListRows.Add(AlwaysInsert:=True)
        totalRowsTransferred = totalRowsTransferred + 1
        section2Rows = section2Rows + 1

        WriteHeader_Generic nextRow, targetTable, quoteDate, salesRep, referenceNumber, contactName, customerName, _
                            materialType, shippingCost, additionalCharge, qualitySpec1, qualitySpec2, _
                            totalRowsTransferred, totalDataRows, totalAmount

        TransferValue sourceWs.Range("B28").Offset(rowOffset, 0), nextRow, targetTable, "Quantity"
        TransferValue sourceWs.Range("C28").Offset(rowOffset, 0), nextRow, targetTable, "Shape"
        TransferValue sourceWs.Range("D28").Offset(rowOffset, 0), nextRow, targetTable, "ItemType"
        TransferValue sourceWs.Range("E28").Offset(rowOffset, 0), nextRow, targetTable, "Thickness"
        TransferValue sourceWs.Range("F28").Offset(rowOffset, 0), nextRow, targetTable, "OuterDiameter"
        TransferValue sourceWs.Range("I28").Offset(rowOffset, 0), nextRow, targetTable, "Finish"
        TransferValue sourceWs.Range("J28").Offset(rowOffset, 0), nextRow, targetTable, "EdgeType"
        TransferValue sourceWs.Range("M28").Offset(rowOffset, 0), nextRow, targetTable, "Direction"
        TransferValue sourceWs.Range("N28").Offset(rowOffset, 0), nextRow, targetTable, "Protection"
        TransferValue sourceWs.Range("U28").Offset(rowOffset, 0), nextRow, targetTable, "UnitPrice"

NextRow2:
    Next rowOffset

    ' =========================
    ' SECTION 3 (rows 33–37)
    ' =========================
    For rowOffset = 0 To 4
        rowHasData = RowHasAnyData(sourceWs, 33, rowOffset, Array("B", "C", "D", "E", "F", "H", "I", "J", "M", "N"))
        If Not rowHasData Then GoTo NextRow3

        Set nextRow = targetTable.ListRows.Add(AlwaysInsert:=True)
        totalRowsTransferred = totalRowsTransferred + 1
        section3Rows = section3Rows + 1

        WriteHeader_Generic nextRow, targetTable, quoteDate, salesRep, referenceNumber, contactName, customerName, _
                            materialType, shippingCost, additionalCharge, qualitySpec1, qualitySpec2, _
                            totalRowsTransferred, totalDataRows, totalAmount

        TransferValue sourceWs.Range("B33").Offset(rowOffset, 0), nextRow, targetTable, "Quantity"
        TransferValue sourceWs.Range("C33").Offset(rowOffset, 0), nextRow, targetTable, "Shape"
        TransferValue sourceWs.Range("D33").Offset(rowOffset, 0), nextRow, targetTable, "ItemType"
        TransferValue sourceWs.Range("E33").Offset(rowOffset, 0), nextRow, targetTable, "Thickness"
        TransferValue sourceWs.Range("F33").Offset(rowOffset, 0), nextRow, targetTable, "OuterDiameter"
        TransferValue sourceWs.Range("H33").Offset(rowOffset, 0), nextRow, targetTable, "Length"
        TransferValue sourceWs.Range("I33").Offset(rowOffset, 0), nextRow, targetTable, "Finish"
        TransferValue sourceWs.Range("J33").Offset(rowOffset, 0), nextRow, targetTable, "EdgeType"
        TransferValue sourceWs.Range("M33").Offset(rowOffset, 0), nextRow, targetTable, "Direction"
        TransferValue sourceWs.Range("N33").Offset(rowOffset, 0), nextRow, targetTable, "Protection"
        TransferValue sourceWs.Range("U33").Offset(rowOffset, 0), nextRow, targetTable, "UnitPrice"

NextRow3:
    Next rowOffset

    ' =========================
    ' SECTION 4 (rows 40–44)
    ' =========================
    For rowOffset = 0 To 4
        rowHasData = RowHasAnyData(sourceWs, 40, rowOffset, Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "M", "N"))
        If Not rowHasData Then GoTo NextRow4

        Set nextRow = targetTable.ListRows.Add(AlwaysInsert:=True)
        totalRowsTransferred = totalRowsTransferred + 1
        section4Rows = section4Rows + 1

        WriteHeader_Generic nextRow, targetTable, quoteDate, salesRep, referenceNumber, contactName, customerName, _
                            materialType, shippingCost, additionalCharge, qualitySpec1, qualitySpec2, _
                            totalRowsTransferred, totalDataRows, totalAmount

        TransferValue sourceWs.Range("B40").Offset(rowOffset, 0), nextRow, targetTable, "Quantity"
        TransferValue sourceWs.Range("C40").Offset(rowOffset, 0), nextRow, targetTable, "Shape"
        TransferValue sourceWs.Range("D40").Offset(rowOffset, 0), nextRow, targetTable, "ItemType"
        TransferValue sourceWs.Range("E40").Offset(rowOffset, 0), nextRow, targetTable, "OuterDiameter"
        TransferValue sourceWs.Range("F40").Offset(rowOffset, 0), nextRow, targetTable, "Thickness"
        TransferValue sourceWs.Range("G40").Offset(rowOffset, 0), nextRow, targetTable, "Width"
        TransferValue sourceWs.Range("H40").Offset(rowOffset, 0), nextRow, targetTable, "Length"
        TransferValue sourceWs.Range("I40").Offset(rowOffset, 0), nextRow, targetTable, "Finish"
        TransferValue sourceWs.Range("J40").Offset(rowOffset, 0), nextRow, targetTable, "EdgeType"
        TransferValue sourceWs.Range("M40").Offset(rowOffset, 0), nextRow, targetTable, "Direction"
        TransferValue sourceWs.Range("N40").Offset(rowOffset, 0), nextRow, targetTable, "Protection"
        TransferValue sourceWs.Range("U40").Offset(rowOffset, 0), nextRow, targetTable, "UnitPrice"

NextRow4:
    Next rowOffset

    targetWb.Save

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Transfer complete!" & vbCrLf & _
           "Section 1: " & section1Rows & " row(s)" & vbCrLf & _
           "Section 2: " & section2Rows & " row(s)" & vbCrLf & _
           "Section 3: " & section3Rows & " row(s)" & vbCrLf & _
           "Section 4: " & section4Rows & " row(s)" & vbCrLf & _
           "Total: " & totalRowsTransferred & " row(s)", vbInformation, "Done"
    Exit Sub

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbExclamation, "Transfer Failed"
End Sub

' ============================================================
' Write header-level fields (generic column names)
' ============================================================
Private Sub WriteHeader_Generic(ByVal nextRow As ListRow, ByVal targetTable As ListObject, _
                                ByVal quoteDate As Variant, ByVal salesRep As Variant, _
                                ByVal referenceNumber As Variant, ByVal contactName As Variant, _
                                ByVal customerName As Variant, ByVal materialType As Variant, _
                                ByVal shippingCost As Variant, ByVal additionalCharge As Variant, _
                                ByVal qualitySpec1 As Variant, ByVal qualitySpec2 As Variant, _
                                ByVal currentRow As Long, ByVal totalRows As Long, _
                                ByVal totalAmount As Variant)

    SetTableValue nextRow, targetTable, "QuoteDate", quoteDate
    SetTableValue nextRow, targetTable, "SalesRep", salesRep
    SetTableValue nextRow, targetTable, "ReferenceNumber", referenceNumber
    SetTableValue nextRow, targetTable, "ContactName", contactName
    SetTableValue nextRow, targetTable, "CustomerName", customerName
    SetTableValue nextRow, targetTable, "MaterialType", materialType
    SetTableValue nextRow, targetTable, "ShippingCost", shippingCost
    SetTableValue nextRow, targetTable, "AdditionalCharge", additionalCharge
    SetTableValue nextRow, targetTable, "QualitySpec1", qualitySpec1
    SetTableValue nextRow, targetTable, "QualitySpec2", qualitySpec2

    If currentRow = totalRows Then
        SetTableValue nextRow, targetTable, "TotalAmount", totalAmount
    End If
End Sub

' ============================================================
' Transfer one cell value into a named table column
' ============================================================
Private Sub TransferValue(ByVal sourceRange As Range, ByVal nextRow As ListRow, _
                          ByVal targetTable As ListObject, ByVal columnName As String)
    Dim v As Variant
    v = GetMergedCellValue(sourceRange)

    If Not IsEmptyOrZero(v) Then
        SetTableValue nextRow, targetTable, columnName, v
    End If
End Sub

' ============================================================
' Safe table write by column header name
' ============================================================
Private Sub SetTableValue(ByVal nextRow As ListRow, ByVal targetTable As ListObject, _
                          ByVal columnName As String, ByVal v As Variant)
    Dim colIndex As Long
    If IsEmptyOrZero(v) Then Exit Sub

    On Error Resume Next
    colIndex = targetTable.ListColumns(columnName).Index
    On Error GoTo 0

    If colIndex > 0 Then nextRow.Range(1, colIndex).Value = v
End Sub

' ============================================================
' Count rows with any data across a set of columns
' ============================================================
Private Function CountDataRows(ByVal ws As Worksheet, ByVal startRow As Long, _
                               ByVal rowCount As Long, ByVal columns As Variant) As Long
    Dim rowOffset As Long, col As Variant
    Dim hasData As Boolean
    Dim count As Long

    count = 0
    For rowOffset = 0 To rowCount
        hasData = False
        For Each col In columns
            If Not IsEmptyOrZero(ws.Range(CStr(col) & startRow).Offset(rowOffset, 0).Value) Then
                hasData = True
                Exit For
            End If
        Next col
        If hasData Then count = count + 1
    Next rowOffset

    CountDataRows = count
End Function

' ============================================================
' Row-level "any data" check
' ============================================================
Private Function RowHasAnyData(ByVal ws As Worksheet, ByVal baseRow As Long, _
                               ByVal rowOffset As Long, ByVal cols As Variant) As Boolean
    Dim c As Variant
    For Each c In cols
        If Not IsEmptyOrZero(ws.Range(CStr(c) & baseRow).Offset(rowOffset, 0).Value) Then
            RowHasAnyData = True
            Exit Function
        End If
    Next c
    RowHasAnyData = False
End Function

' ============================================================
' Merged-cell safe value read
' ============================================================
Private Function GetMergedCellValue(ByVal rng As Range) As Variant
    If rng.MergeCells Then
        GetMergedCellValue = rng.MergeArea.Cells(1, 1).Value
    Else
        GetMergedCellValue = rng.Value
    End If
End Function

' ============================================================
' Empty/zero check
' ============================================================
Private Function IsEmptyOrZero(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyOrZero = True
    ElseIf IsNumeric(v) Then
        IsEmptyOrZero = (CDbl(v) = 0)
    Else
        IsEmptyOrZero = (Trim$(CStr(v)) = vbNullString)
    End If
End Function

' ============================================================
' Find an already-open workbook by name
' ============================================================
Private Function GetOpenWorkbook(ByVal wbName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, wbName, vbTextCompare) = 0 Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOpenWorkbook = Nothing
End Function