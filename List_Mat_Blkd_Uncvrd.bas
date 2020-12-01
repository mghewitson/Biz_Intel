Attribute VB_Name = "List_Mat_Blkd_Uncvrd"
Sub List_Mat_Blkd_Uncvrd()
Attribute List_Mat_Blkd_Uncvrd.VB_ProcData.VB_Invoke_Func = " \n14"
'
' List_Mat_Blkd_Uncvrd Macro
'

'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    Range("J5") = "Quantity"
    
    Range("1:4,6:6").Delete Shift:=xlUp
    Range("B:B,G:G,I:I,K:K,L:L").Delete Shift:=xlToLeft
    
    Dim lastRow As Long
    Dim lastRowAddress As String

    ActiveSheet.Range("A2").Select

'Check for erroneous data
    Do Until ActiveCell.Address(0, 0) = lastRowAddress
        
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        lastRowAddress = "A" & lastRow + 1
        
        If ActiveCell = "Program:" Or ActiveCell = "Material #" Or ActiveCell = "Run date:" Or ActiveCell = "Run time:" Then
            Rows(ActiveCell.Row).EntireRow.Delete
          
        ElseIf Application.CountA(ActiveCell.EntireRow) = 0 Then
            Rows(ActiveCell.Row).EntireRow.Delete
        
        Else: ActiveCell.Offset(1, 0).Select
        
        End If
                
    Loop

        Range("A" & lastRow).Select
        Rows(ActiveCell.Row).EntireRow.Delete

        lastRow = Cells(Rows.Count, 4).End(xlUp).Row
        lastRowAddress = "A" & lastRow
        
'Column A Fill-in
        Range("A2:A" & lastRow).Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        ActiveCell.FormulaR1C1 = "=R[-1]C"
        Selection.FillDown
        
        Columns("A:A").Copy
        Columns("A:A").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

'Column B Fill-in
        Range("B2:B" & lastRow).Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        ActiveCell.FormulaR1C1 = "=R[-1]C"
        Selection.FillDown
        
        Columns("B:B").Copy
        Columns("B:B").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

'Column K Fill-in
        Range("K2:K" & lastRow).Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        ActiveCell.FormulaR1C1 = "=R[-1]C"
        Selection.FillDown
        
        Columns("K:K").Copy
        Columns("K:K").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        ActiveSheet.Range("A1").Select
        Selection.AutoFilter
        Columns.AutoFit
        
        
'Pivot data

'        Dim pvtCache As PivotCache
'        Dim pvt As PivotTable
'        Dim SrcData As String
'        Dim newSheet As Worksheet
'        Dim StartPvt As String
'        Dim pvtFld As PivotField
'
'
'        SrcData = ActiveSheet.Name & "!" & Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
'
'        Set newSheet = Sheets.Add
'        ActiveSheet.Name = "Pivot"
'
'        StartPvt = newSheet.Name & "!" & newSheet.Range("B2").Address(ReferenceStyle:=xlR1C1)
'
'        Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
'        SourceType:=xlDatabase, _
'        SourceData:=SrcData)
'
'        Set pvt = pvtCache.CreatePivotTable( _
'        TableDestination:=StartPvt, _
'        TableName:="PivotTable1")
'
'        With ActiveSheet.PivotTables("PivotTable1")
'            .ColumnGrand = True
'            .HasAutoFormat = True
'            .DisplayErrorString = False
'            .DisplayNullString = True
'            .EnableDrilldown = True
'            .ErrorString = ""
'            .MergeLabels = False
'            .NullString = ""
'            .PageFieldOrder = 2
'            .PageFieldWrapCount = 0
'            .PreserveFormatting = True
'            .RowGrand = True
'            .SaveData = True
'            .PrintTitles = False
'            .RepeatItemsOnEachPrintedPage = False
'            .TotalsAnnotation = False
'            .CompactRowIndent = 1
'            .InGridDropZones = False
'            .DisplayFieldCaptions = True
'            .DisplayMemberPropertyTooltips = False
'            .DisplayContextTooltips = True
'            .ShowDrillIndicators = True
'            .PrintDrillIndicators = False
'            .AllowMultipleFilters = False
'            .SortUsingCustomLists = True
'            .FieldListSortAscending = False
'            .ShowValuesRow = False
'            .CalculatedMembersInFilters = False
'            .RowAxisLayout xlTabularRow
'        End With
'
'        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material #")
'            .Orientation = xlRowField
'            .Position = 1
'            .RepeatLabels = True
'            .Subtotals(1) = False
'        End With
'
'        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
'            "PivotTable1").PivotFields("Quantity"), "Blocked Quantity", xlSum
'        ActiveSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow
'
'        Columns.AutoFit
'        Range("B2").Select
    
        'MsgBox ("Done!")
    
End Sub

