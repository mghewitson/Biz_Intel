Attribute VB_Name = "ZMMR_PO_FORMAT"
Sub ZMMR_PO_FORMAT()
Attribute ZMMR_PO_FORMAT.VB_Description = "Create common pivot table based on ZMMR_VALIDATE report"
Attribute ZMMR_PO_FORMAT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ZMMR_PO_Format Macro
' Create common pivot table based on ZMMR_VALIDATE report
' mhewi3 6/3/2020

'
    Application.ScreenUpdating = False
    
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim SrcData As String
Dim StartPvt As String
Dim newSheet As Worksheet
Dim dataSheet As Worksheet

    Set dataSheet = ActiveSheet
    
    SrcData = ActiveSheet.Name & "!" & Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
    
    'ActiveSheet.Name = "DATA"
    Set newSheet = Sheets.Add
    'ActiveSheet.Name = "PIVOT"
    

    StartPvt = newSheet.Name & "!" & newSheet.Range("B2").Address(ReferenceStyle:=xlR1C1)

    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)
    
    Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable2")
    
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = True
        .HasAutoFormat = False
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Plant")
        .Orientation = xlRowField
        .Position = 1
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Season code")
        .Orientation = xlRowField
        .Position = 2
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Season Year")
        .Orientation = xlRowField
        .Position = 3
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Vendor")
        .Orientation = xlRowField
        .Position = 4
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("PurchOrder")
        .Orientation = xlRowField
        .Position = 5
        .RepeatLabels = True
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Item")
        .Orientation = xlRowField
        .Position = 6
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Material")
        .Orientation = xlRowField
        .Position = 7
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Material Description")
        .Orientation = xlRowField
        .Position = 8
        .RepeatLabels = False
        .Subtotals(1) = False
    End With

    With ActiveSheet.PivotTables("PivotTable2").PivotFields("GAC Date")
        .Orientation = xlRowField
        .Position = 9
        .RepeatLabels = False
        .Subtotals(1) = False
    End With
        
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Qty Request"), "Sum of Qty Request", xlSum
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Contract").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Princ.agreement item").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Contract Qty").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Purchase Requisition").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item of requisition").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("DRS").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Material").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Material Description").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Plant").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Delivery date").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Qty Request").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Vendor").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("PurchOrder").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("MRP Controller").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("OGAC").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("GAC Date").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("RGAC date").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Acpt Date").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Search term").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Plan Month").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Season code").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Season Year").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Stock Category").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Category").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Description").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Sub Category").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Sub-Cat Desc").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Type Group").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Launch Indicator").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Launch date").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Sales Doc.").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Sales Document Item").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Purch Org").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Purch Grp").Subtotals(1) = False
'
'    ActiveSheet.PivotTables("PivotTable2").PivotFields("Rsn for Ord").Subtotals(1) = False
    
    ActiveSheet.PivotTables("PivotTable2").TableStyle2 = "PivotStyleMedium2"
    
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    With ActiveSheet.Columns("J")
        .ColumnWidth = .ColumnWidth * 2.5
    End With
    
    Range("A1").Select
    newSheet.Name = "PIVOT"
    dataSheet.Name = "DATA"
    
    
End Sub
