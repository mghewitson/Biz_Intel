Attribute VB_Name = "GEL_MMX_EDP_FORMAT"
Sub GEL_MMX_EDP_FORMAT()
Attribute GEL_MMX_EDP_FORMAT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GEL_MMX_EDP_FORMAT Macro
' Created by mhewi3 3/25/2022

Application.ScreenUpdating = False
Application.Calculation = xlAutomatic

Dim OG_wb As Workbook
Dim new_wb As Workbook
Dim last_row As Long
Dim data_tab As Worksheet
Dim rng As Range
Dim rnge As Range
Dim pe As String
    
Set OG_wb = ActiveWorkbook
Set new_wb = Workbooks.Add
pe = OG_wb.ActiveSheet.Range("B2").Value

If pe = "APPAREL DIVISION" Then
    
    With new_wb.ActiveSheet
        .Range("A1") = "PPPriorityID"
        .Range("B1") = "Plant"
        .Range("C1") = "DemandSeason"
        .Range("D1") = "CategoryDesc"
        .Range("E1") = "SubCategoryDesc"
        .Range("F1") = "LeagueID"
        .Range("G1") = "LeagueDesc"
        .Range("H1") = "StyleCode"
        .Range("I1") = "ColorCode"
        .Range("J1") = "PriorityDesc"
        .Range("K1") = "Reason"
        .Range("L1") = "RequestedBy"
        .Range("M1") = "Priority"
        .Range("N1") = "DefaultPriority"
        .Range("O1") = "updFlag"
        .Range("P1") = "Error"
        .Range("A1:P1").Interior.Color = vbYellow
        .Range("A1").Select
        Selection.AutoFilter
        Columns.AutoFit
    End With
    
OG_wb.Activate
    Set source_tab = ActiveSheet
    last_row = source_tab.Cells(Rows.Count, "D").End(xlUp).Row
    Set data_tab = Sheets.Add
    source_tab.Activate
    Range("D1:E" & last_row).Select
    Set rng = Application.Selection
    rng.Copy Destination:=data_tab.Range("A1")
    
data_tab.Activate

    With data_tab
        'copy style + season columns, remove duplicates, delete blank rows
        last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        Range("A1:B" & last_row).Select
        Set rng = Application.Selection
        rng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
        data_tab.Range("A1").Select
        Selection.AutoFilter
        Range("A1:B1").AutoFilter Field:=2, Criteria1:="="
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set rnge = .Range("A2:A" & last_row).SpecialCells(xlCellTypeVisible)
        rnge.EntireRow.Delete
        .ShowAllData
    
        'add formula, autofill and paste values
        Range("C2").Select
        last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        ActiveCell.Formula = "=CONCAT(LEFT(A2,2),""20"",RIGHT(A2,2))"
        Selection.AutoFill Destination:=Range("C2:C" & last_row), Type:=xlFillDefault
        Range("C1") = "Full Date"
        Range("C1").Interior.Color = vbYellow
        Range("C2:C" & last_row).Copy
        Range("C2:C" & last_row).PasteSpecial Paste:=xlPasteValues
        Columns.AutoFit
        Application.CutCopyMode = False
    
        'copy style column to correct position for template + copy cleaned data to upload template
        Range("B1:B" & last_row).Select
        Set rng = Application.Selection
        rng.Copy Destination:=ActiveSheet.Range("H1")
        Range("C2:H" & last_row).Select
        Set rng = Application.Selection
        rng.Copy Destination:=new_wb.ActiveSheet.Range("C2")
    End With
    
new_wb.Activate
    With ActiveSheet
        last_row = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
        
        'initial values
        Range("J2") = "P"
        Range("K2") = "GEL"
        Range("L2") = "GOVERNANCE STANDARD"
        Range("M2") = "50"
        Range("N2") = "100"
        Range("O2") = "I"
        Range("J2:O2").Interior.Color = RGB(198, 239, 206)
        Range("J2:O2").Font.Color = RGB(0, 97, 0)
        
        'autfill
        Range("J2:O2").Select
        Selection.AutoFill Destination:=Range("J2:O" & last_row), Type:=xlFillDefault
        Range("A1").Select
        Columns.AutoFit
    End With

Application.CutCopyMode = False
    
Else: With new_wb.ActiveSheet
        .Range("A1") = "PPPriorityID"
        .Range("B1") = "Plant"
        .Range("C1") = "DemandSeason"
        .Range("D1") = "CategoryDesc"
        .Range("E1") = "SubCategoryDesc"
        .Range("F1") = "StyleCode"
        .Range("G1") = "ColorCode"
        .Range("H1") = "Reason"
        .Range("I1") = "RequestedBy"
        .Range("J1") = "PriorityDesc"
        .Range("K1") = "Priority"
        .Range("L1") = "DefaultPriority"
        .Range("M1") = "updFlag"
        .Range("N1") = "Error"
        .Range("A1:N1").Interior.Color = vbYellow
        .Range("A1").Select
        Selection.AutoFilter
        Columns.AutoFit
    End With
    
OG_wb.Activate
    Set source_tab = ActiveSheet
    last_row = source_tab.Cells(Rows.Count, "D").End(xlUp).Row
    Set data_tab = Sheets.Add
    source_tab.Activate
    Range("D1:E" & last_row).Select
    Set rng = Application.Selection
    rng.Copy Destination:=data_tab.Range("A1")
    
    
data_tab.Activate

    With data_tab
        'copy style + season columns, remove duplicates, delete blank rows
        last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        Range("A1:B" & last_row).Select
        Set rng = Application.Selection
        rng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
        data_tab.Range("A1:B1").Select
        Selection.AutoFilter
        Range("A1:B1").AutoFilter Field:=2, Criteria1:="="
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set rnge = .Range("A2:A" & last_row).SpecialCells(xlCellTypeVisible)
        rnge.EntireRow.Delete
        .ShowAllData
    
        'add formula, autofill and paste values
        Range("C2").Select
        last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        ActiveCell.Formula = "=CONCAT(LEFT(A2,2),""20"",RIGHT(A2,2))"
        Selection.AutoFill Destination:=Range("C2:C" & last_row), Type:=xlFillDefault
        Range("C1") = "Full Date"
        Range("C1").Interior.Color = vbYellow
        Range("C2:C" & last_row).Copy
        Range("C2:C" & last_row).PasteSpecial Paste:=xlPasteValues

        Application.CutCopyMode = False
    
        'copy style column to correct position for template + copy cleaned data to upload template
        Range("B1:B" & last_row).Select
        Set rng = Application.Selection
        rng.Copy Destination:=ActiveSheet.Range("F1")
        Range("C2:F" & last_row).Select
        Set rng = Application.Selection
        rng.Copy Destination:=new_wb.ActiveSheet.Range("C2")
    End With
    
new_wb.Activate
    With ActiveSheet
        last_row = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
        
        'initial values
        Range("H2") = "GEL"
        Range("I2") = "GOVERNANCE STANDARD"
        Range("J2") = "P"
        Range("K2") = "50"
        Range("L2") = "100"
        Range("M2") = "I"
        Range("H2:M2").Interior.Color = RGB(198, 239, 206)
        Range("H2:M2").Font.Color = RGB(0, 97, 0)
        
        'autfill
        Range("H2:M2").Select
        Selection.AutoFill Destination:=Range("H2:M" & last_row), Type:=xlFillDefault
        Range("A1").Select
        Columns.AutoFit
    End With

End If
Application.CutCopyMode = False

 Dim savePath As String
 Dim currentDate As String
 Dim myFileName As String
    
    savePath = "C:\Users\mhewi3\Box\Global Supply And Inventory\Global S&IP Operations\03 SUPPLY\OPERATIONS\PARAMETER MANAGEMENT\EDP - PRIORITY\GEL\Uploads\"
    currentDate = Format(Date, "mm_dd_yy")
    myFileName = savePath & pe & "_Upload_" & currentDate
    
    Application.DisplayAlerts = False

    new_wb.SaveAs Filename:=myFileName, FileFormat:=xlWorkbookDefault
End Sub
