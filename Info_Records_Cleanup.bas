Attribute VB_Name = "Info_Records_Cleanup"
Sub Info_Rec_Cleanup()
Attribute Info_Rec_Cleanup.VB_Description = "Info Records by Material Data Cleanup"
Attribute Info_Rec_Cleanup.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Info_Rec_Cleanup Macro
' Info Records by Material Data Cleanup
'

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    Dim OG_Sheet As Worksheet
    Dim Raw_Data As Worksheet
    Dim rng As Range
    
    Set OG_Sheet = activesheet
    
    Sheets.Add
    activesheet.Name = "Raw_Data"
    Set Raw_Data = activesheet
    
    OG_Sheet.Activate
    
    Range("A1").CurrentRegion.Select
    Set rng = Application.Selection
    rng.Copy Destination:=Raw_Data.Range("A1")
    Raw_Data.Activate
    Raw_Data.Range("A1").CurrentRegion.Select
    Application.Selection.Copy
    Raw_Data.Range("A1").PasteSpecial Paste:=xlPasteValues
    Raw_Data.Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveWindow.Zoom = 90
    
    Application.DisplayAlerts = False
    OG_Sheet.Delete
    Application.DisplayAlerts = True
    
    Dim r As Range, j As Integer
    Dim rnge As Range, LstRw As Long
    
    Set r = Range(Range("B1"), Range("B1").End(xlDown))
    
    With Raw_Data
        Range("A1:U1").AutoFilter field:=9, Criteria1:="<>J3AP"
        LstRw = .Cells(.Rows.Count, "B").End(xlUp).Row
        Set rnge = .Range("A2:A" & LstRw).SpecialCells(xlCellTypeVisible)
        rnge.EntireRow.Delete
        .ShowAllData
    End With
        
        Raw_Data.Range("A1:U1").AutoFilter field:=11, Criteria1:="0.00"
            j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
                If j = 1 Then
                    GoTo SKIP1
                End If
    With Raw_Data
        LstRw = .Cells(.Rows.Count, "B").End(xlUp).Row
        Set rnge = .Range("A2:A" & LstRw).SpecialCells(xlCellTypeVisible)
        rnge.EntireRow.Delete
    End With
        
SKIP1:
        Raw_Data.ShowAllData
        
        Range("A1:U1").AutoFilter field:=4, Criteria1:=""
            j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
                If j = 1 Then
                    GoTo SKIP2
                End If
                
    With Raw_Data
        LstRw = .Cells(.Rows.Count, "B").End(xlUp).Row
        Set rnge = .Range("A2:A" & LstRw).SpecialCells(xlCellTypeVisible)
        rnge.EntireRow.Delete
    End With
        
SKIP2:
    Raw_Data.ShowAllData
        
    Raw_Data.Range("S:U").Delete
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Raw_Data.Range("Q2").Select
        ActiveCell.Formula = "=CONCAT(P2,B2)"
        Selection.AutoFill Destination:=Range("Q2:Q" & lastRow), Type:=xlFillDefault

    Raw_Data.Range("R2").Select
        ActiveCell.Formula = "=CONCAT(P2,B2,D2)"
        Selection.AutoFill Destination:=Range("R2:R" & lastRow), Type:=xlFillDefault
    
    With Raw_Data
        Range("Q1") = "Concat"
        Range("R1") = "Concat2"
        Range("Q1:R1").Interior.Color = vbYellow
        Range("Q2:R" & lastRow).Copy
        Range("Q2:R" & lastRow).PasteSpecial Paste:=xlPasteValues
        Range("A:A").Delete
        Range("A1").Select
        Columns.AutoFit
    End With

    
    
End Sub
