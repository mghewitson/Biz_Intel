Attribute VB_Name = "Clear_Data_All"
Sub Clear_Data_All()
Attribute Clear_Data_All.VB_Description = "Delete all data from 'For MPO', 'CUP_Blocked_Qty', 'Blkd Data - Final', 'DRS PR's', 'ZMMR_VALIDATE' and 'Size Grid Data'"
Attribute Clear_Data_All.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear_Data_All Macro
' Delete all data from 'For MPO', 'CUP_Blocked_Qty', 'Blkd Data - Final', 'DRS PR's', 'ZMMR_VALIDATE' and 'Size Grid Data'
'

Application.ScreenUpdating = False

Dim answer As Integer
answer = MsgBox("Are you sure you want to delete data from tables?", vbYesNo + vbQuestion, "Delete Table Data")
If answer = vbYes Then
GoTo CLEARMPO
Else
GoTo ENDSUB
End If


CLEARMPO:
'Run clear_mpo_tables macro
Application.Run "Clear_MPO_Tables.Clear_MPO_Tables"

'Unfilter & Delete Tables
ThisWorkbook.Worksheets("SAP PIR's").ListObjects("PIR_DATA").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("CUP_Blocked_Qty").ListObjects("Blkd_Qty_CUP").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("Blkd Data - Final").ListObjects("BLKD_DATA_FINAL").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("DRS PR's").ListObjects("DRS_PRS").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("ZMMR_VALIDATE").ListObjects("ZMMR_VALIDATE").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("Size Grid Data").ListObjects("Size_Grid").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("PR Report").ListObjects("PR_Report").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").ListObjects("Buy_Plan_Align_Flat").AutoFilter.ShowAllData

Dim r As Range, j As Integer

ThisWorkbook.Worksheets("SAP PIR's").Activate

Set r = Range(Range("A1"), Range("A1").End(xlDown))
Set inputSheet = ThisWorkbook.Worksheets("SAP PIR's")

    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo CUP_BLOCKED_QTY
    ElseIf j = 2 Then
    Call Delete_Data_PIR_1(inputSheet)
    Else: Call Delete_Data_PIR_2(inputSheet)
    End If

CUP_BLOCKED_QTY:

    ThisWorkbook.Worksheets("CUP_Blocked_Qty").Activate

    Set r = Range(Range("A1"), Range("A1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("CUP_Blocked_Qty")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo BLKD_DATA_FINAL
        ElseIf j = 2 Then
        Call Delete_Data_Blocked_1(inputSheet)
        Else: Call Delete_Data_Blocked_2(inputSheet)
        End If
        
BLKD_DATA_FINAL:
    
    ThisWorkbook.Worksheets("Blkd Data - Final").Activate

    Set r = Range(Range("A1"), Range("A1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Blkd Data - Final")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo DRS_PRS
        ElseIf j = 2 Then
        Call Delete_Data_Blocked_3(inputSheet)
        Else: Call Delete_Data_Blocked_4(inputSheet)
        End If

DRS_PRS:

    ThisWorkbook.Worksheets("DRS PR's").Activate

    Set r = Range(Range("R1"), Range("R1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("DRS PR's")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo ZMMR
        ElseIf j = 2 Then
        Call Delete_Data_DRSPR1(inputSheet)
        Else: Call Delete_Data_DRSPR2(inputSheet)
        End If

ZMMR:

    ThisWorkbook.Worksheets("ZMMR_VALIDATE").Activate

    Set r = Range(Range("H1"), Range("H1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("ZMMR_VALIDATE")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo SIZES
        ElseIf j = 2 Then
        Call Delete_Data_ZMMR1(inputSheet)
        Else: Call Delete_Data_ZMMR2(inputSheet)
        End If
        
SIZES:

    ThisWorkbook.Worksheets("Size Grid Data").Activate

    Set r = Range(Range("B1"), Range("B1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Size Grid Data")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo PR_REPORT
        ElseIf j = 2 Then
        Call Delete_Data_SIZES1(inputSheet)
        Else: Call Delete_Data_SIZES2(inputSheet)
        End If
        
PR_REPORT:

    ThisWorkbook.Worksheets("PR Report").Activate

    Set r = Range(Range("A1"), Range("A1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("PR Report")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo BUY_ALIGN
        ElseIf j = 2 Then
        Call Delete_PR_Report_1(inputSheet)
        Else: Call Delete_PR_Report_2(inputSheet)
        End If

BUY_ALIGN:

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Activate

    Set r = Range(Range("A4"), Range("A4").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Buy_Plan_Align_Flat")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo LAST
        ElseIf j = 2 Then
        Call Delete_BP_Align_Flat_1(inputSheet)
        Else: Call Delete_BP_Align_Flat_2(inputSheet)
        End If
        
LAST:
    ThisWorkbook.Worksheets("SAP PIR's").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("CUP_Blocked_Qty").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("Blkd Data - Final").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("DRS PR's").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("ZMMR_VALIDATE").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("Size Grid Data").Activate
        ActiveSheet.Range("A2").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("PR Report").Activate
        ActiveSheet.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Activate
        ActiveSheet.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    ThisWorkbook.Worksheets("For MPO").Activate
        ActiveSheet.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    
    MsgBox "Table Data Deleted", vbOKOnly, "Table Delete Macro"
    
ENDSUB:
End Sub

Function Delete_Data_PIR_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("SAP PIR's").Range("A2:Q2").Delete
    
End Function

Function Delete_Data_PIR_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("SAP PIR's").Range("A2:Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_Blocked_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("CUP_Blocked_Qty").Range("A2:K2").Delete
    
End Function

Function Delete_Data_Blocked_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("CUP_Blocked_Qty").Range("A2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_Blocked_3(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Blkd Data - Final").Range("A2:AB2").Delete
    
End Function

Function Delete_Data_Blocked_4(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Blkd Data - Final").Range("A2:AB2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_DRSPR1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("DRS PR's").Range("A2:BD2").Delete
    
End Function

Function Delete_Data_DRSPR2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("DRS PR's").Range("A2:BD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_ZMMR1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("ZMMR_VALIDATE").Range("A2:AK2").Delete
    
End Function

Function Delete_Data_ZMMR2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("ZMMR_VALIDATE").Range("A2:AK2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_SIZES1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Size Grid Data").Range("A2:I2").Delete
    
End Function

Function Delete_Data_SIZES2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Size Grid Data").Range("A2:I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_PR_Report_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("PR Report").Range("A2:CT2").Delete
    
End Function

Function Delete_PR_Report_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("PR Report").Range("A2:CT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_BP_Align_Flat_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Range("A5:AV5").Delete
    
End Function

Function Delete_BP_Align_Flat_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Range("A5:AV5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function
