Attribute VB_Name = "Clear_Data_All"
Sub Clear_Data_All()
Attribute Clear_Data_All.VB_Description = "Delete all data from 'For MPO', 'CUP_Blocked_Qty', 'Blkd Data - Final', 'DRS PR's', 'ZMMR_VALIDATE' and 'Size Grid Data'"
Attribute Clear_Data_All.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear_Data_All Macro
' Delete all data from 'For MPO', 'CUP_Blocked_Qty', 'Blkd Data - Final', 'DRS PR's', 'ZMMR_VALIDATE' and 'Size Grid Data'
'

Application.ScreenUpdating = False
Application.Calculation = xlManual

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
ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").ListObjects("Buy_Plan_Align_Flat").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("Coverage Data").ListObjects("Coverage").AutoFilter.ShowAllData
ThisWorkbook.Worksheets("Global Buy").ListObjects("Glbl_Buy").AutoFilter.ShowAllData

Dim r As Range, j As Long

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
        GoTo BUY_ALIGN
        ElseIf j = 2 Then
        Call Delete_Data_SIZES1(inputSheet)
        Else: Call Delete_Data_SIZES2(inputSheet)
        End If

BUY_ALIGN:

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Activate

    Set r = Range(Range("C4"), Range("C4").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Buy_Plan_Align_Flat")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo COVERAGE
        ElseIf j = 2 Then
        Call Delete_BP_Align_Flat_1(inputSheet)
        Else: Call Delete_BP_Align_Flat_2(inputSheet)
        End If
        
COVERAGE:

    ThisWorkbook.Worksheets("Coverage Data").Activate

    Set r = Range(Range("A1"), Range("A1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Coverage Data")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo Glbl_Buy
        ElseIf j = 2 Then
        Call Delete_Coverage_Data_1(inputSheet)
        Else: Call Delete_Coverage_Data_2(inputSheet)
        End If
        
Glbl_Buy:

    ThisWorkbook.Worksheets("Global Buy").Activate

    Set r = Range(Range("A1"), Range("A1").End(xlDown))
    Set inputSheet = ThisWorkbook.Worksheets("Global Buy")

        j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
        If j = 1 Then
        GoTo LAST
        ElseIf j = 2 Then
        Call Delete_Glbl_Data_1(inputSheet)
        Else: Call Delete_Glbl_Data_2(inputSheet)
        End If
        

LAST:
    Dim WS_Count As Integer
    Dim I As Integer

         WS_Count = ActiveWorkbook.Worksheets.Count

         For I = 1 To WS_Count
            ActiveWorkbook.Worksheets(I).Activate
            ActiveSheet.Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
            'MsgBox ActiveWorkbook.Worksheets(I).Name
         Next I

ThisWorkbook.Worksheets("Promo AP CUP").Activate
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

    ThisWorkbook.Worksheets("Blkd Data - Final").Range("A2:AE2").Delete
    
End Function

Function Delete_Data_Blocked_4(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Blkd Data - Final").Range("A2:AE2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_DRSPR1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("DRS PR's").Range("A2:BF2").Delete
    
End Function

Function Delete_Data_DRSPR2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("DRS PR's").Range("A2:BF2").Select
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

Function Delete_BP_Align_Flat_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Range("A5:AZ5").Delete
    
End Function

Function Delete_BP_Align_Flat_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Buy_Plan_Align_Flat").Range("A5:AZ5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Coverage_Data_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Coverage Data").Range("A2:R2").Delete
    
End Function

Function Delete_Coverage_Data_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Coverage Data").Range("A2:R2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Glbl_Data_1(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Global Buy").Range("A2:AB2").Delete
    
End Function

Function Delete_Glbl_Data_2(ByRef inputSheet) As Worksheet

    ThisWorkbook.Worksheets("Global Buy").Range("A2:AB2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function
