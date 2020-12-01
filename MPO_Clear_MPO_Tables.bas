Attribute VB_Name = "Clear_MPO_Tables"
Sub Clear_MPO_Tables()

Application.ScreenUpdating = False

ThisWorkbook.Worksheets("For MPO").Activate

With ThisWorkbook.Worksheets("For MPO")
    .ListObjects("DC_FOR_MPO").AutoFilter.ShowAllData
    .ListObjects("DRS_FOR_MPO").AutoFilter.ShowAllData
    .ListObjects("CAN_FOR_MPO").AutoFilter.ShowAllData
End With

Dim r As Range, j As Integer

Set r = Range(Range("C2"), Range("C2").End(xlDown))
Set inputSheet = ThisWorkbook.Worksheets("For MPO")

    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo DRS
    ElseIf j = 2 Then
    Call Delete_Data_1(inputSheet)
    Else: Call Delete_Data_2(inputSheet)
    End If

    
DRS:
    Set r = Range(Range("M2"), Range("M2").End(xlDown))

    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo CANADA
    ElseIf j = 2 Then
    Call Delete_Data_3(inputSheet)
    Else: Call Delete_Data_4(inputSheet)
    End If
    

CANADA:
    Set r = Range(Range("W2"), Range("W2").End(xlDown))

    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo LAST
    ElseIf j = 2 Then
    Call Delete_Data_5(inputSheet)
    Else: Call Delete_Data_6(inputSheet)
    End If


LAST:
ActiveSheet.Range("A1").Select

End Sub

Function Delete_Data_1(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("B3:K3").Delete

End Function

Function Delete_Data_2(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("B3:K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_3(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("M3:U3").Delete
    
End Function

Function Delete_Data_4(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("M3:U3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function

Function Delete_Data_5(ByRef inputSheet) As Worksheet

ActiveSheet.Range("W3:AD3").Delete

End Function

Function Delete_Data_6(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("W3:AD3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete

End Function
