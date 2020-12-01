Attribute VB_Name = "Contract_Finder"

Sub Contract_Finder()

Dim ws As Worksheet

Application.ScreenUpdating = False
Application.Calculation = xlManual

'Add a new worksheet
    Set ws = Worksheets.Add
    ActiveSheet.Name = "Output"
    ActiveSheet.Range("A1") = "PO #"
    
'Autofilter Column 69
    Worksheets("DLR Data").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2" & LastRow).AutoFilter field:=69, Criteria1:="X"
    
    Range("BQ1").Select
    Selection.Copy
    Worksheets("Output").Range("A2").PasteSpecial Paste:=xlPasteValues
    
    Range("E1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Worksheets("Output").Range("F1").PasteSpecial Paste:=xlPasteValues
    Worksheets("Output").Range("F:F").RemoveDuplicates Columns:=1
    Worksheets("Output").Activate
    ActiveSheet.Range("F1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        ActiveSheet.Range("B1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Columns(6).ClearContents
    
    Worksheets("DLR Data").Activate
        Range("M1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
    Worksheets("Output").Range("F1").PasteSpecial Paste:=xlPasteValues
    Worksheets("Output").Range("F:F").RemoveDuplicates Columns:=1
    Worksheets("Output").Activate
    ActiveSheet.Range("F1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        ActiveSheet.Range("C1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Columns(6).ClearContents

'Dim startCol As Range
'Dim column: column = 70

Dim rng As Range
Dim startCol As Range
Dim colNum As Integer
Set startCol = Sheets("DLR Data").Range("BR1")
colNum = startCol.Cells(1).column

Set rng = Sheets("DLR Data").Range("BR1") '-- you may change the sheet name according to yours.

'-- here is your loop
i = 1

Do
'Iterate and filter columns
    Sheets("DLR Data").Activate
    rng.Select
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2" & LastRow).AutoFilter field:=colNum, Criteria1:="X"
    
    rng.Select
    Selection.Copy
    Worksheets("Output").Range("B" & Rows.Count).End(xlUp).Offset(1, -1).PasteSpecial Paste:=xlPasteValues
    
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Worksheets("Output").Range("F1").PasteSpecial Paste:=xlPasteValues
    Worksheets("Output").Range("F:F").RemoveDuplicates Columns:=1
    Worksheets("Output").Activate
        ActiveSheet.Range("F1").CurrentRegion.Select
        Selection.Copy
        ActiveSheet.Range("B" & Rows.Count).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Columns(6).ClearContents
    
    Worksheets("DLR Data").Activate
        Range("M2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
    Worksheets("Output").Range("F1").PasteSpecial Paste:=xlPasteValues
    Worksheets("Output").Range("F:F").RemoveDuplicates Columns:=1
    Worksheets("Output").Activate
        ActiveSheet.Range("F1").CurrentRegion.Select
        Selection.Copy
        ActiveSheet.Range("C" & Rows.Count).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Columns(6).ClearContents
    
    Sheets("DLR Data").Activate
    rng.Offset(0, i).Select
        'i = i + 1
    Set rng = ActiveCell
    colNum = rng.Cells(1).column
    
Loop Until ActiveCell.Value = ""


With Worksheets("Output")
    .Columns(1).Font.Bold = True
    .Range("A1:C1").AutoFilter
    .Columns("A:C").AutoFit
End With

    Sheets("DLR Data").Activate
    ActiveSheet.Range("A1").Select
    ActiveSheet.AutoFilter.ShowAllData

    Worksheets("Output").Activate
    ActiveSheet.Range("A1").Select

MsgBox "Macro Complete", , "Contract Output Macro"


End Sub
