Attribute VB_Name = "Format_Planner_Report"
Sub Format_Planner_Report()
Attribute Format_Planner_Report.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_Planner_Report Macro
'

'
    Application.ScreenUpdating = False

    Dim WS_Count As Integer
    Dim I As Integer
    Dim currentSheet As Worksheet
    
    WS_Count = ActiveWorkbook.Worksheets.Count

    For I = 1 To WS_Count
        
        Set currentSheet = ActiveWorkbook.Worksheets(I)
        currentSheet.Activate
        ActiveSheet.Range("A1").Select
        Selection.AutoFilter
        ActiveWindow.Zoom = 90
        Cells.Select
        Cells.EntireColumn.AutoFit
        ActiveSheet.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        
        'MsgBox ActiveWorkbook.Worksheets(I).Name
    
    Next I

End Sub
