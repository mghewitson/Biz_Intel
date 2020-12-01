Attribute VB_Name = "Tab_Sorting"
Sub Tab_Sorting()
'
' Tab_Sorting Macro created by mhewi3 11/5/18
'
Application.Calculation = xlManual
Application.ScreenUpdating = False

Dim inputSheet As Worksheet
Dim r As Range, j As Integer

Set r = Range(Range("E2"), Range("E2").End(xlDown))

MRUNFW:
    Set inputSheet = ThisWorkbook.Worksheets("M Run FW")
    
    ActiveWorkbook.Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Running"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MRUNAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
    
MRUNAPP:
    Set inputSheet = ThisWorkbook.Worksheets("M Run App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Running"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
     
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WRUNFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WRUNFW:
    Set inputSheet = ThisWorkbook.Worksheets("W Run FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Running"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WRUNAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WRUNAPP:
    Set inputSheet = ThisWorkbook.Worksheets("W Run App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Running"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WTRAINAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WTRAINAPP:
    Set inputSheet = ThisWorkbook.Worksheets("W Train App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Women's Training", "Womens Training", "Training"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WTRAINFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WTRAINFW:
    Set inputSheet = ThisWorkbook.Worksheets("W Train FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Women's Training", "Womens Training", "Training"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MTRAINAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

MTRAINAPP:
    Set inputSheet = ThisWorkbook.Worksheets("M Train APP")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Men's Training", "Training"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MTRAINFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

MTRAINFW:
    Set inputSheet = ThisWorkbook.Worksheets("M Train FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Men's Training", "Training"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo TRAINEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

TRAINEQ:
    Set inputSheet = ThisWorkbook.Worksheets("Train EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Women's Training", "Men's Training", "Training"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MNSWAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
MNSWAPP:
    Set inputSheet = ThisWorkbook.Worksheets("M NSW APP")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("NSW", "Nike Sportswear"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MNSWFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
MNSWFW:
    Set inputSheet = ThisWorkbook.Worksheets("M NSW FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("NSW", "Nike Sportswear"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WNSWAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WNSWAPP:
    Set inputSheet = ThisWorkbook.Worksheets("W NSW APP")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("NSW", "Nike Sportswear"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WNSWFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
                
WNSWFW:
    Set inputSheet = ThisWorkbook.Worksheets("W NSW FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("NSW", "Nike Sportswear"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo NSWEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
                               
NSWEQ:
    Set inputSheet = ThisWorkbook.Worksheets("NSW EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("NSW", "Nike Sportswear"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo BBALLAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
                                                     
BBALLAPP:
    Set inputSheet = ThisWorkbook.Worksheets("B-ball App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Basketball", "Bball"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo BBALLFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
BBALLFW:
    Set inputSheet = ThisWorkbook.Worksheets("B-ball FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Basketball", "Bball"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo BBALLEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
    
BBALLEQ:
    Set inputSheet = ThisWorkbook.Worksheets("B-ball EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Basketball", "Bball"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo JORDANFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
                                                                
JORDANFW:
    Set inputSheet = ThisWorkbook.Worksheets("Jordan FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Jordan"
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo JORDANAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
JORDANAPP:
    Set inputSheet = ThisWorkbook.Worksheets("Jordan App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Jordan"
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo JORDANEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
JORDANEQ:
    Set inputSheet = ThisWorkbook.Worksheets("Jordan EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Jordan"
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo YAFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
YAFW:
    Set inputSheet = ThisWorkbook.Worksheets("YA FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="<>Jordan", Operator:=xlAnd, Criteria2:="<>Soccer"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Kids", "Kid's", "Kids'", "YA", "Young Athletes"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo YAAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
YAAPP:
    Set inputSheet = ThisWorkbook.Worksheets("YA App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="<>Jordan", Operator:=xlAnd, Criteria2:="<>Soccer"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Kids", "Kid's", "Kids'", "YA", "Young Athletes"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo SBFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
SBFW:
    Set inputSheet = ThisWorkbook.Worksheets("SB FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("SB", "Skate", "Skateboarding"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo SBAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
SBAPP:
    Set inputSheet = ThisWorkbook.Worksheets("SB App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("SB", "Skate", "Skateboarding"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo SBEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
SBEQ:
    Set inputSheet = ThisWorkbook.Worksheets("SB EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("SB", "Skate", "Skateboarding"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo SOCCERAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

SOCCERAPP:
    Set inputSheet = ThisWorkbook.Worksheets("Soccer App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Global Football", "Soccer"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo SOCCERFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
SOCCERFW:
    Set inputSheet = ThisWorkbook.Worksheets("Soccer FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Global Football", "Soccer"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo FOOTBALLAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
     
FOOTBALLAPP:
    Set inputSheet = ThisWorkbook.Worksheets("Football App")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Football", "Performance Football", "American Football"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo FOOTBALLFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
FOOTBALLFW:
    Set inputSheet = ThisWorkbook.Worksheets("Football FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Football", "Performance Football", "American Football"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo FOOTBALLEQ
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

FOOTBALLEQ:
    Set inputSheet = ThisWorkbook.Worksheets("Football EQ")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:=Array("Football", "Performance Football", "American Football"), Operator:=xlFilterValues
    'ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("EQ", "Equipment"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WTENNISAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

WTENNISAPP:
    Set inputSheet = ThisWorkbook.Worksheets("W TENNIS APP")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Tennis"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo WTENNISFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If
        
WTENNISFW:
    Set inputSheet = ThisWorkbook.Worksheets("W TENNIS FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Tennis"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Womens", "Women's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MTENNISAPP
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

MTENNISAPP:
    Set inputSheet = ThisWorkbook.Worksheets("M TENNIS APP")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Tennis"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("AP", "App", "App", "App ", "Apparel"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo MTENNISFW
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

MTENNISFW:
    Set inputSheet = ThisWorkbook.Worksheets("M TENNIS FW")
    
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("E2:K2").AutoFilter field:=1, Criteria1:="Tennis"
    ActiveSheet.Range("E2:K2").AutoFilter field:=2, Criteria1:=Array("Mens", "Men's"), Operator:=xlFilterValues
    ActiveSheet.Range("E2:K2").AutoFilter field:=3, Criteria1:=Array("FW", "Footwear"), Operator:=xlFilterValues
    
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo ENDSUB
    End If
    If j = 2 Then
    Call Paste_Data_2(inputSheet)
    End If
    If j = 3 Then
    Call Paste_Data_3(inputSheet)
    End If
    If j > 3 Then
    Call Paste_Data(inputSheet)
    End If

ENDSUB:
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("A1").Select
    MsgBox "Macro complete", , "Tab Sorting Macro"
    
End Sub

Function Paste_Data(ByRef inputSheet) As Worksheet

    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 4).Select
    ActiveCell.Resize(1, 7).Copy
        inputSheet.Activate
        inputSheet.Range("C12").PasteSpecial Paste:=xlPasteValues
        
    Worksheets("material list").Activate
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(2, 4).Select
    ActiveCell.Resize(1, 7).Select
    Range(Selection, Selection.End(xlDown)).Copy
    
        inputSheet.Activate
        inputSheet.Range("C13").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        inputSheet.Range("C12").Select

End Function

Function Paste_Data_2(ByRef inputSheet) As Worksheet

    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 4).Select
    ActiveCell.Resize(1, 7).Copy
        inputSheet.Activate
        inputSheet.Range("C12").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    inputSheet.Range("C12").Select
    
End Function

Function Paste_Data_3(ByRef inputSheet) As Worksheet

    ActiveSheet.Range("E3:K3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        inputSheet.Activate
        inputSheet.Range("C12").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
        inputSheet.Range("C12").Select
        
End Function
