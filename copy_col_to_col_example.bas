Attribute VB_Name = "copy_col_to_col"
Sub test()

'create the variables
Set rng_1 = Range("G:G")
Set rng_2 = Range("G2:G6")
ttl_items = WorksheetFunction.CountA(rng_1)
first_items = WorksheetFunction.CountA(rng_2)
ttl_scroll = Round(ttl_items / first_items, 0)

'select starting point
ActiveSheet.Range("F2").Select

'first set of visible items loop
i = 1
For j = 1 To first_items
    i = i + 1
    Cells(i, 2).Value = Cells(j + 1, 7).Text
Next j
    
'scrolling through nested loops
For scroll = 1 To ttl_scroll
    Selection.End(xlDown).Select
    start_point = ActiveCell.Address
    Selection.Offset(0, 1).Select
    new_rng = ActiveCell.Row - 1
    Range(Selection, Selection.Offset(4, 0)).Select
    visible_items = WorksheetFunction.CountA(Selection)
    
    If i = ttl_items Then
        Exit Sub
    Else
    'interior nested loop
        For j = 1 To visible_items
            i = i + 1
            Cells(i, 2).Value = Cells(j + new_rng, 7).Text
        Next j
    End If

    Range(start_point).Select
Next scroll

'return to starting point
ActiveSheet.Range("F2").Select

End Sub
