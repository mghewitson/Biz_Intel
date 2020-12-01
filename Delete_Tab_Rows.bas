Attribute VB_Name = "Delete_Tab_Rows"
Sub Delete_Tab_Rows()
Attribute Delete_Tab_Rows.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_Tab_Rows Macro created by mhewi3 9/25/18
'
Application.ScreenUpdating = False

Dim answer As Integer
answer = MsgBox("Are you sure you want to delete all tab rows?", vbYesNo + vbQuestion, "Delete Tab Rows")
If answer = vbYes Then
GoTo MRUNFW
Else
GoTo ENDSUB2
End If

MRUNFW:
    Worksheets("M Run FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MRUNAPP
    End If
MRUNAPP:
    Worksheets("M Run App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WRUNFW
    End If
WRUNFW:
    Worksheets("W Run FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WRUNAPP
    End If
WRUNAPP:
    Worksheets("W Run App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WTRAINAPP
    End If
WTRAINAPP:
    Worksheets("W Train App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WTRAINFW
    End If
WTRAINFW:
    Worksheets("W Train FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MTRAINAPP
    End If
MTRAINAPP:
    Worksheets("M Train App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MTRAINFW
    End If
MTRAINFW:
    Worksheets("M Train FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo TRAINEQ
    End If
TRAINEQ:
    Worksheets("Train EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MNSWAPP
    End If
MNSWAPP:
    Worksheets("M NSW App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MNSWAFW
    End If
MNSWAFW:
    Worksheets("M NSW FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WNSWAPP
    End If
WNSWAPP:
    Worksheets("W NSW App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WNSWFW
    End If
WNSWFW:
    Worksheets("W NSW FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo NSWEQ
    End If
NSWEQ:
    Worksheets("NSW EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo BBALLAPP
    End If
BBALLAPP:
    Worksheets("B-ball App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo BBALLFW
    End If
BBALLFW:
    Worksheets("B-ball FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo BBALLEQ
    End If
BBALLEQ:
    Worksheets("B-ball EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo JORDANFW
    End If
JORDANFW:
    Worksheets("Jordan FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo JORDANAPP
    End If
JORDANAPP:
    Worksheets("Jordan App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo JORDANEQ
    End If
JORDANEQ:
    Worksheets("Jordan EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo SOCCERAPP
    End If
SOCCERAPP:
    Worksheets("Soccer App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo SOCCERFW
    End If
SOCCERFW:
    Worksheets("Soccer FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo YAFW
    End If
YAFW:
    Worksheets("YA FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo YAAPP
    End If
YAAPP:
    Worksheets("YA App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo SBFW
    End If
SBFW:
    Worksheets("SB FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo SBAPP
    End If
SBAPP:
    Worksheets("SB App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo SBEQ
    End If
SBEQ:
    Worksheets("SB EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo FOOTBALLAPP
    End If
FOOTBALLAPP:
    Worksheets("Football App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo FOOTBALLFW
    End If
FOOTBALLFW:
    Worksheets("Football FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo FOOTBALLEQ
    End If
FOOTBALLEQ:
    Worksheets("Football EQ").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MTENNISAPP
    End If
MTENNISAPP:
    Worksheets("M Tennis App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo MTENNISFW
    End If
MTENNISFW:
    Worksheets("M Tennis FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WTENNISAPP
    End If
WTENNISAPP:
    Worksheets("W Tennis App").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo WTENNISFW
    End If
WTENNISFW:
    Worksheets("W Tennis FW").Activate
    If IsEmpty(ActiveSheet.Range("C12")) = False Then
    Delete
    Else: GoTo ENDSUB
    End If

ENDSUB:
Worksheets("M Run FW").Activate
ActiveSheet.Range("A12").Select
MsgBox "All tab rows deleted", , "Delete Tab Rows Macro"

ENDSUB2:

End Sub

Function Delete()

ActiveSheet.Rows("12:12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A12").Select

End Function
