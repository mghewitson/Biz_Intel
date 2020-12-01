Attribute VB_Name = "Launch_Email"

Sub Launch_Email()
'Created by mhewi3 on 10/25/18
'Updated 1/9/2019 - highlight priority doors red
'Updated 3/18/2019 - Removing Signature space

Application.ScreenUpdating = False

Dim xWs As Worksheet
Dim Rng As Range
Dim xWb As String
Dim launchName As String
Dim launchDate As String

'Check if valid file path
xWb = ActiveWorkbook.Name
savePath = Workbooks(xWb).Worksheets("Setup").Range("B2").Value

    If Dir(savePath, vbDirectory) <> vbNullString Then
    GoTo launchName
    Else
        MsgBox "Folder doesn't exist", vbInformation, "Launch Email Macro"
        Workbooks(xWb).Worksheets("Setup").Activate
        ActiveSheet.Range("B2").Select
        Exit Sub
    End If
    
'inputbox for launch naming convention
launchName:

    launchName = Application.InputBox("What is the name of this launch?", "Launch Name")

        If LenB(Trim$(launchName)) = 0 Then
            MsgBox "Launch name cannot be blank!", vbExclamation
        GoTo launchName
        Else

        If launchName Like "*[*/=']*" Then
            MsgBox "Launch name cannot contain invalid characters (*[*/=']*)!", vbExclamation
        GoTo launchName
        Else
        
        If launchName = "False" Then
            MsgBox "Canceled", , "Launch Email Macro"
        Exit Sub
        
        End If
        End If
        End If

'inputbox for launch date
launchDate:

    launchDate = Application.InputBox("What is the launch date? Format MM-DD-YY", "Launch Date")

'        If LenB(Trim$(launchDate)) = 0 Then
'            MsgBox "Launch Date cannot be blank!", vbExclamation
'        GoTo launchName
'        Else

        If launchDate Like "*[*/=']*" Then
            MsgBox "Launch date cannot contain invalid characters (*[*/=']*)!", vbExclamation
        GoTo launchDate
        Else
        
        If launchDate = "False" Then
            MsgBox "Canceled", , "Launch Email Macro"
        Exit Sub
        
        End If
        End If


xWb = ActiveWorkbook.Name

    ActiveCell.CurrentRegion.Select
    Set Rng = Application.Selection
    Application.Workbooks.Add
    Set xWs = Application.ActiveSheet
    Rng.Copy Destination:=xWs.Range("B2")
    xWs.Range("B2").CurrentRegion.Select
    Application.Selection.Copy
    xWs.Range("B2").PasteSpecial Paste:=xlPasteValues
    xWs.Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Columns.AutoFit

'Highlight FedEx Ground Doors
Dim r As Range, j As Integer
Set r = Range(Range("B2"), Range("B2").End(xlDown))

Dim gdoor1 As String
Dim gdoor2 As String
Dim gdoor3 As String
Dim gdoor4 As String
Dim gdoor5 As String

gdoor1 = Workbooks(xWb).Worksheets("Setup").Range("D2").Value
gdoor2 = Workbooks(xWb).Worksheets("Setup").Range("D3").Value
gdoor3 = Workbooks(xWb).Worksheets("Setup").Range("D4").Value
gdoor4 = Workbooks(xWb).Worksheets("Setup").Range("D5").Value
gdoor5 = Workbooks(xWb).Worksheets("Setup").Range("D6").Value

    ActiveSheet.Range("B2").AutoFilter field:=1, Criteria1:=Array(gdoor1, gdoor2, gdoor3, gdoor4, gdoor5), Operator:=xlFilterValues
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo PDOORS
    Else: Call Highlight
    End If

PDOORS: 'Highlight priority doors red

ActiveSheet.AutoFilter.ShowAllData

Dim pdoor1 As String
Dim pdoor2 As String
Dim pdoor3 As String
Dim pdoor4 As String
Dim pdoor5 As String

pdoor1 = Workbooks(xWb).Worksheets("Setup").Range("E2").Value
pdoor2 = Workbooks(xWb).Worksheets("Setup").Range("E3").Value
pdoor3 = Workbooks(xWb).Worksheets("Setup").Range("E4").Value
pdoor4 = Workbooks(xWb).Worksheets("Setup").Range("E5").Value
pdoor5 = Workbooks(xWb).Worksheets("Setup").Range("E6").Value

    ActiveSheet.Range("B2").AutoFilter field:=1, Criteria1:=Array(pdoor1, pdoor2, pdoor3, pdoor4, pdoor5), Operator:=xlFilterValues
    j = WorksheetFunction.CountA(r.Cells.SpecialCells(xlCellTypeVisible))
    If j = 1 Then
    GoTo EMAIL
    Else: Call Highlight2
    End If


EMAIL: 'Save As path
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("A1").Select
    
'Shows Save As dialog box to allow for unique filenames, saved to dynamic savePath
Dim workbook_Name As Variant
Dim userResponse As Boolean
Dim savePath2 As String
Dim userName As String
Dim category As String

savePath2 = Workbooks(xWb).Worksheets("Setup").Range("B2").Value
userName = Workbooks(xWb).Worksheets("Setup").Range("B1").Value
category = Workbooks(xWb).Worksheets("Setup").Range("B3").Value
    
    ChDir (savePath)
    On Error Resume Next
        userResponse = Application.Dialogs(xlDialogSaveAs).Show(category & " Launch PO's " & "- " & launchName & " - " & launchDate & ".xlsx")
    On Error GoTo 0
        If userResponse = False Then
            MsgBox "Canceled", , "Launch Email Macro"
            Exit Sub
        End If
    
'Create email and attach saved file
Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String
Dim sig As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

'Body of email (HTML format)
    strbody = "Hi Team,<br><br>" & _
    "I just wrote the attached Launch orders. Please clear any <b><i><font size=""3"" color=""red"">RED</font></i></b> PO's first, and code all the <span style=""background-color: #FFFF00"">yellow highlighted</span> PO's FedEx Ground. " & _
    "Let me know if there are any issues.<br><br>" & _
    "Thanks,<br>" '& userName
    
    On Error Resume Next

'Email information
    With OutMail
        .Display
        .To = "NAMarketplaceOps.RetailNSO@Nike.com"
        .CC = ""
        .BCC = ""
        .Subject = category & " Launch PO's " & "- " & launchName & " - " & launchDate
        sig = .HTMLBody
            
            'Finds and replaces extra carriage return with only one return. HTML creates two lines between body and signature, this removes extra line.
            sig = Replace(sig, _
            "<p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p>", _
            "<p class=MsoNormal><o:p>&nbsp;</o:p></p>")
        
        .HTMLBody = strbody & sig
        .Attachments.Add ActiveWorkbook.FullName
        
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Function Highlight()

ActiveSheet.Range("B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Interior.Color = vbYellow

End Function

Function Highlight2()

ActiveSheet.Range("B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select

With Selection
.Font.Color = vbRed
.Font.Bold = True
.Font.Italic = True
End With

End Function


