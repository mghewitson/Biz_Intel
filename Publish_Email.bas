Attribute VB_Name = "Publish_Email"
Sub Publish()

Dim wB As Workbook
Dim xWb As String
Dim savePath As String
Dim currentDate As String
Dim uniqueComments As String
Dim myFileName As String
Dim lastRow As Long

Set wB = ActiveWorkbook
lastRow = wB.Sheets("Pivots").Cells(Rows.Count, 9).End(xlUp).Row

'Check if valid file path
xWb = ActiveWorkbook.Name
savePath = Workbooks(xWb).Worksheets("Setup").Range("B2").Value

    If Dir(savePath, vbDirectory) <> vbNullString Then
    GoTo SAVE
    Else
        MsgBox "Folder doesn't exist", vbInformation, "Launch Email Macro"
        Workbooks(xWb).Worksheets("Setup").Activate
        ActiveSheet.Range("B2").Select
        Exit Sub
    End If

SAVE:

currentDate = Format(Date, "mm.dd.yyyy")
uniqueComments = Workbooks(xWb).Worksheets("Setup").Range("B3").Value
myFileName = savePath & "Auto-Rapid Replen Summary " & currentDate

'Save to folder
    ActiveWorkbook.SaveCopyAs myFileName & ".xlsm"
    

'Create email and attach saved file
Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String
Dim sig As String
Dim rangeToSend As Range

    Set OutApp = CreateObject("Outlook.Application")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'define a temp path for your image
    tmpImageName = VBA.Environ$("temp") & "\tempo.jpg"

    'Range to save as an image
    Set rangeToSend = Worksheets("Pivots").Range("A3:J" & lastRow)
    ' Now copy that range as a picture
    rangeToSend.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' To save this as an Image we need to do a workaround
    ' First add a temporary sheet and add a Chart there
    ' Resize the chart same as the size of the range
    ' Make the Chart border as Zero
    ' Later once we export that chart as an image
    ' and save it in the above temporary path
    ' will delete this temp sheet

    Set sht = Sheets.Add
    sht.Shapes.AddChart
    sht.Shapes.Item(1).Select
    Set objChart = ActiveChart

    With objChart
        .ChartArea.Height = rangeToSend.Height
        .ChartArea.Width = rangeToSend.Width
        .ChartArea.Fill.Visible = msoFalse
        .ChartArea.Border.LineStyle = xlLineStyleNone
        .Paste
        .Export Filename:=tmpImageName, FilterName:="JPG"
    End With

    'Now delete that temporary sheet
    sht.Delete

    
    'Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

'Body of email (HTML format)
    strbody = "<BODY style=font-size:11pt;font-family:Calibri>Hi Team,<br><br>" & _
    "See below for a high level review of today's Auto/Rapid Replenishment orders. Holler if you have any questions.<br><br>" & _
    "<b><i>" & uniqueComments & "</b></i>" & _
    "<br><br><img src=" & "'" & tmpImageName & "'/><br><br>" & _
    "Let me know if there are any issues.<br><br>" & _
    "Thanks,<br></BODY>" '& userName
    
    On Error Resume Next

'Email information
    With OutMail
        .Display
        .to = "Lst-NA.DSM.NikeDirect.NikeStores.Allocations"
        .CC = ""
        .BCC = ""
        .Subject = "Auto-Rapid Replen Summary " & currentDate
        sig = .HTMLBody
            
            'Finds and replaces extra carriage return with only one return. HTML creates two lines between body and signature, this removes extra line.
            sig = Replace(sig, _
            "<p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p>", _
            "<p class=MsoNormal><o:p>&nbsp;</o:p></p>")
        
        .HTMLBody = strbody & sig
        
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

