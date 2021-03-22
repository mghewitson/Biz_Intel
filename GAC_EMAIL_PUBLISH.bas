Attribute VB_Name = "GAC_EMAIL_PUBLISH"
Sub GAC_EMAIL_PUBLISH()
Attribute GAC_EMAIL_PUBLISH.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim currentDate As String
    Dim fileDate As String
    Dim savePath As String
    Dim myFileName As String
    
    currentDate = Format(Date, "m/d/yyyy")
    fileDate = Format(Date, "m_d_yy")
    savePath = "C:\Users\mhewi3\Box Sync\Inventory Planning - Promo Apparel\Weekly GAC Change Report\"
    myFileName = savePath & "Weekly GAC Change Report " & fileDate
    
    'Save to folder
    ActiveWorkbook.SaveAs myFileName & ".xlsm"

    'Create email
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    Dim NA_WoW_Change As Chart
    Set NA_WoW_Change = ThisWorkbook.Worksheets("Pivots").ChartObjects("NA_WoW_Change").Chart
    
    Dim Cur_10_Vndrs As Chart
    Set Cur_10_Vndrs = ThisWorkbook.Worksheets("Pivots").ChartObjects("Cur_10_Vndrs").Chart
    
    Dim Cur_Seasonal_Impact As Chart
    Set Cur_Seasonal_Impact = ThisWorkbook.Worksheets("Pivots").ChartObjects("Cur_Seasonal_Impact").Chart


    Dim ttlPromo As String
    Dim newPromo As String
    Dim furtherPromo As String
    Dim promoPercent As String
    Dim ttlNA As String
    Dim season As String
    Dim seasonVal As String
    Dim pt As PivotTable
    Dim myPicture As String
    Dim myPicture1 As String
    Dim myPicture2 As String
    Dim fileName As String
    Dim fileName1 As String
    Dim fileName2 As String
    Dim myPath As String
    Dim strbody As String
    Dim sig As String
    
    Set pt = ThisWorkbook.Worksheets("Pivots").PivotTables("PivotTable3")
    ttlPromo = Format(pt.GetPivotData("[Measures].[Ttl Promo Change Qty]").Value, "0,""K""")
    newPromo = Format(pt.GetPivotData("[Measures].[Promo New Delay Qty]").Value, "0,""K""")
    furtherPromo = Format(pt.GetPivotData("[Measures].[Promo Further Delay Qty]").Value, "0,""K""")
    promoPercent = Format(pt.GetPivotData("[Measures].[NA Promo %]").Value, "0%")
    ttlNA = ThisWorkbook.Worksheets("Hidden Pivots").Range("H1").text
    season = ThisWorkbook.Worksheets("Hidden Pivots").Range("X4").text
    seasonVal = Format(ThisWorkbook.Worksheets("Hidden Pivots").PivotTables("PivotTable8").GetPivotData("[Measures].[Ttl Promo Change Qty]").Value, "0,""K""")
    seasonPercent = Format(ThisWorkbook.Worksheets("Hidden Pivots").PivotTables("PivotTable8").GetPivotData("[Measures].[NA Promo %]").Value, "0%")

    myPicture = "NA_WoW_Change.png"
    myPicture1 = "Cur_10_Vndrs.png"
    myPicture2 = "Cur_Seasonal_Impact.png"
    
    myPath = "C:\Users\mhewi3\Downloads\"

    fileName = myPath & myPicture
    fileName1 = myPath & myPicture1
    fileName2 = myPath & myPicture2
    
    NA_WoW_Change.Export fileName
    Cur_10_Vndrs.Export fileName1
    Cur_Seasonal_Impact.Export fileName2


        strbody = "<BODY style=font-size:11pt;font-family:Calibri><p>Hi Team - </p>" & _
                    "<p>Please see this week's Promo AP GAC Slip Report.</p>" & _
                    "<img src=cid:" & Replace(myPicture, " ", "%20") & " height=372 width=820>" & _
                    "<br><br>" & _
                    "<b><u>This week for NA Promo:</u></b>" & _
                    "<li>We're seeing " & ttlPromo & " in <b>New</b> (" & newPromo & ") or <b>Further</b> (" & furtherPromo & ") delays. " & _
                    "The " & ttlPromo & " represents " & promoPercent & " of the total impact to NA (" & ttlNA & ").</li><br><br>" & _
                    "<img src=cid:" & Replace(myPicture1, " ", "%20") & " height=372 width=820><br><br>" & _
                    "<li>" & season & " is seeing the biggest impact with " & seasonVal & " (" & seasonPercent & " of Ttl NA) in New and Further delays.</li><br><br>" & _
                    "<img src=cid:" & Replace(myPicture2, " ", "%20") & " height=372 width=820><br><br>" & _
                                                        "Please let us know if you have any questions.<br><br>Thanks,</BODY>"
                                                        
    With OutMail
        .Display
        .TO = "NASM.Ops@Nike.com"
        .CC = "Trevor.Rembe@nike.com;Jen.Grissinger@nike.com;Chris.Kondrath@nike.com;Rachel.Uyan@nike.com;Nate.Boyden@nike.com;Ian.Coleman-Berger@nike.com"
        .Subject = "Weekly Promo GAC Slip Report " & currentDate
         sig = .HTMLBody
        .Attachments.Add myFileName & ".xlsm"
        .Attachments.Add fileName
        .Attachments.Add fileName1
        .Attachments.Add fileName2
        
            
            'Finds and replaces extra carriage return with only one return. HTML creates two lines between body and signature, this removes extra line.
            sig = Replace(sig, "<p class=MsoNormal><o:p>&nbsp;</o:p></p>", "")

        .HTMLBody = strbody & sig

    End With

    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing


End Sub

