Attribute VB_Name = "Module1"
Option Explicit


Sub buy_pull_data()
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AutoRecover.Enabled = False
ActiveWorkbook.EnableAutoRecover = False
Application.Calculation = xlCalculationAutomatic


Dim s As Worksheet
Set s = ActiveWorkbook.Worksheets("Plan Data")

Dim p As Worksheet
Set p = ActiveWorkbook.Worksheets("Plan Pivot")

'Season
Dim k As Long
Dim str1 As String
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

's.Range("L2").Value = str1

'If str1 = "" Then
'ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ListIndex = -1
'End If

'For first part of query
Dim addsql As String
Dim addsql1 As String
addsql = ""
addsql1 = ""
If str1 = "" Then
addsql = ""
addsql1 = ""
Else
addsql = "and upper(season) in (" & str1 & ")"
addsql1 = "and upper(season) in (select merchandiseabbreviation from NA_NFS_Master.dbo.NA_NFS_Master_Season_ID where Season_ID in (" & _
"select Season_ID - 4 from NA_NFS_Master.dbo.NA_NFS_Master_Season_ID where merchandiseabbreviation in (" & str1 & ")))"
End If


'Concept
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 9")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql2 As String
Dim addsq22 As String
addsql2 = ""
addsq22 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql2 = ""
addsq22 = ""
Else
addsql2 = " and upper(concept) in (" & str1 & ")"
addsq22 = " and upper(s.concept) in (" & str1 & ")"
End If

'RPT
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 10")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql3 As String
Dim addsq33 As String
addsql3 = ""
addsq33 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql3 = ""
addsq33 = ""
Else
addsql3 = " and Product_Category in (" & str1 & ")"
addsq33 = " and RPT in (" & str1 & ")"
End If



'Division
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 12")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With


Dim addsq14 As String
Dim addsq44 As String
addsq14 = ""
addsq44 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsq14 = ""
addsq44 = ""
Else
addsq14 = " and Division_Desc in (" & str1 & ")"
addsq44 = " and replace(s.Division,'Division',' ') in (" & str1 & ")"
End If

'Category
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 11")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql5 As String
Dim addsq55 As String
addsql5 = ""
addsq55 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql5 = ""
addsq55 = ""
Else
addsql5 = " and Category_Desc in (" & str1 & ")"
str1 = Replace(str1, "Nike Training", "Athletic Training")
addsq55 = " and s.Category in (" & str1 & ")"
End If


'Gender
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 13")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql6 As String
Dim addsq66 As String
addsql6 = ""
addsq66 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql6 = ""
addsq66 = ""
Else
addsql6 = " and Dept_Desc in (" & str1 & ")"
addsq66 = " and Dept in (" & str1 & ")"
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TrendQry As String
Dim Plan_Qry As String
Dim Qry As String
Dim Qry0 As String
Dim Qry1 As String
Dim Qry2 As String
Dim Qry3 As String
Dim Qry4 As String
Dim Qry5 As String
Dim tdydate As String
Dim usrnm As String

'stc.Columns.EntireColumn.Hidden = False
'stc.Rows.EntireRow.Hidden = False

Worksheets("Plan Data").AutoFilterMode = False

With Worksheets("Plan Data")
.Rows("4:1000000").ClearContents
End With
 
 Qry1 = "select Season, Season_ID, UPPER(Concept) as Concept, Product_Category as RPT, Division_Desc as Division, Category_Desc as Category, Product_ID, Style, Style_Color, Style_Color_Desc, Full_Color_Desc, Color_Code, " & _
 "Attribute_1,  Attribute_2,  Attribute_3,  Attribute_4,  Attribute_5,   Attribute_6,  Attribute_7,  Attribute_8,Attribute_9,  Attribute_10, Attribute_11,  Attribute_12, Attribute_13, Comments, " & _
 "Brand_Initiative_1,  Brand_Initiative_2,  Brand_Initiative_3,  Brand_Initiative_4,  Brand_Initiative_5, Brand_Initiative_6,  Brand_Initiative_7,  Brand_Initiative_8,  Brand_Initiative_9,  Brand_Initiative_10, " & _
 "Wholesale_Cost, Cost_Discount_Pct, Cost_Final, MSRP, IMU_Pct, Presentation_Date, Weeks_at_Reg, Clearance_Date, Weeks_at_Clrnc, Doors, APS_Target as Style_APS, Style_Rnk, Pct_Color, APS_Target_Units as Style_Color_APS, " & _
 "APS_Final_Units, APS_Pct_Inc_On_Clrnc, AUR_Deg_Pct, CASE WHEN Reg_Promo_Sls_Units <= 0 OR Reg_Promo_Sls_Units IS NULL THEN NULL ELSE Reg_Promo_AUR end as Reg_Promo_AUR, Reg_Promo_Sls_Units, Reg_Promo_NetSls_Rtl, " & _
 "Reg_Promo_NetSls_Cost, Reg_Promo_ST_Pct, BOP_Units, BOP_Cost,  Plan_Pct_Off, Clrnc_AUR, Clrnc_NetSls_Units, Clrnc_NetSls_Rtl, Clrnc_NetSls_Cost, " & _
 "Clrnc_ST_Pct, Tot_NetSls_Units, Tot_NetSls_Cost, Tot_NetSls_Retail,Tot_AUC, Total_ST, Receipt_Units_Target as Receipt_units_override, Receipt_Units_Target, Receipt_Cost, Receipt_Units, Invest_Units, EOP_Units, Reg_Promo_Prod_Mgn_Amt, Reg_Promo_Prod_Mgn, " & _
 "Clrnc_Prod_Mgn , Clrnc_Prod_Mgn_Amt, Tot_ProdMgn, Tot_ProdMgn_Amt from NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Assortment_StagingTable_New n " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Division d on n.Division_CD = d.Division_CD left join NA_NFS_Master.dbo.NA_NFS_Master_Retail_Category c on n.Category_ID = c.Category_ID left join [NA_NFS_Master].[dbo].[NA_NFS_Master_Department] g on n.Dept_ID = g.Dept_ID where 1=1 and Add_Drop = 1 " & _
 " " & addsql & addsql2 & addsql3 & addsq14 & addsql5 & addsql6 & " order by Product_Category desc"


Dim connAP As ADODB.Connection
Set connAP = New ADODB.Connection
'connAP.Open "Driver={SQL Server};Server=NA_NFS_Master; Uid = 'GMM'; Pwd= 'Nike1234';"
'cn.Open "Data Source=poedw2; Database=edw_access_views; Persist Security Info=True; User ID=" & UserName & "; Password=" & Teradatapassword & "; Session Mode=ASCII;"
connAP.Open "Data Source=NA_NSO_PLAN; Database=NA_NSO_PLAN; Persist Security Info=True; User ID= 'NFSGMM'; Password='Nike1234'; Session Mode=ASCII;"
'set for 20 min
connAP.CommandTimeout = 1200


Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
rst1.Open Qry1, connAP
s.Range("A4").CopyFromRecordset rst1

'Dim rst2 As ADODB.Recordset
'Set rst2 = New ADODB.Recordset
'rst2.Open Qry1, connAP
'Worksheets("TY Plan").Range("A2").CopyFromRecordset rst2
'
''Dim rst4 As ADODB.Recordset
''Set rst4 = New ADODB.Recordset
''rst4.Open Qry3, connAP
''Worksheets("Top Styles").Range("A2").CopyFromRecordset rst4
'
'Dim rst5 As ADODB.Recordset
'Set rst5 = New ADODB.Recordset
'rst5.Open Qry5, connAP
'Worksheets("Combined Data").Range("A2").CopyFromRecordset rst5


connAP.Close

Set connAP = Nothing

Worksheets("Plan Pivot").PivotTables("PivotTable2").RefreshTable


'Worksheets("Plan Data").Shapes("ListBox1").Activate
's.Range("A4").Activate
Call ActiveWorkbook.Sheets("Plan Data").Activate

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AutoRecover.Enabled = True
ActiveWorkbook.EnableAutoRecover = True

End Sub



Function RunQuery(Query As String, conn As ADODB.Connection, RSrange As Range)

Dim rsAP1 As ADODB.Recordset
Set rsAP1 = New ADODB.Recordset

rsAP1.Open Query, conn, adOpenStatic, adLockReadOnly

RSrange.CopyFromRecordset rsAP1

End Function

Sub Clearall()
Dim i As Long
i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With


i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 9")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 10")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 11")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 12")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 13")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With
End Sub


Sub Clearall2()
Dim i As Long
i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 8")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With
i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 9")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 10")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 11")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 12")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 13")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With
End Sub

Sub Clearall_SR()
Dim i As Long
i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 2")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With
i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 9")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 10")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 11")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 12")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With

i = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 13")
    For i = 1 To .ListCount
        .Selected(i) = False
    Next i
End With
End Sub

Sub Clearalldata()

Dim ws As Worksheet
Set ws = ActiveSheet

'With Worksheets("Plan Data")
With ws
.Rows("4:1000000").ClearContents
End With

ws.Range("A4").Select
End Sub

'Store Level
Sub store_data()
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AutoRecover.Enabled = False
ActiveWorkbook.EnableAutoRecover = False
Application.Calculation = xlCalculationAutomatic


Dim s As Worksheet
Set s = ActiveWorkbook.Worksheets("Store Data")

s.AutoFilterMode = False
ActiveWorkbook.Worksheets("Grid").AutoFilterMode = False

'Season
Dim k As Long
Dim str1 As String
k = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 8")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql As String
Dim addsql1 As String
addsql = ""
addsql1 = ""
If str1 = "" Then
addsql = ""
addsql1 = ""
Else
addsql = "and upper(m.season) in (" & str1 & ")"
addsql1 = "and upper(n.season) in (" & str1 & ")"
End If


'Concept
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 9")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql2 As String
Dim addsq22 As String
addsql2 = ""
addsq22 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql2 = ""
addsq22 = ""
Else
addsql2 = " and upper(m.concept) in (" & str1 & ")"
addsq22 = " and upper(n.concept) in (" & str1 & ")"
End If

'RPT
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 10")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql3 As String
addsql3 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql3 = ""
Else
addsql3 = " and Product_Category in (" & str1 & ")"
End If



'Division
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Store Data").ListBoxes("List Box 12")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With


Dim addsq14 As String
addsq14 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsq14 = ""
Else
addsq14 = " and Division_Desc in (" & str1 & ")"
End If

'Category
str1 = ""
k = 0
With s.ListBoxes("List Box 11")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql5 As String
addsql5 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql5 = ""
Else
addsql5 = " and Category_Desc in (" & str1 & ")"
End If


'Gender
str1 = ""
k = 0
With s.ListBoxes("List Box 13")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql6 As String
addsql6 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql6 = ""
Else
addsql6 = " and Dept_Desc in (" & str1 & ")"
End If


Dim TrendQry As String
Dim Plan_Qry As String
Dim Qry As String
Dim Qry2 As String
Dim Qry3 As String
Dim tdydate As String
Dim usrnm As String

'stc.Columns.EntireColumn.Hidden = False
'stc.Rows.EntireRow.Hidden = False


With s
.Rows("4:1000000").ClearContents
End With

With ActiveWorkbook.Worksheets("Grid")
.Rows("4:1000000").ClearContents
End With


'Qry = ts.Range("A1")

Qry = "select d.Style_Color, case when d.Cluster_ID = '0150' then '0400' else d.Cluster_ID end as Store_No, Store_Desc as Store_Name, " & _
 "convert(int,Presentation_Date) as Phase_In, dateadd(dd,Weeks_at_Reg * 7,convert(int,Presentation_Date)) as Phase_Out, Weeks_at_Reg as Selling_Weeks, APS AS APS_Units, " & _
 "ceiling(Rcpt_Units_Final) as  Receipt_Units, m.Reg_Promo_ST_Pct as Planned_Sell_Through, Minimum as Minimum_Presentation, " & _
 " null as Maximum_Presentation, null as WOC, MerchandiseAbbreviation as SeasonCode, d.Description as StyleCode_Description, BOP_Units as BOP_from_MAT, Comments, Product_Category as RPT, Division_Desc as Division, " & _
 "Attribute_13 as League, Dept_Desc as Gender, Attribute_9 as Team " & _
 "from NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Style_Color_Cluster_Data d " & _
 "left join NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Assortment_StagingTable_New m on d.Style_Color = m.Style_Color " & _
 "and d.Season_ID = m.Season_ID " & _
 "and case when d.Cluster_ID = '0150' then 'DIGITAL' " & _
 "when d.Cluster_ID in ('0269','0019') then 'EMPLOYEE' " & _
 "when d.Cluster_ID not in ('0269','0019','0150') then 'INLINE' end = m.Concept " & _
 "inner join NA_NFS_Master.dbo.NA_NFS_Master_Stores s on d.Cluster_ID = s.Store_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Season_ID i on d.Season_ID = i.Season_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Division di on m.Division_CD = di.Division_CD left join [NA_NFS_Master].[dbo].[NA_NFS_Master_Department] g on m.Dept_ID = g.Dept_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Retail_Category c on m.Category_ID = c.Category_ID where 1=1 and d.Add_Drop = 1 and Min_Pres > 0 and Weeks_at_Reg is not null and Weeks_at_Reg >0 " & _
 " " & addsql & addsql2 & addsql3 & addsq14 & addsql5 & addsql6 & " " & _
 "order by 1,2,3,4,5,6,8,9"
 
 
 Qry2 = "Select n.Style_Color, MAX(Style_Color_Desc) as Description, MAX(Full_Color_Desc) AS Full_Color_Desc, MAX(cast(MSRP as dec(20,2))) as Retail, MAX(round(MSRP/2,0)) as ES_Price, MAX(Brand_Initiative_1) as Initiative, null as Staff," & _
"MAX(case when n.Concept = 'INLINE' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'INLINE' then Weeks_at_Reg end) as Weeks_at_Reg," & _
"/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'INLINE' then Doors end) as Doors, MAX(case when n.Concept = 'INLINE' then APS_Final_Units end) as APS," & _
"MAX(case when n.Concept = 'INLINE' then Comments end) as Comments, NULL as Deep_Buy, null as Contract_Units, MAX(case when n.Concept = 'INLINE' then Division_Desc end) as Divison," & _
"MAX(case when BOP_Units > 1 and n.Concept = 'INLINE' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up " & _
", max(case when s.Cluster_ID = '0028' and Min_Pres = 1 then s.Min_Pres end) as '0028',  max(case when s.Cluster_ID = '0051' and Min_Pres = 1 then s.Min_Pres end) as '0051' " & _
", max(case when s.Cluster_ID = '0060' and Min_Pres = 1 then s.Min_Pres end) as '0060',  max(case when s.Cluster_ID = '0081' and Min_Pres = 1 then s.Min_Pres end) as '0081' " & _
", max(case when s.Cluster_ID = '0082' and Min_Pres = 1 then s.Min_Pres end) as '0082',  max(case when s.Cluster_ID = '0086' and Min_Pres = 1 then s.Min_Pres end) as '0086' " & _
", max(case when s.Cluster_ID = '0093' and Min_Pres = 1 then s.Min_Pres end) as '0093',  max(case when s.Cluster_ID = '0201' and Min_Pres = 1 then s.Min_Pres end) as '0201' " & _
", max(case when s.Cluster_ID = '0240' and Min_Pres = 1 then s.Min_Pres end) as '0240',  max(case when s.Cluster_ID = '0246' and Min_Pres = 1 then s.Min_Pres end) as '0246' " & _
", max(case when s.Cluster_ID = '0303' and Min_Pres = 1 then s.Min_Pres end) as '0303',  max(case when s.Cluster_ID = '0305' and Min_Pres = 1 then s.Min_Pres end) as '0305' " & _
", max(case when s.Cluster_ID = '0307' and Min_Pres = 1 then s.Min_Pres end) as '0307',  max(case when s.Cluster_ID = '0322' and Min_Pres = 1 then s.Min_Pres end) as '0322' " & _
", max(case when s.Cluster_ID = '0323' and Min_Pres = 1 then s.Min_Pres end) as '0323',  max(case when s.Cluster_ID = '0325' and Min_Pres = 1 then s.Min_Pres end) as '0325' " & _
", max(case when s.Cluster_ID = '0350' and Min_Pres = 1 then s.Min_Pres end) as '0350',  max(case when s.Cluster_ID = '0351' and Min_Pres = 1 then s.Min_Pres end) as '0351' " & _
", max(case when s.Cluster_ID = '0352' and Min_Pres = 1 then s.Min_Pres end) as '0352',  max(case when s.Cluster_ID = '0359' and Min_Pres = 1 then s.Min_Pres end) as '0359' " & _
", max(case when s.Cluster_ID = '0360' and Min_Pres = 1 then s.Min_Pres end) as '0360',  max(case when s.Cluster_ID = '0364' and Min_Pres = 1 then s.Min_Pres end) as '0364' " & _
", max(case when s.Cluster_ID = '0365' and Min_Pres = 1 then s.Min_Pres end) as '0365',  max(case when s.Cluster_ID = '0367' and Min_Pres = 1 then s.Min_Pres end) as '0367' " & _
", max(case when s.Cluster_ID = '0368' and Min_Pres = 1 then s.Min_Pres end) as '0368',  max(case when s.Cluster_ID = '0379' and Min_Pres = 1 then s.Min_Pres end) as '0379' " & _
", max(case when s.Cluster_ID = '0381' and Min_Pres = 1 then s.Min_Pres end) as '0381',  max(case when s.Cluster_ID = '0382' and Min_Pres = 1 then s.Min_Pres end) as '0382' " & _
", max(case when s.Cluster_ID = '0268' and Min_Pres = 1 then s.Min_Pres end) as '0268', null, null, null, null, null, null,"
 
 Qry3 = Qry2 & " MAX(case when n.Concept = 'EMPLOYEE' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'EMPLOYEE' then Weeks_at_Reg end) as Weeks_at_Reg," & _
 "/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'EMPLOYEE' then Doors end) as Doors, MAX(case when n.Concept = 'EMPLOYEE' then APS_Final_Units end) as APS, " & _
"MAX(case when n.Concept = 'EMPLOYEE' then Comments end) as Comments, NULL as Deep_Buy, null as Contract_Units, " & _
"MAX(case when BOP_Units > 1 and n.Concept = 'EMPLOYEE' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator_2, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up, " & _
"max(case when s.Cluster_ID = '0019' and Min_Pres = 1 then s.Min_Pres end) as '0019',  max(case when s.Cluster_ID = '0269' and Min_Pres = 1 then s.Min_Pres end) as '0269', NULL, " & _
"MAX(case when n.Concept = 'DIGITAL' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'DIGITAL' then Weeks_at_Reg end) as Weeks_at_Reg, " & _
"/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'DIGITAL' then Doors end) as Doors, MAX(case when n.Concept = 'DIGITAL' then APS_Final_Units end) as APS, " & _
"MAX(case when n.Concept = 'DIGITAL' then Comments end) as Comments, NULL as Active, null as Contract_Units, " & _
"MAX(case when BOP_Units > 1 and n.Concept = 'DIGITAL' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up " & _
"from NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Assortment_StagingTable_New n left join NA_NFS_Master.dbo.NA_NFS_Master_Division d on n.Division_CD = d.Division_CD left join [NA_NFS_Master].[dbo].[NA_NFS_Master_Department] g on n.Dept_ID = g.Dept_ID " & _
"left join NA_NFS_Master.dbo.NA_NFS_Master_Retail_Category c on n.Category_ID = c.Category_ID left join NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Style_Color_Cluster_Data s on n.Style_Color = s.Style_Color " & _
"and n.Season_ID = s.Season_ID and case when s.Cluster_ID = '0150' then 'DIGITAL' when s.Cluster_ID in ('0269','0019') then 'EMPLOYEE' " & _
"when s.Cluster_ID not in ('0269','0019','0150') then 'INLINE' end = n.Concept and n.Add_Drop = 1 and Min_Pres > 0 " & _
"where 1=1 " & addsql1 & addsq22 & addsql3 & addsq14 & addsql5 & addsql6 & "" & _
"group by n.Style_Color "
 
 
'ActiveWorkbook.Worksheets("Grid").Range("F1").Value = Qry3

Dim connAP As ADODB.Connection
Set connAP = New ADODB.Connection
'connAP.Open "Driver={SQL Server};Server=NA_NFS_Master; Uid = 'GMM'; Pwd= 'Nike1234';"
'cn.Open "Data Source=poedw2; Database=edw_access_views; Persist Security Info=True; User ID=" & UserName & "; Password=" & Teradatapassword & "; Session Mode=ASCII;"
connAP.Open "Data Source=NA_NSO_PLAN; Database=NA_NSO_PLAN; Persist Security Info=True; User ID= 'NFSGMM'; Password='Nike1234'; Session Mode=ASCII;"
'set for 20 min
connAP.CommandTimeout = 1200


Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
rst1.Open Qry, connAP
s.Range("A4").CopyFromRecordset rst1

Dim rst2 As ADODB.Recordset
Set rst2 = New ADODB.Recordset
rst2.Open Qry3, connAP
ActiveWorkbook.Worksheets("Grid").Range("B4").CopyFromRecordset rst2

'sc.Range("C6").CopyFromRecordset rst1(2)


'TrendQry = RunQuery(Qry, connAP, s.Range("A1"))


'rst1.Open Qry, conn, adOpenStatic, adLockReadOnly

'RSrange.CopyFromRecordset rst1

'tdydate = Application.WorksheetFunction.Text(Now(), "m/d/yy")
'usrnm = Environ("UserName")
'fl.Range("D2") = Chr(17) & "Updated " & tdydate & " by " & usrnm


connAP.Close

Set connAP = Nothing

'Worksheets("Store Pivot").PivotTables("PivotTable2").RefreshTable

Worksheets("Store Data").Range("A4").Activate
Worksheets("Store Data").Range("A3").AutoFilter

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AutoRecover.Enabled = True
ActiveWorkbook.EnableAutoRecover = True

End Sub

'Seasonal Readiness Tab
Sub seasonal_readiness()
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AutoRecover.Enabled = False
ActiveWorkbook.EnableAutoRecover = False
Application.Calculation = xlCalculationAutomatic


Dim s As Worksheet
Set s = ActiveWorkbook.Worksheets("Seasonal Readiness")

s.AutoFilterMode = False
ActiveWorkbook.Worksheets("Grid").AutoFilterMode = False

'Season
Dim k As Long
Dim str1 As String
k = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 2")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql As String
Dim addsql1 As String
addsql = ""
addsql1 = ""
If str1 = "" Then
addsql = ""
addsql1 = ""
Else
addsql = "and upper(m.season) in (" & str1 & ")"
addsql1 = "and upper(n.season) in (" & str1 & ")"
End If


'Concept
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 9")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql2 As String
Dim addsq22 As String
addsql2 = ""
addsq22 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql2 = ""
addsq22 = ""
Else
addsql2 = " and upper(m.concept) in (" & str1 & ")"
addsq22 = " and upper(n.concept) in (" & str1 & ")"
End If

'RPT
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 10")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql3 As String
addsql3 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql3 = ""
Else
addsql3 = " and Product_Category in (" & str1 & ")"
End If



'Division
str1 = ""
k = 0
With ActiveWorkbook.Worksheets("Seasonal Readiness").ListBoxes("List Box 12")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With


Dim addsq14 As String
addsq14 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsq14 = ""
Else
addsq14 = " and Division_Desc in (" & str1 & ")"
End If

'Category
str1 = ""
k = 0
With s.ListBoxes("List Box 11")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql5 As String
addsql5 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql5 = ""
Else
addsql5 = " and Category_Desc in (" & str1 & ")"
End If


'Gender
str1 = ""
k = 0
With s.ListBoxes("List Box 13")
    For k = 1 To .ListCount
        If .Selected(k) Then
             If str1 = "" Then
          ' MsgBox .List(k)
            str1 = "'" + .List(k) + "'"
            Else
            str1 = str1 + ",'" + .List(k) + "'"
            End If
        End If
    Next k
End With

Dim addsql6 As String
addsql6 = ""
'If ActiveWorkbook.Worksheets("Plan Data").ListBoxes("List Box 8").ControlFormat.ListIndex <> -1 Then
If str1 = "" Then
addsql6 = ""
Else
addsql6 = " and Dept_Desc in (" & str1 & ")"
End If


Dim TrendQry As String
Dim Plan_Qry As String
Dim Qry As String
Dim Qry2 As String
Dim Qry3 As String
Dim tdydate As String
Dim usrnm As String

'stc.Columns.EntireColumn.Hidden = False
'stc.Rows.EntireRow.Hidden = False


With s
.Rows("4:1000000").ClearContents
End With

With ActiveWorkbook.Worksheets("Grid")
.Rows("4:1000000").ClearContents
End With


'Qry = ts.Range("A1")

Qry = "select d.Style_Color, d.Cluster_ID as Store_No, Store_Desc as Store_Name, " & _
"Category_Desc as Category, Product_Category as RPT, Division_Desc as Division, Dept_Desc as Gender, " & _
 "convert(int,Presentation_Date) as Phase_In, dateadd(dd,Weeks_at_Reg * 7,convert(int,Presentation_Date)) as Phase_Out, Weeks_at_Reg as Selling_Weeks, d.APS AS APS_Units, " & _
 "ceiling(d.Rcpt_Units_Final) as  Receipt_Units, m.Reg_Promo_ST_Pct as Planned_Sell_Through, d.Minimum as Minimum_Presentation, " & _
 " MerchandiseAbbreviation as SeasonCode, d.Description as StyleCode_Description, BOP_Units as BOP_from_MAT, case when cd.Style_Color is not null then 'Carryover' else null end as Carryover, Attribute_13 as League " & _
 "from NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Style_Color_Cluster_Data d " & _
 "left join NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Assortment_StagingTable_New m on d.Style_Color = m.Style_Color " & _
 "and d.Season_ID = m.Season_ID " & _
 "and case when d.Cluster_ID = '0150' then 'DIGITAL' " & _
 "when d.Cluster_ID in ('0269','0019') then 'EMPLOYEE' " & _
 "when d.Cluster_ID not in ('0269','0019','0150') then 'INLINE' end = m.Concept " & _
 "left join NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Style_Color_Cluster_Data cd on d.Season_ID-1 = cd.Season_ID and d.Style_Color = cd.Style_Color and d.Cluster_ID = cd.Cluster_ID and cd.Min_Pres = 1 " & _
 "inner join NA_NFS_Master.dbo.NA_NFS_Master_Stores s on d.Cluster_ID = s.Store_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Season_ID i on d.Season_ID = i.Season_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Division di on m.Division_CD = di.Division_CD left join [NA_NFS_Master].[dbo].[NA_NFS_Master_Department] g on m.Dept_ID = g.Dept_ID " & _
 "left join NA_NFS_Master.dbo.NA_NFS_Master_Retail_Category c on m.Category_ID = c.Category_ID where 1=1 and d.Add_Drop = 1 and d.Min_Pres > 0 " & _
 " " & addsql & addsql2 & addsql3 & addsq14 & addsql5 & addsql6 & " " & _
 "order by Store_Desc, Product_Category"
 '1,2,3,4,5,6,8,9"
 
 
 Qry2 = "Select n.Style_Color, MAX(Style_Color_Desc) as Description, MAX(Full_Color_Desc) AS Full_Color_Desc, MAX(cast(MSRP as dec(20,2))) as Retail, MAX(round(MSRP/2,0)) as ES_Price, MAX(Brand_Initiative_1) as Initiative, null as Staff," & _
"MAX(case when n.Concept = 'INLINE' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'INLINE' then Weeks_at_Reg end) as Weeks_at_Reg," & _
"/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'INLINE' then Doors end) as Doors, MAX(case when n.Concept = 'INLINE' then APS_Final_Units end) as APS," & _
"MAX(case when n.Concept = 'INLINE' then Comments end) as Comments, NULL as Deep_Buy, null as Contract_Units, MAX(case when n.Concept = 'INLINE' then Division_Desc end) as Divison," & _
"MAX(case when BOP_Units > 1 and n.Concept = 'INLINE' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up " & _
", max(case when s.Cluster_ID = '0028' and Min_Pres = 1 then s.Min_Pres end) as '0028',  max(case when s.Cluster_ID = '0051' and Min_Pres = 1 then s.Min_Pres end) as '0051' " & _
", max(case when s.Cluster_ID = '0060' and Min_Pres = 1 then s.Min_Pres end) as '0060',  max(case when s.Cluster_ID = '0081' and Min_Pres = 1 then s.Min_Pres end) as '0081' " & _
", max(case when s.Cluster_ID = '0082' and Min_Pres = 1 then s.Min_Pres end) as '0082',  max(case when s.Cluster_ID = '0086' and Min_Pres = 1 then s.Min_Pres end) as '0086' " & _
", max(case when s.Cluster_ID = '0093' and Min_Pres = 1 then s.Min_Pres end) as '0093',  max(case when s.Cluster_ID = '0201' and Min_Pres = 1 then s.Min_Pres end) as '0201' " & _
", max(case when s.Cluster_ID = '0240' and Min_Pres = 1 then s.Min_Pres end) as '0240',  max(case when s.Cluster_ID = '0246' and Min_Pres = 1 then s.Min_Pres end) as '0246' " & _
", max(case when s.Cluster_ID = '0303' and Min_Pres = 1 then s.Min_Pres end) as '0303',  max(case when s.Cluster_ID = '0305' and Min_Pres = 1 then s.Min_Pres end) as '0305' " & _
", max(case when s.Cluster_ID = '0307' and Min_Pres = 1 then s.Min_Pres end) as '0307',  max(case when s.Cluster_ID = '0322' and Min_Pres = 1 then s.Min_Pres end) as '0322' " & _
", max(case when s.Cluster_ID = '0323' and Min_Pres = 1 then s.Min_Pres end) as '0323',  max(case when s.Cluster_ID = '0325' and Min_Pres = 1 then s.Min_Pres end) as '0325' " & _
", max(case when s.Cluster_ID = '0350' and Min_Pres = 1 then s.Min_Pres end) as '0350',  max(case when s.Cluster_ID = '0351' and Min_Pres = 1 then s.Min_Pres end) as '0351' " & _
", max(case when s.Cluster_ID = '0352' and Min_Pres = 1 then s.Min_Pres end) as '0352',  max(case when s.Cluster_ID = '0359' and Min_Pres = 1 then s.Min_Pres end) as '0359' " & _
", max(case when s.Cluster_ID = '0360' and Min_Pres = 1 then s.Min_Pres end) as '0360',  max(case when s.Cluster_ID = '0364' and Min_Pres = 1 then s.Min_Pres end) as '0364' " & _
", max(case when s.Cluster_ID = '0365' and Min_Pres = 1 then s.Min_Pres end) as '0365',  max(case when s.Cluster_ID = '0367' and Min_Pres = 1 then s.Min_Pres end) as '0367' " & _
", max(case when s.Cluster_ID = '0368' and Min_Pres = 1 then s.Min_Pres end) as '0368',  max(case when s.Cluster_ID = '0379' and Min_Pres = 1 then s.Min_Pres end) as '0379' " & _
", max(case when s.Cluster_ID = '0381' and Min_Pres = 1 then s.Min_Pres end) as '0381',  max(case when s.Cluster_ID = '0382' and Min_Pres = 1 then s.Min_Pres end) as '0382' " & _
", max(case when s.Cluster_ID = '0268' and Min_Pres = 1 then s.Min_Pres end) as '0268', null, null, null, null, null, null,"
 
 Qry3 = Qry2 & " MAX(case when n.Concept = 'EMPLOYEE' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'EMPLOYEE' then Weeks_at_Reg end) as Weeks_at_Reg," & _
 "/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'EMPLOYEE' then Doors end) as Doors, MAX(case when n.Concept = 'EMPLOYEE' then APS_Final_Units end) as APS, " & _
"MAX(case when n.Concept = 'EMPLOYEE' then Comments end) as Comments, NULL as Deep_Buy, null as Contract_Units, " & _
"MAX(case when BOP_Units > 1 and n.Concept = 'EMPLOYEE' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator_2, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up, " & _
"max(case when s.Cluster_ID = '0019' and Min_Pres = 1 then s.Min_Pres end) as '0019',  max(case when s.Cluster_ID = '0269' and Min_Pres = 1 then s.Min_Pres end) as '0269', NULL, " & _
"MAX(case when n.Concept = 'DIGITAL' then cast(Presentation_Date as date) end) as Presentation_Date, MAX(case when n.Concept = 'DIGITAL' then Weeks_at_Reg end) as Weeks_at_Reg, " & _
"/*DMCA_OR_MCA*/ null as AA, NULL as Tier, MAX(case when n.Concept = 'DIGITAL' then Doors end) as Doors, MAX(case when n.Concept = 'DIGITAL' then APS_Final_Units end) as APS, " & _
"MAX(case when n.Concept = 'DIGITAL' then Comments end) as Comments, NULL as Active, null as Contract_Units, " & _
"MAX(case when BOP_Units > 1 and n.Concept = 'DIGITAL' THEN 'C/O' else null end) as CO_Indicator, null as Drop_Indicator, null as Never_Outs, null as Omega, null as Style_Guide, null as Gear_Up " & _
"from NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Assortment_StagingTable_New n left join NA_NFS_Master.dbo.NA_NFS_Master_Division d on n.Division_CD = d.Division_CD left join [NA_NFS_Master].[dbo].[NA_NFS_Master_Department] g on n.Dept_ID = g.Dept_ID " & _
"left join NA_NFS_Master.dbo.NA_NFS_Master_Retail_Category c on n.Category_ID = c.Category_ID left join NA_NSO_PLAN.dbo.NA_NFS_Mplan_APT_Style_Color_Cluster_Data s on n.Style_Color = s.Style_Color " & _
"and n.Season_ID = s.Season_ID and case when s.Cluster_ID = '0150' then 'DIGITAL' when s.Cluster_ID in ('0269','0019') then 'EMPLOYEE' " & _
"when s.Cluster_ID not in ('0269','0019','0150') then 'INLINE' end = n.Concept and n.Add_Drop = 1 and Min_Pres > 0 " & _
"where 1=1 " & addsql1 & addsq22 & addsql3 & addsq14 & addsql5 & addsql6 & "" & _
"group by n.Style_Color "
 
 
'ActiveWorkbook.Worksheets("Grid").Range("F1").Value = Qry3

Dim connAP As ADODB.Connection
Set connAP = New ADODB.Connection
'connAP.Open "Driver={SQL Server};Server=NA_NFS_Master; Uid = 'GMM'; Pwd= 'Nike1234';"
'cn.Open "Data Source=poedw2; Database=edw_access_views; Persist Security Info=True; User ID=" & UserName & "; Password=" & Teradatapassword & "; Session Mode=ASCII;"
connAP.Open "Data Source=NA_NSO_PLAN; Database=NA_NSO_PLAN; Persist Security Info=True; User ID= 'NFSGMM'; Password='Nike1234'; Session Mode=ASCII;"
'set for 20 min
connAP.CommandTimeout = 1200


Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
rst1.Open Qry, connAP
s.Range("A4").CopyFromRecordset rst1

Dim rst2 As ADODB.Recordset
Set rst2 = New ADODB.Recordset
rst2.Open Qry3, connAP
ActiveWorkbook.Worksheets("Grid").Range("B4").CopyFromRecordset rst2

'sc.Range("C6").CopyFromRecordset rst1(2)


'TrendQry = RunQuery(Qry, connAP, s.Range("A1"))


'rst1.Open Qry, conn, adOpenStatic, adLockReadOnly

'RSrange.CopyFromRecordset rst1

'tdydate = Application.WorksheetFunction.Text(Now(), "m/d/yy")
'usrnm = Environ("UserName")
'fl.Range("D2") = Chr(17) & "Updated " & tdydate & " by " & usrnm


connAP.Close

Set connAP = Nothing

'Worksheets("Store Pivot").PivotTables("PivotTable2").RefreshTable

Worksheets("Seasonal Readiness").Range("A4").Activate

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AutoRecover.Enabled = True
ActiveWorkbook.EnableAutoRecover = True

End Sub

Sub exporttext()
Dim xlCalc As XlCalculation
Dim fd As String
Dim output As String
Dim FName As String
Dim FNameHelper As String
Dim vendor As String

xlCalc = Application.Calculation
Application.Calculation = xlCalculationManual
On Error GoTo CalcBack

vendor = InputBox("If this Assortment Plan is for a 3rd Party Vendor, then Enter Vendor Code here.  If this Assortment Plan is for Nike products, then simply click OK", "Enter Vendor Code", "NIKE")

'Prompting for a FileName for the output plan
'FName = InputBox("Enter text for the FileName of the JE Assortment Plan File Being Created.", "FileName")
'FName = "\AssortmentPlan_" + FName + ".txt"

Dim savefilename As String
savefilename = Application.GetSaveAsFilename(InitialFileName:="\\NIKE\GroupShare\Buyers\Allocation Analysts\Just Enough\AssortmentPlan_")

FName = savefilename + "txt"
'MsgBox (FName)

output = FName 'ThisWorkbook.path & FName
Open output For Output As #1

Dim x As Object

For Each x In ActiveWorkbook.Worksheets("Store Data").Range("A4:A10000")
    If Len(x.Value) > 0 And Not (IsError(x.Offset(0, 1).Value)) Then
        Print #1, x.Value & "-" & vendor & "-01|" & x.Offset(0, 1).Value & "|" & Format(x.Offset(0, 3).Value, "YYYY-MM-DD") & "|" & Format(x.Offset(0, 4).Value, "YYYY-MM-DD") & "|" & x.Offset(0, 6).Value & "|" & x.Offset(0, 7).Value & "|" & x.Offset(0, 9).Value & "|" & x.Offset(0, 10).Value & "|" & x.Offset(0, 11).Value & "|" & x.Offset(0, 12).Value
    End If
Next x
'
'Dim fso As Object
'Set fso = CreateObject("Scripting.FileSystemObject")
'Dim path As String
'
'With Application.FileDialog(msoFileDialogOpen)
'    .Show
'    If .SelectedItems.Count = 1 Then
'        path = .SelectedItems(1)
'    End If
'End With
'
'If path <> "" Then
'    Open path For Output As path
'End If
'
'Dim oFile As Object
'Set oFile = fso.CreateTextFile(path)
'oFile.WriteLine "test"
'oFile.Close
'Set fso = Nothing
'Set oFile = Nothing



CalcBack:
Close #1
Application.Calculation = xlCalc
End Sub


Sub exporttext2()
Dim xlCalc As XlCalculation
Dim fd As String
Dim output As String
Dim FName As String
Dim FNameHelper As String
Dim vendor As String

xlCalc = Application.Calculation
Application.Calculation = xlCalculationManual
On Error GoTo CalcBack

vendor = InputBox("If this Assortment Plan is for a 3rd Party Vendor, then Enter Vendor Code here.  If this Assortment Plan is for Nike products, then simply click OK", "Enter Vendor Code", "NIKE")

'Prompting for a FileName for the output plan
'FName = InputBox("Enter text for the FileName of the JE Assortment Plan File Being Created.", "FileName")
'FName = "\AssortmentPlan_" + FName + ".txt"

Dim savefilename As String
savefilename = Application.GetSaveAsFilename(InitialFileName:="\\NIKE\GroupShare\Buyers\Allocation Analysts\Just Enough\AssortmentPlan_")

FName = savefilename + "txt"
'MsgBox (FName)

output = FName 'ThisWorkbook.path & FName
Open output For Output As #1

Dim x As Object

For Each x In ActiveWorkbook.Worksheets("Final Store Data").Range("A4:A10000")
    If Len(x.Value) > 0 And Not (IsError(x.Offset(0, 1).Value)) Then
        Print #1, x.Value & "-" & vendor & "-01|" & x.Offset(0, 1).Value & "|" & Format(x.Offset(0, 3).Value, "YYYY-MM-DD") & "|" & Format(x.Offset(0, 4).Value, "YYYY-MM-DD") & "|" & x.Offset(0, 6).Value & "|" & x.Offset(0, 7).Value & "|" & x.Offset(0, 9).Value & "|" & x.Offset(0, 10).Value & "|" & x.Offset(0, 11).Value & "|" & x.Offset(0, 12).Value
    End If
Next x

CalcBack:
Close #1
Application.Calculation = xlCalc
End Sub



Sub pushdata()
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AutoRecover.Enabled = False
ActiveWorkbook.EnableAutoRecover = False
Application.Calculation = xlCalculationAutomatic


Dim lRow1 As Long
Dim lRow2 As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
lRow1 = ActiveWorkbook.Worksheets("Final Store Data").Cells(Rows.Count, 1).End(xlUp).Row
lRow2 = ActiveWorkbook.Worksheets("Store Data").Cells(Rows.Count, 1).End(xlUp).Row
lCol = ActiveWorkbook.Worksheets("Store Data").Cells(4, Columns.Count).End(xlToLeft).Column

Dim r As Range

Set r = ActiveWorkbook.Worksheets("Store Data").Range(Cells(4, 1), Cells(lRow2, lCol + 1))
r.Copy

'ActiveWorkbook.Worksheets("Final Store Data").Range(Cells(lRow1, 1)).PasteSpecial xlPasteAll
ActiveWorkbook.Worksheets("Final Store Data").Range("A" & lRow1 + 1).PasteSpecial xlPasteAll
ActiveWorkbook.Worksheets("Final Store Data").Rows(3).Font.Bold = True

ActiveWorkbook.Worksheets("Store Data").Range("A3").Select

Call ActiveWorkbook.Sheets("Final Store Data").Activate

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AutoRecover.Enabled = True
ActiveWorkbook.EnableAutoRecover = True
End Sub


Sub Clearalldatafinal()
With Worksheets("Final Store Data")
.Rows("4:1000000").ClearContents
End With
End Sub

Sub Validate_All_Cells()

Dim currow As Integer
Dim lastrow As Integer
Dim lenstyle As Integer
Dim curstyle As String

currow = 4

Application.ScreenUpdating = False

Dim x

'loop through each row to find last bit of data.
For Each x In ActiveSheet.Range("A4:A40000")
    If Len(x.Value) > 0 Then lastrow = x.Row
Next x
Do Until currow > lastrow
    If Len(ActiveSheet.Cells(currow, 1).Value) > 0 Then
        'for each row, begin by clearing previous marks
        ActiveSheet.Range(Cells(currow, 1), Cells(currow, 11)).Interior.ColorIndex = xlNone
        ActiveSheet.Range(Cells(currow, 1), Cells(currow, 11)).Font.Italic = False
        lenstyle = Len(ActiveSheet.Cells(currow, 1).Value)
        curstyle = ActiveSheet.Cells(currow, 1).Value
        'Test Style-Color format
        If (lenstyle <> 10 Or Mid(curstyle, 7, 1) <> "-") And (lenstyle <> 14 Or Mid(curstyle, 11, 1) <> "-") Then Call colorcell(currow, 1)
        'Flag Missing Store Numbers
        If IsError(ActiveSheet.Cells(currow, 2).Value) = True Then Call colorcell(currow, 2)
        'Ensure Store Name has its formula
        If Len(ActiveSheet.Cells(currow, 3)) = 0 Then ActiveSheet.Cells(currow, 3).Formula = "=iferror(VLOOKUP(" & ActiveSheet.Cells(currow, 2).Address & ",StoreList!$a$2:$b$244,2,FALSE)," & Chr(34) & Chr(34) & ")"
        'Test Phase In and Phase Out dates
        If ActiveSheet.Cells(currow, 4).Value < Now() - 365 Then Call colorcell(currow, 4)
        If ActiveSheet.Cells(currow, 5).Value < ActiveSheet.Cells(currow, 4).Value Then Call colorcell(currow, 5)
        'Ensure Selling Weeks has its formula
        ActiveSheet.Cells(currow, 6).Formula = "=ROUND((" & ActiveSheet.Cells(currow, 5).Address & "-" & ActiveSheet.Cells(currow, 4).Address & ")/7,0)"
        'Test APS Units
        If ActiveSheet.Cells(currow, 7) <= 0 Then Call colorcell(currow, 7)
        'Test Receipt Units
        If ActiveSheet.Cells(currow, 8) <= 0 Or ActiveSheet.Cells(currow, 8) > 10000 Then Call colorcell(currow, 8)
        'Test Planned Sell Through
        If ActiveSheet.Cells(currow, 9) <= 0 Or ActiveSheet.Cells(currow, 9) > 1 Then Call colorcell(currow, 9)
        'Test Minimum Presentation
        If ActiveSheet.Cells(currow, 10) <= 0 Or ActiveSheet.Cells(currow, 10) > 5000 Then Call colorcell(currow, 10)
    End If
    currow = currow + 1
Loop
    
Application.ScreenUpdating = True

End Sub

Sub colorcell(currow As Integer, curclmn As Integer)

ActiveSheet.Cells(currow, curclmn).Interior.Color = RGB(255, 192, 0)
ActiveSheet.Cells(currow, curclmn).Font.Italic = True

End Sub



