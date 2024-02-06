# 3D CNC - FAI data read
---
Faith Hall and Thomas Keene
---
## Objective
List of active jobs needing a FAI
 
Details
Display list of active jobs needing a First article inspection reports in Excel. This data will need to be pulled from a SQL Database. As this list gets updated from quality personnel finishing FAI reports push the information back into the SQL database. 
- Filter data to only show active jobs with the word “Required” in the field labeled FAI. 
- Prefer to have a button in Excel that starts the process of grabbing the information out of the database. 
- Have a second button in Excel to write/update the SQL database with the information that has changed in the list. As the Quality personnel updates the list when finishing the reports needed. 

## Proposal
To have either of you or both of you load the SQL engine on your machine. Then I will provide the database for you to load. This way you will be able to test at your own pace without needing to work around my schedule at 3D CNC. We do have some scripts built to read from the database into excel but I haven't the knowledge to write the data back. I will give you a tour of the Excel spreadsheet and database in today's zoom meeting.
- Start with getting SQL engine installed locally. 
- 3D CNC provides a backup copy of the database for you. 
- Use the data for testing your progress.

## Resources 
- [Microsoft SQL Server 2019 Essential Training](https://www.linkedin.com/learning-login/share?forceAccount=false&redirect=https%3A%2F%2Fwww.linkedin.com%2Flearning%2Fmicrosoft-sql-server-2019-essential-training%3Ftrk%3Dshare_ent_url%26shareId%3DEYYpokdQQ1y5WMY6ToeNpA%253D%253D)
- [Intermediate SQL for Data Scientists](https://www.linkedin.com/learning/intermediate-sql-for-data-scientists/the-need-for-sql-in-data-science)
- [Updating Data in Tables Script](https://stackoverflow.com/questions/22229765/update-sql-server-table-from-excel-vba?rq=4)
- [Inserting Data - YouTube Video](https://www.youtube.com/watch?v=lwa56Pdm7Sk)
- [Exporting Data - YouTube Video](https://www.youtube.com/watch?v=AR3fiGr9q44)
- [Exporting Data part 2 - YouTube Video](https://www.youtube.com/watch?v=khdNk0j5Wco)

## Scripts
### SQL Server - Joining Tables
```
SELECT J.[Job] ,j.[Customer], c.[Name], [Top_Lvl_Job], J.[Status], O.[Status], [Open_Operations],
O.[Work_Center], O.[WC_Vendor], O.[Description], O.[Est_Total_Hrs] AS 'Est Operatiom Hours' -- FYI Called Alias
 
,J.[Est_Total_Hrs], J.[Sched_End], J.[Lead_Days], [Part_Number], J.[Description], J.[Priority] 
,D.[Promised_Date], O.[Lead_Days], D.[DeliveryKey], J.[Sched_Start], O.[Sched_Start], O.[Description], J.[Job]
 
,U.[Text1] AS 'Master R'-- FYI Called Alias
,U.[Text2] AS 'Project L'
,U.[Text3] AS 'Repeat'
,U.[Text4] AS 'Packaging R'
,U.[Text5] AS 'FAI_Report'
 
FROM dbo.Job J
INNER JOIN dbo.Customer C ON J.Customer = C.Customer
INNER JOIN dbo.Job_Operation O ON J.job = O.Job
INNER JOIN dbo.User_Values U ON J.User_Values = U.User_Values
INNER JOIN dbo.Delivery D ON J.job = D.job

WHERE J.STATUS = ('active')and O.STATUS != ('C') ORDER BY J.Job --and O.status = ('O')
```

### Excel - Joining Tables
```
= Odbc.Query("dsn=jobboss32", "select #(lf)      J.[Job]#(lf)      ,j.[Customer]#(lf)      ,c.[Name]#(lf)      ,[Top_Lvl_Job]#(lf)      ,J.[Status]#(lf)#(tab)  ,O.[Status][OP STATUS]#(lf)  ,[Open_Operations]#(lf)#(tab)  ,O.[Work_Center]#(lf)#(tab)  ,O.[WC_Vendor]#(lf)#(tab)  ,O.[Description]#(lf)      ,O.[Est_Total_Hrs] [Est Operation Hours]#(lf)      ,J.[Est_Total_Hrs]#(lf)      ,J.[Sched_End]#(lf)#(tab)  ,J.[Lead_Days]#(lf)      ,[Part_Number]#(lf)      ,J.[Description]#(lf)      ,J.[Priority]#(lf)      ,D.[Promised_Date]#(lf)   ,O.[Lead_Days]#(lf), D.[DeliveryKey]#(lf), J.[Sched_Start]#(lf), O.[Sched_Start]#(lf)--#(tab)  ,O.[Description]#(lf),J.[Job]#(lf),U.[Text2]#(lf),J.[Customer_PO],J.[Order_Quantity],J.[Make_Quantity],J.[Completed_Quantity]#(lf)#(lf)from dbo.Job J#(lf)#(lf)inner join dbo.Customer C#(lf)on J.Customer = C.Customer#(lf)inner join dbo.User_Values U#(lf)on J.User_values = U.User_values#(lf)inner join dbo.Job_Operation O#(lf)on J.job = O.Job#(lf)inner join dbo.Delivery D#(lf)on J.Top_Lvl_job = D.job#(lf)where J.status = ('active')and O.status != ('C') Order By J.Job")
```

### SQL Server - Turing View Into Table
```
USE [JB Training]

SELECT TOP (1000) [Job]
      ,[Customer]
      ,[Name]
      ,[Top_Lvl_Job]
      ,[Status]
      ,[Expr1]
      ,[Open_Operations]
      ,[Work_Center]
      ,[WC_Vendor]
      ,[Description]
      ,[Est Operatiom Hours]
      ,[Est_Total_Hrs]
      ,[Sched_End]
      ,[Lead_Days]
      ,[Part_Number]
      ,[Expr3]
      ,[Priority]
      ,[Promised_Date]
      ,[Expr4]
      ,[DeliveryKey]
      ,[Sched_Start]
      ,[Expr5]
      ,[Expr2]
      ,[Expr6]
      ,[Master R]
      ,[Project L]
      ,[Repeat]
      ,[Packaging R]
      ,[FAI_Report]

INTO FAI
FROM [FAI_Report]
```
### VBA - Export and data from SQL Server to Excel 
```
Option Explicit

Private Sub cmdExport_Click()
'    On Error GoTo ErrExit
    
    Dim cn_ADO As ADODB.Connection
    Dim rs_ADO As ADODB.Recordset
    Dim cmd_ADO As ADODB.Command

    
    Dim SQLUser As String
    Dim SQLPassword As String
    Dim SQLServer As String
    Dim DBName As String
    Dim DbConn As String
    
    Dim SQLQuery As String
    
    Dim strStatus As String
    Dim i As Integer
    Dim j As Integer
    Dim jOffset As Integer
    Dim iStartRow As Integer
    Dim iStep As Integer

    iStep = 100
    jOffset = 1
    iStartRow = 1
    i = iStartRow
    
    SQLUser = "super_sa"
    SQLPassword = "super_sa"
    SQLServer = "Faiths-HP\SQLSERVER2022"
    DBName = "JB Training"
    
    DbConn = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & SQLUser & ";Password=" & SQLPassword & ";Initial Catalog=" & DBName & ";" & _
            "Data Source=" & SQLServer & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;" & _
            "Use Encryption for Data=False;Tag with column collation when possible=False"
    
    Set cn_ADO = New ADODB.Connection
    cn_ADO.Open DbConn
    
    SQLQuery = "SELECT "
    SQLQuery = SQLQuery + "[ID], "
    SQLQuery = SQLQuery + "[Job], "
    SQLQuery = SQLQuery + "[Customer], "
    SQLQuery = SQLQuery + "[Name], "
    SQLQuery = SQLQuery + "[FAI_Report] "
    SQLQuery = SQLQuery + "From "
    SQLQuery = SQLQuery + "[JB Training].[dbo].[FAI] "
      
    Application.Cursor = xlWait
    Application.StatusBar = "Logging onto database..."
    
    Set cmd_ADO = New ADODB.Command
    
    cmd_ADO.CommandText = SQLQuery
    cmd_ADO.ActiveConnection = cn_ADO
    cmd_ADO.Execute
    
    ' Open the recordset
    Set rs_ADO = New ADODB.Recordset
    Set rs_ADO.ActiveConnection = cn_ADO
    rs_ADO.Open cmd_ADO
    
    Range(Cells(i, 1), Cells(Rows.Count, jOffset + rs_ADO.Fields.Count)).Clear
    Cells(1, 1).Select

    Application.StatusBar = "Formatting columns..."
    
    'Output Column names
    For j = 0 To rs_ADO.Fields.Count - 1
        Cells(i, j + jOffset).Value = rs_ADO.Fields(CLng(j)).Name
        Cells(i, j + jOffset).Font.Bold = True
        Next j
        
    strStatus = "Loading data..."
    Application.StatusBar = strStatus
    'dataset output
    While Not rs_ADO.EOF
        i = i + 1
        For j = 0 To rs_ADO.Fields.Count - 1
            Cells(i, j + jOffset).Value = rs_ADO.Fields(j).Value
        Next j
        rs_ADO.MoveNext
    Wend
    
    'Close ADO and recordset
    rs_ADO.Close
    Set cn_ADO = Nothing
    Set cmd_ADO = Nothing
    Set rs_ADO = Nothing

    'Application.StatusBar = False
    Application.StatusBar = "Total record count: " & i - iStartRow
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    
'    Exit Sub
'ErrExit:
'            MsgBox "Error: " & Err & " " & Error(Err)
'            Application.StatusBar = False
'            Application.Cursor = xlDefault
'
'            If Not cn_ADO Is Nothing Then
'                Set cn_ADO = Nothing
'            End If
'            If Not cmd_ADO Is Nothing Then
'                Set cmd_ADO = Nothing
'            End If
'            If Not rs_ADO Is Nothing Then
'                Set rs_ADO = Nothing
'            End If
End Sub

Private Sub cmdImport_Click()
'    On Error GoTo ErrExit

    Dim cn_ADO As ADODB.Connection
    Dim cmd_ADO As ADODB.Command

    
    Dim SQLUser As String
    Dim SQLPassword As String
    Dim SQLServer As String
    Dim DBName As String
    Dim DbConn As String
   
    Dim SQLQuery As String
    Dim strWhere As String
    
    'Dim strStatus As String
    Dim i As Integer
    'Dim j As Integer
    Dim jOffset As Integer
    Dim iStartRow As Integer
    'Dim iStep As Integer
   
     'Data Columns
    Dim strID As String
    Dim strJob As String
    Dim strCustomer As String
    Dim strName As String
    Dim strFAI_Report As String

    'iStep = 100
    jOffset = 1
    iStartRow = 2
    i = iStartRow

    SQLUser = "super_sa"
    SQLPassword = "super_sa"
    SQLServer = "Faiths-HP\SQLSERVER2022"
    DBName = "JB Training"
    
    DbConn = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & SQLUser & ";Password=" & SQLPassword & ";Initial Catalog=" & DBName & ";" & _
            "Data Source=" & SQLServer & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;" & _
            "Use Encryption for Data=False;Tag with column collation when possible=False"
    
    Set cn_ADO = New ADODB.Connection
    cn_ADO.Open DbConn

    Set cmd_ADO = New ADODB.Command

    While Cells(i, jOffset).Value <> ""
        strID = Cells(i, 0 + jOffset).Value
        strJob = Cells(i, 1 + jOffset).Value
        strCustomer = Cells(i, 2 + jOffset).Value
        strName = Cells(i, 3 + jOffset).Value
        strFAI_Report = Cells(i, 4 + jOffset).Value
        
        strWhere = "ID = " & strID

        SQLQuery = "UPDATE dbo.FAI " & _
                    "SET " & _
                    "FAI_Report = '" & strFAI_Report & "' " & _
                    "WHERE " & strWhere
              
        cmd_ADO.CommandText = SQLQuery
        cmd_ADO.ActiveConnection = cn_ADO
        cmd_ADO.Execute

        i = i + 1
    Wend

    Set cmd_ADO = Nothing
    Set cn_ADO = Nothing

'    Exit Sub
'
'ErrExit:
'        MsgBox "Error: " & Err & " " & Error(Err)
'        Application.StatusBar = False
'        Application.Cursor = xlDefault
'
'        If Not cn_ADO Is Nothing Then
'            Set cn_ADO = Nothing
'        End If
'            If Not cmd_ADO Is Nothing Then
'            Set cmd_ADO = Nothing


End Sub
```
