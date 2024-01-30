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

## Scripts
### SQL Server
```
SELECT J.[Job] ,j.[Customer], c.[Name], [Top_Lvl_Job], J.[Status], O.[Status], [Open_Operations],
O.[Work_Center], O.[WC_Vendor], O.[Description], O.[Est_Total_Hrs] AS 'Est Operatiom Hours' -- FYI Called Alias
 
,J.[Est_Total_Hrs], J.[Sched_End], J.[Lead_Days], [Part_Number], J.[Description], J.[Priority] 
,D.[Promised_Date], O.[Lead_Days], D.[DeliveryKey], J.[Sched_Start], O.[Sched_Start], O.[Description], J.[Job]
 
,U.[Text1] AS 'Master R'-- FYI Called Alias
,U.[Text2] AS 'Project L'
,U.[Text3] AS 'Repeat'
,U.[Text4] AS 'Packaging R'
 
FROM dbo.Job J
INNER JOIN dbo.Customer C ON J.Customer = C.Customer
INNER JOIN dbo.Job_Operation O ON J.job = O.Job
INNER JOIN dbo.User_Values U ON J.User_Values = U.User_Values
INNER JOIN dbo.Delivery D ON J.job = D.job
WHERE J.STATUS = ('active')and O.STATUS != ('C') ORDER BY J.Job --and O.status = ('O')
```

### Excel
```
= Odbc.Query("dsn=jobboss32", "select #(lf)      J.[Job]#(lf)      ,j.[Customer]#(lf)      ,c.[Name]#(lf)      ,[Top_Lvl_Job]#(lf)      ,J.[Status]#(lf)#(tab)  ,O.[Status][OP STATUS]#(lf)  ,[Open_Operations]#(lf)#(tab)  ,O.[Work_Center]#(lf)#(tab)  ,O.[WC_Vendor]#(lf)#(tab)  ,O.[Description]#(lf)      ,O.[Est_Total_Hrs] [Est Operation Hours]#(lf)      ,J.[Est_Total_Hrs]#(lf)      ,J.[Sched_End]#(lf)#(tab)  ,J.[Lead_Days]#(lf)      ,[Part_Number]#(lf)      ,J.[Description]#(lf)      ,J.[Priority]#(lf)      ,D.[Promised_Date]#(lf)   ,O.[Lead_Days]#(lf), D.[DeliveryKey]#(lf), J.[Sched_Start]#(lf), O.[Sched_Start]#(lf)--#(tab)  ,O.[Description]#(lf),J.[Job]#(lf),U.[Text2]#(lf),J.[Customer_PO],J.[Order_Quantity],J.[Make_Quantity],J.[Completed_Quantity]#(lf)#(lf)from dbo.Job J#(lf)#(lf)inner join dbo.Customer C#(lf)on J.Customer = C.Customer#(lf)inner join dbo.User_Values U#(lf)on J.User_values = U.User_values#(lf)inner join dbo.Job_Operation O#(lf)on J.job = O.Job#(lf)inner join dbo.Delivery D#(lf)on J.Top_Lvl_job = D.job#(lf)where J.status = ('active')and O.status != ('C') Order By J.Job")
```
