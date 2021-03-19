Option Compare Database
Option Explicit

Public CurrentTimeTable As String
'Set CurrentTimeTable = "z_tblRunTime" This is set in the startTime()
'z_tblRunTime.FunctionName; Short Text
'z_tblRunTime.StartTime; Number-Double
'z_tblRunTime.EndTime; Number-Double
'z_tblRunTime.RunTime; Number-Double
'z_tblRunTime.RunSec; Number-Double
'z_tblRunTime.RunMinSec; Short Text
'z_tblRunTime.RunHrMinSec; Short Text
'z_tblRunTime.RunDate; Date/Time

Public myStartTime As Double
Public myEndTime As Double
Public myRunTime As Double

Private Function checkTable(callingProcedure As String) As Boolean
'This function checks to see if the calling Function Name is already _
in the Table. If it is it returns True. If it is not found, It will _
add the calling procedure name to the Table and then return True.
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim dbsTime As DAO.Database
Dim rstTime As DAO.Recordset
Dim notFound As Boolean

'Open a pointer to current database and create a record set
Set dbsTime = CurrentDb()
Set rstTime = dbsTime.OpenRecordset(CurrentTimeTable)

rstTime.Index = "PrimaryKey"
rstTime.Seek "=", callingProcedure

notFound = rstTime.NoMatch

'rstTime.Close
'dbsTime.Close

If notFound = True Then
    'Create Record in Table
    Set dbs = CurrentDb()
    Set rst = dbs.OpenRecordset(CurrentTimeTable)
    'Create a new record and the save it back
    rst.AddNew
    rst!FunctionName = callingProcedure
    rst.Update
    'Close out the objects
    rstTime.Close
    dbsTime.Close
    dbs.Close
    'rst.Close
    checkTable = True
ElseIf notFound = False Then
    'Record exsist, do nothing
    rstTime.Close
    dbsTime.Close
    checkTable = True
Else
    'Failed, return FALSE
    rstTime.Close
    dbsTime.Close
    checkTable = False
End If

End Function

Public Function startTime(callingProcedure As String) As Boolean
' This function sets the start time, and writes it to the tblRunTime Table
Dim dbs As DAO.Database
Dim myUpDate_SQL As String
Dim timeTableError As Boolean

myStartTime = Now()
CurrentTimeTable = "z_tblRunTime"
timeTableError = False

'Call function to check if record exsist
timeTableError = checkTable(callingProcedure)

If timeTableError = True Then
    'Open a pointer to current database
    Set dbs = CurrentDb()
    
    ' "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Aggregated_Screen] = " & aggregatedScreensVar & " WHERE (((TC_Screen_Agg.Trial_Card)=" & curTrial_Card & "));"
    myUpDate_SQL = "UPDATE DISTINCTROW " & CurrentTimeTable & " SET " & CurrentTimeTable & ".StartTime = """ & myStartTime & """ WHERE (((" & CurrentTimeTable & ".FunctionName)=""" & callingProcedure & """));"
    ' Debug.Print (myUpDate_SQL)
    dbs.Execute myUpDate_SQL
    
    Debug.Print ("Procedure: " & callingProcedure)
    Debug.Print ("Start Time: " & myStartTime)
    'True will let the time functions contiune to run
    startTime = True
Else
    'Assumed timeTableError = False
    'False will skip all the time functions due to a table error
    startTime = False
End If

End Function

Public Function endTime(callingProcedure As String)
' This Function sets the finish time, writes it to the tblRunTime Table
' and then calls the runTime function
Dim dbs As DAO.Database
Dim myUpDate_SQL As String

myEndTime = Now() ' or could be Time()

'Open a pointer to current database
Set dbs = CurrentDb()

' "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Aggregated_Screen] = " & aggregatedScreensVar & " WHERE (((TC_Screen_Agg.Trial_Card)=" & curTrial_Card & "));"
myUpDate_SQL = "UPDATE DISTINCTROW " & CurrentTimeTable & " SET " & CurrentTimeTable & ".EndTime = """ & myEndTime & """ WHERE (((" & CurrentTimeTable & ".FunctionName)=""" & callingProcedure & """));"
' Debug.Print (myUpDate_SQL)
dbs.Execute myUpDate_SQL

'Debug.Print ("Procedure: " & callingProcedure)
Debug.Print ("End Time: " & myEndTime)

runTime (callingProcedure)

'Debug.Print ("Run Time: " & myRunTime)

End Function

Private Function runTime(callingProcedure As String)
' This function sets the run time, and writes it to the tblRunTime Table
Dim dbs As DAO.Database
Dim myUpDate_SQL As String

myRunTime = myEndTime - myStartTime

Dim myRunSec As Integer
Dim myRunMinSec As String
Dim myRunHrMinSec As String
Dim myRunDate As Date

myRunSec = Int(CSng(myRunTime * 24 * 3600)) ' This computes the run time in Seconds

myRunMinSec = Int(CSng(myRunTime * 24 * 60)) & ":" & Format(myRunTime, "ss") & " Minutes:Seconds" ' This computes the run time in Minutes : Seconds

myRunHrMinSec = Int(CSng(myRunTime * 24)) & ":" & Format(myRunTime, "nn:ss") & " Hours:Minutes:Seconds" ' This computes the run time in Hours : Minutes : Seconds

myRunDate = Date

'Open a pointer to current database
Set dbs = CurrentDb()

' "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Aggregated_Screen] = " & aggregatedScreensVar & " WHERE (((TC_Screen_Agg.Trial_Card)=" & curTrial_Card & "));"
myUpDate_SQL = "UPDATE DISTINCTROW " & CurrentTimeTable & "" _
& " SET " _
& "" & CurrentTimeTable & ".RunTime = """ & myRunTime & """" _
& ", " & CurrentTimeTable & ".RunSec = """ & myRunSec & """" _
& ", " & CurrentTimeTable & ".RunMinSec = """ & myRunMinSec & """" _
& ", " & CurrentTimeTable & ".RunHrMinSec = """ & myRunHrMinSec & """" _
& ", " & CurrentTimeTable & ".RunDate = """ & myRunDate & """" _
& " WHERE (((" & CurrentTimeTable & ".FunctionName)=""" & callingProcedure & """));"
Debug.Print (myUpDate_SQL)
dbs.Execute myUpDate_SQL

End Function
