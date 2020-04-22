Option Compare Database
Option Explicit


Public Sub SetCurrentWorkingTable_TCSA()
'This sub is setting the current working _
table that is used in all the SQL statements.

'CurrentTable = "TC_Screen_Agg"

End Sub


Public Sub SetFirstScreensAndEvents_TCSA()
'This is to load or reload the screening aggrations into the TC_Screen_Agg Table _
1st text <Not Found> is entered in every field _
2nd Final is loaded into [Last_Screen] and into [First_Screen] from XX_Screen_Only

'Set data columns to <Not Found>
Dim myDateVarList As Variant
Dim notFound As String 'The <Not Found> is not used in TSM or elsewhere, becomes a visual that something was missed
notFound = "Not Found"

'Set TC data and first screens as place holder values
Dim emptyID As String 'The dash is not used in TSM or elsewhere, becomes a visual that something was missed
emptyID = "-"
        
Dim emptyEvent As String 'The double E is not used in TSM or elsewhere, becomes a visual that something was missed
emptyEvent = "EE"
        
Dim emptySts_A_T As String 'The dash slash dash is not used in TSM or elsewhere, becomes a visual that something was missed
emptySts_A_T = "-/-"

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [All_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [All_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[First_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[First_Screen]"
        
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [All_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [All_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[Aggregated_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[Aggregated_Screen]"
        
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [All_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [All_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[Last_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[Last_Screen]"
    
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [All_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [All_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".Trial_ID = '" & emptyID & "', " _
        & "" & CurrentTable & ".Final_Sts_A_T = '" & emptySts_A_T & "', " _
        & "" & CurrentTable & ".Event = '" & emptyEvent & "';"
        ' Debug.Print "Completed setting place holder values in columns Trial_ID, and Event."
        
        ' Debug.Print "Completed the" & CurrentTable & "table data set with place holder values Update Query."
    
        'Set Final Screening as Final and as First
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [All_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [All_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".First_Screen = [All_XX_Screen_Only].[First_Screening], " _
        & "" & CurrentTable & ".[Last_Screen] = [All_XX_Screen_Only].[" & columnFinal & "], " _
        & "" & CurrentTable & ".Final_Sts_A_T = [All_XX_Screen_Only].[Final_Sts_A_T], " _
        & "" & CurrentTable & ".Trial_ID = [All_XX_Screen_Only].[Trial_ID], " _
        & "" & CurrentTable & ".Event = [All_XX_Screen_Only].[Event];"
        ' Debug.Print "Completed setting values in columns Trial_ID, Event, First_Screen, Last_Screen."

    ElseIf All_or_Events = "Events" Then
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [Events_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [Events_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[First_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[First_Screen]"
        
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [Events_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [Events_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[Aggregated_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[Aggregated_Screen]"
        
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [Events_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [Events_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".[Last_Screen] = '" & notFound & "';"
        ' Debug.Print "done with Table Column: " & CurrentTable & ".[Last_Screen]"

        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [Events_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [Events_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".Trial_ID = '" & emptyID & "', " _
        & "" & CurrentTable & ".Final_Sts_A_T = '" & emptySts_A_T & "', " _
        & "" & CurrentTable & ".Event = '" & emptyEvent & "';"
        ' Debug.Print "Completed setting place holder values in columns Trial_ID, and Event."
        
        ' Debug.Print "Completed the" & CurrentTable & "table data set with place holder values Update Query."
    
        'Set Final Screening as Final and as First
        CurrentDb.Execute "UPDATE DISTINCTROW " & CurrentTable & " RIGHT JOIN [Events_XX_Screen_Only] ON " & CurrentTable & ".Trial_Card = [Events_XX_Screen_Only].Trial_Card " _
        & "SET " _
        & "" & CurrentTable & ".First_Screen = [Events_XX_Screen_Only].[First_Screening], " _
        & "" & CurrentTable & ".[Last_Screen] = [Events_XX_Screen_Only].[" & columnFinal & "], " _
        & "" & CurrentTable & ".Final_Sts_A_T = [Events_XX_Screen_Only].[Final_Sts_A_T], " _
        & "" & CurrentTable & ".Trial_ID = [Events_XX_Screen_Only].[Trial_ID], " _
        & "" & CurrentTable & ".Event = [Events_XX_Screen_Only].[Event];"
        ' Debug.Print "Completed setting values in columns Trial_ID, Event, First_Screen, Last_Screen."

    Else
        'Un trapped error
    End If

    ' Debug.Print vbCrLf & "The TC_Screen_Agg Table is now ready for ConCat of Screenings into the [Aggregated_Screen] field." & vbCrLf

End Sub


Public Sub Build_and_Set_Aggregated_Screen_TCSA()
'This concats the screen values into a single field to show screen transitions.
Dim myDateVarList As Variant  ' Used to cycle through each date column as an iterable
Dim myIndex As Integer ' This is to get the prior screen
'These are used on the recordset
Dim priorScreenVar As String ' This holds the first screen from the date columns
Dim currentScreenVar As String ' This holds the last column looked at value to see if current column is different
Dim aggregatedScreensVar As String ' This holds the screens as they are collected
Dim curTrial_Card As String ' This is for the update back to the TC_Screen_Agg Table
'Open a recordset to loop through each Record in the recordset
Dim dbs As DAO.Database
Dim rst As DAO.Recordset ' or Recordset2?
Dim mySQLstring As String
'Open a recordset to update the TC_Screen_Agg table
Dim dbs_Agg As DAO.Database
Dim myUpDate_SQL As String

    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        'simple hard typed SQL statement
        mySQLstring = "SELECT All_XX_Screen_Only.* FROM All_XX_Screen_Only INNER JOIN All_TC_Screen_Agg ON All_XX_Screen_Only.Trial_Card = All_TC_Screen_Agg.Trial_Card;"
    ElseIf All_or_Events = "Events" Then
        'simple hard typed SQL statement
        mySQLstring = "SELECT Events_XX_Screen_Only.* FROM Events_XX_Screen_Only INNER JOIN Events_TC_Screen_Agg ON Events_XX_Screen_Only.Trial_Card = Events_TC_Screen_Agg.Trial_Card;"
    Else
        'Un trapped error
    End If
    
    'Open a pointer to current database
    Set dbs = CurrentDb()
    Set dbs_Agg = CurrentDb()
    'Create the recordset with my SQL string
    Set rst = dbs.OpenRecordset(mySQLstring)
    
    'Check for All Reports or only Events
    If All_or_Events = "All" Then
        Do While Not rst.EOF
            myIndex = 0 'This is the first date column in the allColumnsList, reset to 0 on each row
            curTrial_Card = rst![Trial_Card]
            
            For Each myDateVarList In allColumnsList
                'Debug.Print vbCrLf & ("Prior allColumnsList(" & myIndex & "): " & allColumnsList(myIndex))
                'Debug.Print ("Current myDateVarList: " & myDateVarList)
                'Debug.Print ("Current Trial Card: " & curTrial_Card)
                        
                'Assign the prior screen from the first date column
                If myDateVarList = allColumnsList(0) Then
                    'I am on the first date column, set prior screen = [First_Screening]
                    priorScreenVar = rst![First_Screening]
                    currentScreenVar = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    aggregatedScreensVar = rst![First_Screening]
                    'Leave myIndex set to zero here so it is lagging next time arround
                    'Debug.Print ("First date column of record row: " & myDateVarList)
                ElseIf myDateVarList = allColumnsList.Item(allColumnsList.Count - 1) Then
                    'I am at the end of the date columns
                    priorScreenVar = rst("" & allColumnsList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Do nothing to myIndex because this is the last column
                    'Debug.Print ("Last date column of record row: " & myDateVarList)
                Else
                    'I am in the middle of the date columns, set priorScreenVar=x-1 currentScreenVar=x
                    priorScreenVar = rst("" & allColumnsList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Increment the myIndex counter +1
                    myIndex = myIndex + 1
                End If
            
                'Debug.Print ("currentScreenVar =" & currentScreenVar & "=")
                'Debug.Print ("priorScreenVar =" & priorScreenVar & "=")
                    
                If currentScreenVar = "Not Found" _
                    Or currentScreenVar = "POST BT Trial" _
                    Or currentScreenVar = "POST AT Trial" _
                    Or currentScreenVar = "POST FCT Trial" _
                    Or currentScreenVar = "SPLIT" _
                    Or currentScreenVar = "X/X" _
                    Or currentScreenVar = "" _
                    Or currentScreenVar = Empty _
                    Then
                    'Do not add this to the aggragate screening
                    'Pass
                ElseIf currentScreenVar = priorScreenVar Then
                    'The screening has not changed, do not aggragate this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
                    'Pass
                ElseIf Right(currentScreenVar, 2) = Right(aggregatedScreensVar, 2) Then
                    'Debug.Print ("Right(currentScreenVar, 2): " & Right(currentScreenVar, 2))
                    'Debug.Print ("Right(aggregatedScreensVar, 2): " & Right(aggregatedScreensVar, 2))
                    'Pass
                ElseIf Right(currentScreenVar, 2) <> Right(aggregatedScreensVar, 2) Then
                    'Debug.Print ("currentScreenVar: " & currentScreenVar)
                    'Debug.Print ("Right(aggregatedScreensVar, 2): " & Right(aggregatedScreensVar, 2))
                    'Debug.Print ("Right(currentScreenVar, 2): " & Right(currentScreenVar, 2))
                    aggregatedScreensVar = aggregatedScreensVar & "/" & currentScreenVar 'Or use the assigned rst![" & myDateVarList & "]
                End If
                
                If myDateVarList = allColumnsList.Item(allColumnsList.Count - 1) Then
                    'I am at the end of the date columns
                    'Save the aggregated screens to the table
                    '[TC_Screen_Agg].[curTrial_Card] = aggregatedScreensVar
                    'Do nothing to myIndex because this is the last column
                    ' Debug.Print ("Last date column of record row, Writing aggregated values.")
                Else
                    'Do Nothing
                End If
                
            Next myDateVarList
            
            'When last column of row record is reached control is returned to this loop
            'Write the Aggregated screens to the TC_Screen_Agg Table
        
        '    "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Aggregated_Screen] = " & aggregatedScreensVar & " WHERE (((TC_Screen_Agg.Trial_Card)=" & curTrial_Card & "));"
            myUpDate_SQL = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".Aggregated_Screen = """ & aggregatedScreensVar & """ WHERE (((" & CurrentTable & ".Trial_Card)=""" & curTrial_Card & """));"
            ' Debug.Print (myUpDate_SQL)
            dbs_Agg.Execute myUpDate_SQL
            
            myDateVarList = Empty
            ' Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    ElseIf All_or_Events = "Events" Then
        Do While Not rst.EOF
            myIndex = 0 'This is the first date column in the allColumnsList, reset to 0 on each row
            curTrial_Card = rst![Trial_Card]
            
            For Each myDateVarList In trialsOnlyList
                'Debug.Print vbCrLf & ("Prior allColumnsList(" & myIndex & "): " & allColumnsList(myIndex))
                'Debug.Print ("Current myDateVarList: " & myDateVarList)
                'Debug.Print ("Current Trial Card: " & curTrial_Card)
                        
                'Assign the prior screen from the first date column
                If myDateVarList = trialsOnlyList(0) Then
                    'I am on the first date column, set prior screen = [First_Screening]
                    priorScreenVar = rst![First_Screening]
                    currentScreenVar = rst("" & myDateVarList & "") ' rst![ & myDateVarList & ]
                    aggregatedScreensVar = rst![First_Screening]
                    'Leave myIndex set to zero here so it is lagging next time arround
                    'Debug.Print ("First date column of record row: " & myDateVarList)
                ElseIf myDateVarList = trialsOnlyList.Item(trialsOnlyList.Count - 1) Then
                    'I am at the end of the date columns
                    priorScreenVar = rst("" & trialsOnlyList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Do nothing to myIndex because this is the last column
                    'Debug.Print ("Last date column of record row: " & myDateVarList)
                Else
                    'I am in the middle of the date columns, set priorScreenVar=x-1 currentScreenVar=x
                    priorScreenVar = rst("" & trialsOnlyList(myIndex) & "")
                    currentScreenVar = rst("" & myDateVarList & "")
                    'Increment the myIndex counter +1
                    myIndex = myIndex + 1
                End If
            
                'Debug.Print ("currentScreenVar =" & currentScreenVar & "=")
                'Debug.Print ("priorScreenVar =" & priorScreenVar & "=")
                    
                If currentScreenVar = "Not Found" _
                    Or currentScreenVar = "POST BT Trial" _
                    Or currentScreenVar = "POST AT Trial" _
                    Or currentScreenVar = "POST FCT Trial" _
                    Or currentScreenVar = "SPLIT" _
                    Or currentScreenVar = "X/X" _
                    Or currentScreenVar = "" _
                    Or currentScreenVar = Empty _
                    Then
                    'Do not add this to the aggragate screening
                    'Pass
                ElseIf currentScreenVar = priorScreenVar Then
                    'The screening has not changed, do not aggragate this screening
                    'Debug.Print ("currentScreenVar = priorScreenVar : " & currentScreenVar & " EQUALS " & priorScreenVar)
                    'Pass
                ElseIf Right(currentScreenVar, 2) = Right(aggregatedScreensVar, 2) Then
                    'Debug.Print ("Right(currentScreenVar, 2): " & Right(currentScreenVar, 2))
                    'Debug.Print ("Right(aggregatedScreensVar, 2): " & Right(aggregatedScreensVar, 2))
                    'Pass
                ElseIf Right(currentScreenVar, 2) <> Right(aggregatedScreensVar, 2) Then
                    'Debug.Print ("currentScreenVar: " & currentScreenVar)
                    'Debug.Print ("Right(aggregatedScreensVar, 2): " & Right(aggregatedScreensVar, 2))
                    'Debug.Print ("Right(currentScreenVar, 2): " & Right(currentScreenVar, 2))
                    aggregatedScreensVar = aggregatedScreensVar & "/" & currentScreenVar 'Or use the assigned rst![" & myDateVarList & "]
                End If
                
                If myDateVarList = trialsOnlyList.Item(trialsOnlyList.Count - 1) Then
                    'I am at the end of the date columns
                    'Save the aggregated screens to the table
                    '[TC_Screen_Agg].[curTrial_Card] = aggregatedScreensVar
                    'Do nothing to myIndex because this is the last column
                    ' Debug.Print ("Last date column of record row, Writing aggregated values.")
                Else
                    'Do Nothing
                End If
                
            Next myDateVarList
            
            'When last column of row record is reached control is returned to this loop
            'Write the Aggregated screens to the TC_Screen_Agg Table
        
        '    "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".[Aggregated_Screen] = " & aggregatedScreensVar & " WHERE (((TC_Screen_Agg.Trial_Card)=" & curTrial_Card & "));"
            myUpDate_SQL = "UPDATE DISTINCTROW " & CurrentTable & " SET " & CurrentTable & ".Aggregated_Screen = """ & aggregatedScreensVar & """ WHERE (((" & CurrentTable & ".Trial_Card)=""" & curTrial_Card & """));"
            ' Debug.Print (myUpDate_SQL)
            dbs_Agg.Execute myUpDate_SQL
            
            myDateVarList = Empty
            ' Debug.Print ("Next row in Table")
            
            rst.MoveNext
            
        Loop
    
    Else
        'Un trapped error
    End If
    
    rst.Close
    dbs.Close
    
    ' Debug.Print vbCrLf & "The X/X Trial Cards Update Query completed." & vbCrLf

End Sub
