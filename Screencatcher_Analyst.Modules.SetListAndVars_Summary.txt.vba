Option Compare Database
Option Explicit

Public myBTvar As String
Public aCL_BT_index_pos As Long ' Index of the BT Event in the allColumnsList
Public myATvar As String
Public aCL_AT_index_pos As Long ' Index of the AT Event in the allColumnsList
Public myFCTvar As String
Public aCL_FCT_index_pos As Long ' Index of the FCT Event in the allColumnsList
Public myOWLDvar As String
Public aCL_OWLD_index_pos As Long ' Index of the OWLD Event in the allColumnsList
Public myFinalvar As String
Public aCL_Final_index_pos As Long ' Index of the Final Event in the allColumnsList
Public aCL_EventIndexList As Object ' List of allColumnsList Trial dates index
Public EventFilterList As Object ' List of Event query paramiters to filter the returned recordset


Sub AddingToMySummaryDateLists()
'If this breaks ,IndexOf or IndexInArray a reference has broken
'https://excelmacromastery.com/vba-arraylist/
'1.Select Tools and then References from the menu.
'2.Click on the Browse.
'3.Find the file mscorlib.tlb and click Open. It should be in a folder like this C:\Windows\Microsoft.NET\Framework\v4.0.30319.
'4.Scroll down the list and check mscorlib.dll.
'5.Click Ok.

'BT Event
myBTvar = trialsOnlyList(0) 'trialsOnlyList(0).Add "2017/06/30" ' BT
'aCL_BT_index_pos = IndexInArray(allColumnsList, myBTvar)
aCL_BT_index_pos = allColumnsList.IndexOf(myBTvar, 0)

'AT Event
myATvar = trialsOnlyList(1) 'trialsOnlyList(1).Add "2017/08/18" ' AT
'aCL_AT_index_pos = IndexInArray(allColumnsList, myATvar)
aCL_AT_index_pos = allColumnsList.IndexOf(myATvar, 0)

'FCT Event
myFCTvar = trialsOnlyList(2) 'trialsOnlyList(2).Add "2018/10/26" ' FCT
'aCL_FCT_index_pos = IndexInArray(allColumnsList, myFCTvar)
aCL_FCT_index_pos = allColumnsList.IndexOf(myFCTvar, 0)

'OWLD Event
myOWLDvar = trialsOnlyList(3) 'trialsOnlyList(3).Add "2019/09/19" ' OWLD
'aCL_OWLD_index_pos = IndexInArray(allColumnsList, myOWLDvar)
aCL_OWLD_index_pos = allColumnsList.IndexOf(myOWLDvar, 0)

'Final
myFinalvar = trialsOnlyList(4) 'trialsOnlyList(4).Add "2020/04/03" ' Final
'aCL_Final_index_pos = IndexInArray(allColumnsList, myFinalvar)
aCL_Final_index_pos = allColumnsList.IndexOf(myFinalvar, 0)

'List of allColumnsList Trial dates index
'Dim aCL_EventIndexList As Object ' This Declares the object handle, without any properties or methods
'Dim aCL_EventIndexList As New ArrayList ' This Declares the object handle with early binding
Set aCL_EventIndexList = CreateObject("System.Collections.ArrayList") ' This is late binding
'aCL_EventIndexList
' Add items
aCL_EventIndexList.Add aCL_BT_index_pos ' BT aCL_EventIndex(0)
aCL_EventIndexList.Add aCL_AT_index_pos ' AT aCL_EventIndex(1)
aCL_EventIndexList.Add aCL_FCT_index_pos ' FCT aCL_EventIndex(2)
'aCL_EventIndexList.Add aCL_OWLD_index_pos ' OWLD aCL_EventIndex(3)
'aCL_EventIndexList.Add aCL_Final_index_pos ' Final aCL_EventIndex(4)

'List of Event query paramiters to filter the returned recordset
'Dim EventFilterList As Object ' This Declares the object handle, without any properties or methods
'Dim EventFilterList As New ArrayList ' This Declares the object handle with early binding
Set EventFilterList = CreateObject("System.Collections.ArrayList") ' This is late binding
'EventFilterList
' Add items
EventFilterList.Add "((Event = 'BT') AND (Trial_ID LIKE 'B*'))" 'This is BT New Cards written at the BT Event. Trial_ID will have a leading "B" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'BT') AND (Trial_ID LIKE '*B*'))" 'This is BT Roll and Splits
EventFilterList.Add "((Event = 'AT') AND (Trial_ID LIKE 'C*'))" 'This is AT New Cards written at the AT Event. Trial_ID will have a leading "C" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'AT') AND (Trial_ID LIKE '*C*'))" 'This is AT Roll and Splits
EventFilterList.Add "((Event = 'FCT') AND (Trial_ID LIKE 'F*'))" ''This is FCT New Cards written at the FCT Event. Trial_ID will have a leading "F" and may or may not have trailing Trial_ID's.
EventFilterList.Add "((Event <> 'FCT') AND (Trial_ID LIKE '*F*'))" 'This is FCT Roll and Splits
EventFilterList.Add "((Event = 'BT') OR (Event = 'AT') OR (Event = 'FCT'))" 'This is All BT/AT/FCT New and Splits
EventFilterList.Add "(((Event = 'AT') OR (Event = 'FCT')) AND ((Trial_ID LIKE '*C*') OR (Trial_ID LIKE '*F*')))" 'This is INSURV New, Roll and Splits

End Sub


Public Sub ClearMySummaryDateLists()
'This is to empty the lists and make sure no double entry happens

aCL_EventIndexList.Clear
EventFilterList.Clear

End Sub
