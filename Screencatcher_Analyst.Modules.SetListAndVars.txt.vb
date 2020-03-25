Option Compare Database
Option Explicit

Public nonTrialsList As Object ' https://excelmacromastery.com/vba-arraylist/
Public trialsOnlyList As Object 'Declare the Object handle for use
Public allColumnsList As Object 'Declare the Object handle for use
Public allTablesList As Object 'Declare the Object handle for use
Public PreAT_DatesVarList As Object 'Declare the Object handle for use
Public PreFCT_DatesVarList As Object 'Declare the Object handle for use
Public beanBT As String 'This is a varible holding the name reference to the _BT table used in the SQL statments
Public beanAT As String 'This is a varible holding the name reference to the _AT table used in the SQL statments
Public beanFCT As String 'This is a varible holding the name reference to the _FCT table used in the SQL statments
Public beanOWLD As String 'This is a varible holding the name reference to the _OWLD table used in the SQL statments
Public beanFinal As String 'This is a varible holding the name reference to the _Final table used in the SQL statments
Public columnBT As String ' name reference to the BT column used in the SQL statments
Public columnAT As String ' name reference to the AT column used in the SQL statments
Public columnFCT As String ' name reference to the FCT column used in the SQL statments
Public columnOWLD As String ' name reference to the OWLD column used in the SQL statments
Public columnFinal As String ' name reference to the Final column used in the SQL statments
Public curHullNum As String ' This is the current ship hull number that is used in the table names
Public CurrentTable As String ' This is the current table that the module is modifing
Public SparseRefTable As String ' This is the table that the module is using to compare records columns values


Sub AddingToMyDateLists()
'https://excelmacromastery.com/vba-arraylist/

'Hull number currently in use
curHullNum = "LPD25"  ' This is the current ship hull number that is used in the table names

'Tables used for reference to Events
beanBT = "2013/08/19_LPD25_BT" ' name reference to the _BT table used in the SQL statments
beanAT = "2013/09/21_LPD25_AT" ' name reference to the _AT table used in the SQL statments
beanFCT = "2014/11/09_LPD25_FCT" ' name reference to the _FCT table used in the SQL statments
beanOWLD = "2015/07/02_LPD25_OWLD" ' name reference to the _OWLD table used in the SQL statments
beanFinal = "2020/02/03_LPD25_Final" ' name reference to the _Final table used in the SQL statments

'Columns used to reference the Events
columnBT = "2013/08/19" ' name reference to the BT column used in the SQL statments
columnAT = "2013/09/21" ' name reference to the AT column used in the SQL statments
columnFCT = "2014/11/09" ' name reference to the FCT column used in the SQL statments
columnOWLD = "2015/07/02" ' name reference to the OWLD column used in the SQL statments
columnFinal = "2020/02/03" ' name reference to the Final column used in the SQL statments

'Columns used to reset values prior to the AT Event
'Dim PreAT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreAT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreAT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreAT_DatesVarList
' Add items
PreAT_DatesVarList.Add "2013/08/19" ' BT, Called from SetLateAdds_TrialCards as PreAT_DatesVarList(0)
PreAT_DatesVarList.Add "2013/08/21" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(1)
PreAT_DatesVarList.Add "2013/09/06" ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(2)
 ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(3)
 ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(4)
 ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(5)
 ' Called from SetLateAdds_TrialCards as PreAT_DatesVarList(6)

'Columns used to reset values prior to the FCT Event
'Dim PreFCT_DatesVarList As Object ' This Declares the object handle, without any properties or methods
'Dim PreFCT_DatesVarList As New ArrayList ' This Declares the object handle with early binding
Set PreFCT_DatesVarList = CreateObject("System.Collections.ArrayList") ' This is late binding
'PreFCT_DatesVarList
' Add items
PreFCT_DatesVarList.Add "2013/08/19" ' BT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(0)
PreFCT_DatesVarList.Add "2013/08/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(1)
PreFCT_DatesVarList.Add "2013/09/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(2)
PreFCT_DatesVarList.Add "2013/09/21" ' AT, Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(3)
PreFCT_DatesVarList.Add "2013/09/26" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(4)
PreFCT_DatesVarList.Add "2013/10/04" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(5)
PreFCT_DatesVarList.Add "2013/10/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(6)
PreFCT_DatesVarList.Add "2013/11/02" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(7)
PreFCT_DatesVarList.Add "2013/11/08" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(8)
PreFCT_DatesVarList.Add "2013/11/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(9)
PreFCT_DatesVarList.Add "2013/11/23" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(10)
PreFCT_DatesVarList.Add "2013/12/01" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(11)
PreFCT_DatesVarList.Add "2013/12/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(12)
PreFCT_DatesVarList.Add "2013/12/13" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(13)
PreFCT_DatesVarList.Add "2013/12/20" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(14)
PreFCT_DatesVarList.Add "2013/12/27" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(15)
PreFCT_DatesVarList.Add "2014/01/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(16)
PreFCT_DatesVarList.Add "2014/01/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(17)
PreFCT_DatesVarList.Add "2014/01/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(18)
PreFCT_DatesVarList.Add "2014/01/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(19)
PreFCT_DatesVarList.Add "2014/01/31" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(20)
PreFCT_DatesVarList.Add "2014/02/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(21)
PreFCT_DatesVarList.Add "2014/02/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(22)
PreFCT_DatesVarList.Add "2014/02/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(23)
PreFCT_DatesVarList.Add "2014/02/28" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(24)
PreFCT_DatesVarList.Add "2014/03/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(25)
PreFCT_DatesVarList.Add "2014/03/07" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(26)
PreFCT_DatesVarList.Add "2014/03/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(27)
PreFCT_DatesVarList.Add "2014/03/21" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(28)
PreFCT_DatesVarList.Add "2014/03/28" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(29)
PreFCT_DatesVarList.Add "2014/04/04" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(30)
PreFCT_DatesVarList.Add "2014/04/11" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(31)
PreFCT_DatesVarList.Add "2014/04/18" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(32)
PreFCT_DatesVarList.Add "2014/04/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(33)
PreFCT_DatesVarList.Add "2014/05/01" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(34)
PreFCT_DatesVarList.Add "2014/05/09" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(35)
PreFCT_DatesVarList.Add "2014/05/16" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(36)
PreFCT_DatesVarList.Add "2014/05/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(37)
PreFCT_DatesVarList.Add "2014/05/30" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(38)
PreFCT_DatesVarList.Add "2014/06/06" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(39)
PreFCT_DatesVarList.Add "2014/06/13" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(40)
PreFCT_DatesVarList.Add "2014/06/20" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(41)
PreFCT_DatesVarList.Add "2014/06/27" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(42)
PreFCT_DatesVarList.Add "2014/07/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(43)
PreFCT_DatesVarList.Add "2014/07/12" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(44)
PreFCT_DatesVarList.Add "2014/07/18" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(45)
PreFCT_DatesVarList.Add "2014/07/25" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(46)
PreFCT_DatesVarList.Add "2014/07/31" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(47)
PreFCT_DatesVarList.Add "2014/08/08" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(48)
PreFCT_DatesVarList.Add "2014/08/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(49)
PreFCT_DatesVarList.Add "2014/08/22" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(50)
PreFCT_DatesVarList.Add "2014/08/29" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(51)
PreFCT_DatesVarList.Add "2014/09/05" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(52)
PreFCT_DatesVarList.Add "2014/09/14" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(53)
PreFCT_DatesVarList.Add "2014/09/19" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(54)
PreFCT_DatesVarList.Add "2014/09/26" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(55)
PreFCT_DatesVarList.Add "2014/10/03" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(56)
PreFCT_DatesVarList.Add "2014/10/10" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(57)
PreFCT_DatesVarList.Add "2014/10/17" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(58)
PreFCT_DatesVarList.Add "2014/10/24" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(59)
PreFCT_DatesVarList.Add "2014/10/31" ' Called from SetLateAdds_TrialCards as PreFCT_DatesVarList(60)

'Date values used in queries list

'Dim nonTrialsList As Object ' This Declares the object handle, without any properties or methods
'Dim nonTrialsList As New ArrayList ' This Declares the object handle with early binding
Set nonTrialsList = CreateObject("System.Collections.ArrayList") ' This is late binding
'nonTrialsList
' Add items
nonTrialsList.Add "2013/08/21"
nonTrialsList.Add "2013/09/06"
nonTrialsList.Add "2013/09/26"
nonTrialsList.Add "2013/10/04"
nonTrialsList.Add "2013/10/25"
nonTrialsList.Add "2013/11/02"
nonTrialsList.Add "2013/11/08"
nonTrialsList.Add "2013/11/16"
nonTrialsList.Add "2013/11/23"
nonTrialsList.Add "2013/12/01"
nonTrialsList.Add "2013/12/06"
nonTrialsList.Add "2013/12/13"
nonTrialsList.Add "2013/12/20"
nonTrialsList.Add "2013/12/27"
nonTrialsList.Add "2014/01/03"
nonTrialsList.Add "2014/01/10"
nonTrialsList.Add "2014/01/17"
nonTrialsList.Add "2014/01/24"
nonTrialsList.Add "2014/01/31"
nonTrialsList.Add "2014/02/07"
nonTrialsList.Add "2014/02/14"
nonTrialsList.Add "2014/02/21"
nonTrialsList.Add "2014/02/28"
nonTrialsList.Add "2014/03/06"
nonTrialsList.Add "2014/03/07"
nonTrialsList.Add "2014/03/14"
nonTrialsList.Add "2014/03/21"
nonTrialsList.Add "2014/03/28"
nonTrialsList.Add "2014/04/04"
nonTrialsList.Add "2014/04/11"
nonTrialsList.Add "2014/04/18"
nonTrialsList.Add "2014/04/25"
nonTrialsList.Add "2014/05/01"
nonTrialsList.Add "2014/05/09"
nonTrialsList.Add "2014/05/16"
nonTrialsList.Add "2014/05/25"
nonTrialsList.Add "2014/05/30"
nonTrialsList.Add "2014/06/06"
nonTrialsList.Add "2014/06/13"
nonTrialsList.Add "2014/06/20"
nonTrialsList.Add "2014/06/27"
nonTrialsList.Add "2014/07/03"
nonTrialsList.Add "2014/07/12"
nonTrialsList.Add "2014/07/18"
nonTrialsList.Add "2014/07/25"
nonTrialsList.Add "2014/07/31"
nonTrialsList.Add "2014/08/08"
nonTrialsList.Add "2014/08/17"
nonTrialsList.Add "2014/08/22"
nonTrialsList.Add "2014/08/29"
nonTrialsList.Add "2014/09/05"
nonTrialsList.Add "2014/09/14"
nonTrialsList.Add "2014/09/19"
nonTrialsList.Add "2014/09/26"
nonTrialsList.Add "2014/10/03"
nonTrialsList.Add "2014/10/10"
nonTrialsList.Add "2014/10/17"
nonTrialsList.Add "2014/10/24"
nonTrialsList.Add "2014/10/31"
nonTrialsList.Add "2014/11/14"
nonTrialsList.Add "2014/11/21"
nonTrialsList.Add "2014/11/26"
nonTrialsList.Add "2014/12/05"
nonTrialsList.Add "2014/12/12"
nonTrialsList.Add "2014/12/19"
nonTrialsList.Add "2014/12/29"
nonTrialsList.Add "2015/01/02"
nonTrialsList.Add "2015/01/09"
nonTrialsList.Add "2015/01/16"
nonTrialsList.Add "2015/01/23"
nonTrialsList.Add "2015/01/31"
nonTrialsList.Add "2015/02/06"
nonTrialsList.Add "2015/02/13"
nonTrialsList.Add "2015/02/20"
nonTrialsList.Add "2015/02/27"
nonTrialsList.Add "2015/03/07"
nonTrialsList.Add "2015/03/13"
nonTrialsList.Add "2015/03/20"
nonTrialsList.Add "2015/03/27"
nonTrialsList.Add "2015/04/05"
nonTrialsList.Add "2015/04/10"
nonTrialsList.Add "2015/04/17"
nonTrialsList.Add "2015/04/24"
nonTrialsList.Add "2015/05/01"
nonTrialsList.Add "2015/05/08"
nonTrialsList.Add "2015/05/15"
nonTrialsList.Add "2015/05/20"
nonTrialsList.Add "2015/05/27"
nonTrialsList.Add "2015/05/29"
nonTrialsList.Add "2015/06/05"
nonTrialsList.Add "2015/06/12"
nonTrialsList.Add "2015/06/19"
nonTrialsList.Add "2015/06/24"
nonTrialsList.Add "2015/07/10"
nonTrialsList.Add "2015/07/17"
nonTrialsList.Add "2015/07/31"
nonTrialsList.Add "2015/12/03"
nonTrialsList.Add "2016/04/08"
nonTrialsList.Add "2016/08/16"
nonTrialsList.Add "2016/11/16"
nonTrialsList.Add "2019/04/01"
nonTrialsList.Add "2019/09/05"

'Dim trialsOnlyList As Object
'Dim trialsOnlyList As New ArrayList
Set trialsOnlyList = CreateObject("System.Collections.ArrayList")
'trialsOnlyList
' Add items
trialsOnlyList.Add "2013/08/19" ' BT
trialsOnlyList.Add "2013/09/21" ' AT
trialsOnlyList.Add "2014/11/09" ' FCT
trialsOnlyList.Add "2015/07/02" ' OWLD
trialsOnlyList.Add "2020/02/03" ' Final

'Dim allColumnsList As Object
'Dim allColumnsList As New ArrayList
Set allColumnsList = CreateObject("System.Collections.ArrayList")
'allColumnsList
' Add items
allColumnsList.Add "2013/08/19" ' BT
allColumnsList.Add "2013/08/21"
allColumnsList.Add "2013/09/06"
allColumnsList.Add "2013/09/21" ' AT
allColumnsList.Add "2013/09/26"
allColumnsList.Add "2013/10/04"
allColumnsList.Add "2013/10/25"
allColumnsList.Add "2013/11/02"
allColumnsList.Add "2013/11/08"
allColumnsList.Add "2013/11/16"
allColumnsList.Add "2013/11/23"
allColumnsList.Add "2013/12/01"
allColumnsList.Add "2013/12/06"
allColumnsList.Add "2013/12/13"
allColumnsList.Add "2013/12/20"
allColumnsList.Add "2013/12/27"
allColumnsList.Add "2014/01/03"
allColumnsList.Add "2014/01/10"
allColumnsList.Add "2014/01/17"
allColumnsList.Add "2014/01/24"
allColumnsList.Add "2014/01/31"
allColumnsList.Add "2014/02/07"
allColumnsList.Add "2014/02/14"
allColumnsList.Add "2014/02/21"
allColumnsList.Add "2014/02/28"
allColumnsList.Add "2014/03/06"
allColumnsList.Add "2014/03/07"
allColumnsList.Add "2014/03/14"
allColumnsList.Add "2014/03/21"
allColumnsList.Add "2014/03/28"
allColumnsList.Add "2014/04/04"
allColumnsList.Add "2014/04/11"
allColumnsList.Add "2014/04/18"
allColumnsList.Add "2014/04/25"
allColumnsList.Add "2014/05/01"
allColumnsList.Add "2014/05/09"
allColumnsList.Add "2014/05/16"
allColumnsList.Add "2014/05/25"
allColumnsList.Add "2014/05/30"
allColumnsList.Add "2014/06/06"
allColumnsList.Add "2014/06/13"
allColumnsList.Add "2014/06/20"
allColumnsList.Add "2014/06/27"
allColumnsList.Add "2014/07/03"
allColumnsList.Add "2014/07/12"
allColumnsList.Add "2014/07/18"
allColumnsList.Add "2014/07/25"
allColumnsList.Add "2014/07/31"
allColumnsList.Add "2014/08/08"
allColumnsList.Add "2014/08/17"
allColumnsList.Add "2014/08/22"
allColumnsList.Add "2014/08/29"
allColumnsList.Add "2014/09/05"
allColumnsList.Add "2014/09/14"
allColumnsList.Add "2014/09/19"
allColumnsList.Add "2014/09/26"
allColumnsList.Add "2014/10/03"
allColumnsList.Add "2014/10/10"
allColumnsList.Add "2014/10/17"
allColumnsList.Add "2014/10/24"
allColumnsList.Add "2014/10/31"
allColumnsList.Add "2014/11/09" ' FCT
allColumnsList.Add "2014/11/14"
allColumnsList.Add "2014/11/21"
allColumnsList.Add "2014/11/26"
allColumnsList.Add "2014/12/05"
allColumnsList.Add "2014/12/12"
allColumnsList.Add "2014/12/19"
allColumnsList.Add "2014/12/29"
allColumnsList.Add "2015/01/02"
allColumnsList.Add "2015/01/09"
allColumnsList.Add "2015/01/16"
allColumnsList.Add "2015/01/23"
allColumnsList.Add "2015/01/31"
allColumnsList.Add "2015/02/06"
allColumnsList.Add "2015/02/13"
allColumnsList.Add "2015/02/20"
allColumnsList.Add "2015/02/27"
allColumnsList.Add "2015/03/07"
allColumnsList.Add "2015/03/13"
allColumnsList.Add "2015/03/20"
allColumnsList.Add "2015/03/27"
allColumnsList.Add "2015/04/05"
allColumnsList.Add "2015/04/10"
allColumnsList.Add "2015/04/17"
allColumnsList.Add "2015/04/24"
allColumnsList.Add "2015/05/01"
allColumnsList.Add "2015/05/08"
allColumnsList.Add "2015/05/15"
allColumnsList.Add "2015/05/20"
allColumnsList.Add "2015/05/27"
allColumnsList.Add "2015/05/29"
allColumnsList.Add "2015/06/05"
allColumnsList.Add "2015/06/12"
allColumnsList.Add "2015/06/19"
allColumnsList.Add "2015/06/24"
allColumnsList.Add "2015/07/02" ' OWLD
allColumnsList.Add "2015/07/10"
allColumnsList.Add "2015/07/17"
allColumnsList.Add "2015/07/31"
allColumnsList.Add "2015/12/03"
allColumnsList.Add "2016/04/08"
allColumnsList.Add "2016/08/16"
allColumnsList.Add "2016/11/16"
allColumnsList.Add "2019/04/01"
allColumnsList.Add "2019/09/05"
allColumnsList.Add "2020/02/03" ' Final

'Dim allTablesList As Object
'Dim allTablesList As New ArrayList
Set allTablesList = CreateObject("System.Collections.ArrayList")
'allTablesList
' Add items
allTablesList.Add "2013/08/19_LPD25_BT"
allTablesList.Add "2013/08/21_LPD25"
allTablesList.Add "2013/09/06_LPD25"
allTablesList.Add "2013/09/21_LPD25_AT"
allTablesList.Add "2013/09/26_LPD25"
allTablesList.Add "2013/10/04_LPD25"
allTablesList.Add "2013/10/25_LPD25"
allTablesList.Add "2013/11/02_LPD25"
allTablesList.Add "2013/11/08_LPD25"
allTablesList.Add "2013/11/16_LPD25"
allTablesList.Add "2013/11/23_LPD25"
allTablesList.Add "2013/12/01_LPD25"
allTablesList.Add "2013/12/06_LPD25"
allTablesList.Add "2013/12/13_LPD25"
allTablesList.Add "2013/12/20_LPD25"
allTablesList.Add "2013/12/27_LPD25"
allTablesList.Add "2014/01/03_LPD25"
allTablesList.Add "2014/01/10_LPD25"
allTablesList.Add "2014/01/17_LPD25"
allTablesList.Add "2014/01/24_LPD25"
allTablesList.Add "2014/01/31_LPD25"
allTablesList.Add "2014/02/07_LPD25"
allTablesList.Add "2014/02/14_LPD25"
allTablesList.Add "2014/02/21_LPD25"
allTablesList.Add "2014/02/28_LPD25"
allTablesList.Add "2014/03/06_LPD25"
allTablesList.Add "2014/03/07_LPD25"
allTablesList.Add "2014/03/14_LPD25"
allTablesList.Add "2014/03/21_LPD25"
allTablesList.Add "2014/03/28_LPD25"
allTablesList.Add "2014/04/04_LPD25"
allTablesList.Add "2014/04/11_LPD25"
allTablesList.Add "2014/04/18_LPD25"
allTablesList.Add "2014/04/25_LPD25"
allTablesList.Add "2014/05/01_LPD25"
allTablesList.Add "2014/05/09_LPD25"
allTablesList.Add "2014/05/16_LPD25"
allTablesList.Add "2014/05/25_LPD25"
allTablesList.Add "2014/05/30_LPD25"
allTablesList.Add "2014/06/06_LPD25"
allTablesList.Add "2014/06/13_LPD25"
allTablesList.Add "2014/06/20_LPD25"
allTablesList.Add "2014/06/27_LPD25"
allTablesList.Add "2014/07/03_LPD25"
allTablesList.Add "2014/07/12_LPD25"
allTablesList.Add "2014/07/18_LPD25"
allTablesList.Add "2014/07/25_LPD25"
allTablesList.Add "2014/07/31_LPD25"
allTablesList.Add "2014/08/08_LPD25"
allTablesList.Add "2014/08/17_LPD25"
allTablesList.Add "2014/08/22_LPD25"
allTablesList.Add "2014/08/29_LPD25"
allTablesList.Add "2014/09/05_LPD25"
allTablesList.Add "2014/09/14_LPD25"
allTablesList.Add "2014/09/19_LPD25"
allTablesList.Add "2014/09/26_LPD25"
allTablesList.Add "2014/10/03_LPD25"
allTablesList.Add "2014/10/10_LPD25"
allTablesList.Add "2014/10/17_LPD25"
allTablesList.Add "2014/10/24_LPD25"
allTablesList.Add "2014/10/31_LPD25"
allTablesList.Add "2014/11/09_LPD25_FCT"
allTablesList.Add "2014/11/14_LPD25"
allTablesList.Add "2014/11/21_LPD25"
allTablesList.Add "2014/11/26_LPD25"
allTablesList.Add "2014/12/05_LPD25"
allTablesList.Add "2014/12/12_LPD25"
allTablesList.Add "2014/12/19_LPD25"
allTablesList.Add "2014/12/29_LPD25"
allTablesList.Add "2015/01/02_LPD25"
allTablesList.Add "2015/01/09_LPD25"
allTablesList.Add "2015/01/16_LPD25"
allTablesList.Add "2015/01/23_LPD25"
allTablesList.Add "2015/01/31_LPD25"
allTablesList.Add "2015/02/06_LPD25"
allTablesList.Add "2015/02/13_LPD25"
allTablesList.Add "2015/02/20_LPD25"
allTablesList.Add "2015/02/27_LPD25"
allTablesList.Add "2015/03/07_LPD25"
allTablesList.Add "2015/03/13_LPD25"
allTablesList.Add "2015/03/20_LPD25"
allTablesList.Add "2015/03/27_LPD25"
allTablesList.Add "2015/04/05_LPD25"
allTablesList.Add "2015/04/10_LPD25"
allTablesList.Add "2015/04/17_LPD25"
allTablesList.Add "2015/04/24_LPD25"
allTablesList.Add "2015/05/01_LPD25"
allTablesList.Add "2015/05/08_LPD25"
allTablesList.Add "2015/05/15_LPD25"
allTablesList.Add "2015/05/20_LPD25"
allTablesList.Add "2015/05/27_LPD25"
allTablesList.Add "2015/05/29_LPD25"
allTablesList.Add "2015/06/05_LPD25"
allTablesList.Add "2015/06/12_LPD25"
allTablesList.Add "2015/06/19_LPD25"
allTablesList.Add "2015/06/24_LPD25"
allTablesList.Add "2015/07/02_LPD25_OWLD"
allTablesList.Add "2015/07/10_LPD25"
allTablesList.Add "2015/07/17_LPD25"
allTablesList.Add "2015/07/31_LPD25"
allTablesList.Add "2015/12/03_LPD25"
allTablesList.Add "2016/04/08_LPD25"
allTablesList.Add "2016/08/16_LPD25"
allTablesList.Add "2016/11/16_LPD25"
allTablesList.Add "2019/04/01_LPD25"
allTablesList.Add "2019/09/05_LPD25"
allTablesList.Add "2020/02/03_LPD25_Final"
    
' Insert to first position
'   allTablesList.Insert 0, "2020/01/22_LPD23"
' Print a specfic item
'   allTablesList(i)
' Sort
'allTablesList.Sort
'   Debug.Print vbCrLf & "Sorted Ascending"
    ' Add this sub from "Reading through the items" section
'   PrintToImmediateWindow allTablesList
' Reverse sort
'   allTablesList.Reverse
'   Debug.Print vbCrLf & "Sorted Descending"
'   PrintToImmediateWindow allTablesList
' Remove all item
'   allTablesList.Clear

End Sub

Public Sub ClearMyDateLists()
'This is to empty the lists and make sure no double entry happens

PreAT_DatesVarList.Clear
PreFCT_DatesVarList.Clear
nonTrialsList.Clear
trialsOnlyList.Clear
allColumnsList.Clear
allTablesList.Clear

End Sub
