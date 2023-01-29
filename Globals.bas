Attribute VB_Name = "Module1"
Global sRank        As String
Global sLevel       As String
Global sCharName    As String
Global sPoints      As String
Global sSex         As String
Global sType        As String
Global sClass       As String
Global sMastered    As String
Global sCorpName    As String
Global sCEO         As String
Global sCorpExp     As String
Global sCorpAlign   As String
Global sLastPlayed  As String
Global sLastPlayed2 As String
Global sDate        As String
Global sKills       As String
Global sExp         As String
Global sStatus      As String

Global Plvl         As String
Global Slvl         As String
Global dmModName    As String
Global dmLastPlay   As String

Global xSourceFile  As String
Global xDestFile    As String
Global sDestFile    As String

Global xGarbage     As String

Global intEmpFileNbr As Integer
Global Line1 As String
Global Line2 As String
Global Line3 As String
Global Line4 As String
Global LineA As String
Global LineB As String
Global Lines As String
Global strTest As String
Global strArray() As String
Global lvlArray() As String
Global intCount As Integer


Global RSCount  As Integer

Global DBCon As ADODB.Connection
Global Rs As ADODB.Recordset
Global DBQuery As ADODB.Recordset


Public Sub ConnecttoMYSQL()

'Create a connection to the database
Set DBCon = New ADODB.Connection
' DBCon.CursorLocation = adUseClient
DBCon.ConnectionString = "DRIVER={MySQL ODBC 5.2 ANSI Driver}; Server=10.10.10.176;Database=sbbs;User=sbbsrw;Password=sbbsrw;Option=" & 1 + 2 + 8 + 32 + 2048 + 16384  ' 3"
DBCon.Open


End Sub


'Input #intEmpFileNbr, strEmpName
'        If strEmpName = "" Then
'            Call CleanUpLines(Lines)
'            Lines = ""
'        Else
'            Lines = Lines + strEmpName
'        End If
