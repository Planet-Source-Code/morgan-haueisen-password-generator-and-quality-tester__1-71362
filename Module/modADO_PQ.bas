Attribute VB_Name = "modMain"
Option Explicit

'// Valid password characters
Public Const C_Special     As String = "!@#$%^&*()-_=+[]{};':"""",./<>?\|`~"
Public Const C_Lower       As String = "abcdefghijklmnopqrstuvwxyz"
Public Const C_Upper       As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const C_Numbers     As String = "1234567890"

Public gstrDB              As String
Public gstrWordsDB         As String

Public Sub Main()

   Call IsAppRunning
   Call ManifestWrite
   
   '// set default database name and location
   gstrDB = App.Path & "\PwdData.mdb"
   gstrWordsDB = App.Path & "\CommonWords.mdb"
   
   frmUserCheck.Show

End Sub

Public Sub OpenDB(ByRef rConnection As ADODB.Connection, ByVal vDBName As String)

   Set rConnection = New ADODB.Connection
   rConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vDBName & ";Jet OLEDB:Database Password=;Jet OLEDB:Engine Type=5;"

End Sub

Public Sub OpenRS(ByRef rRecordset As ADODB.Recordset, _
                  ByVal oSourceTable As String, _
                  ByRef rConnection As ADODB.Connection, _
                  Optional oCursorType As CursorTypeEnum = adOpenStatic, _
                  Optional oLockType As LockTypeEnum = adLockOptimistic, _
                  Optional ByVal oOptions As Integer = -1)

   Set rRecordset = New ADODB.Recordset
   rRecordset.Open oSourceTable, rConnection, oCursorType, oLockType, oOptions

End Sub

