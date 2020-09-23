Attribute VB_Name = "modMain"
Option Explicit

'// change this number to make each compiled version different
Public Const C_EKEY        As Double = 1.123

'// Valid password characters
Public Const C_Special     As String = "!@#$%^&*()"
Public Const C_Lower       As String = "abcdefghijkkmnopqrstuvwxyz"
Public Const C_Upper       As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const C_Numbers     As String = "1234567890"
Public Const C_StartLength As Integer = 6 '// Must be a number > 5

Public gstrDB              As String
Public gstrWordsDB         As String
Public glngID              As Long

Public Function ADOFindFirst(ByRef MySet As ADODB.Recordset, ByVal Filter As String) As Boolean

  Dim mhRS    As ADODB.Recordset
  Dim mhMatch As Boolean

   On Error GoTo Err_Proc

   Set mhRS = New ADODB.Recordset
   Set mhRS = MySet.Clone
   mhRS.Filter = Filter

   If mhRS.RecordCount > 0 Then
      mhRS.MoveFirst
      MySet.Bookmark = mhRS.Bookmark
      mhMatch = True

   Else
      If MySet.RecordCount > 0 Then
         MySet.MoveLast
         MySet.MoveNext
      End If

      mhMatch = False
   End If

   ADOFindFirst = mhMatch

   Exit Function

Err_Proc:
   On Error Resume Next
   mhRS.Close
   Set mhRS = Nothing
   ADOFindFirst = False

End Function

Public Function ADORecordCount(ByRef MySet As ADODB.Recordset) As Long

  Dim BkMark As Variant
  Dim Rc     As Long

   On Local Error Resume Next

   With MySet
      BkMark = .Bookmark
      .MoveLast
      Rc = .RecordCount
   End With 'MySet

   If Rc = 1 Then
      If IsNull(MySet.Fields(0)) Then
         Rc = 0
      End If
   End If

   ADORecordCount = Rc
   MySet.Bookmark = BkMark
   On Local Error GoTo 0

End Function

Public Function GetCreatedPassPhrase(ByVal vstrSeed As String, _
                                     ByVal vlngLength As Long, _
                                     ByVal vblnUpper As Boolean, _
                                     ByVal vblnNumbers As Boolean, _
                                     ByVal vblnSpecial As Boolean, _
                                     ByRef strPWDO As String) As String

  Dim MyDB     As ADODB.Connection
  Dim strPWD   As String
  Dim lngX     As Long
  Dim lngI     As Long
  Dim strTemp  As String

   '// get seed value
   If LenB(vstrSeed) = 0 Then vstrSeed = App.Title

   For lngX = 1 To Len(vstrSeed)
      lngI = lngI + Asc(Mid$(vstrSeed, lngX, 1))
   Next lngX

   Rnd -1
   Randomize CInt(lngI * C_EKEY)

   '// Open database
   Call OpenDB(MyDB, gstrWordsDB)
   '// Get the number of words in database
   lngI = MyDB.Execute("SELECT Max(CommonWords.ID) AS MaxOfID From CommonWords;")("MaxOfID")
   '// Get random words from database
   Do
      lngX = Int(lngI * Rnd + 1)
      strTemp = MyDB.Execute("SELECT First(CommonWords.Words) AS FirstOfWords From CommonWords WHERE (((CommonWords.ID)=" & _
         CStr(lngX) & "));")("FirstOfWords")
      strPWD = strPWD & strTemp
      strPWDO = strPWDO & " " & strTemp
      If Len(strPWD) >= vlngLength Then Exit Do
   Loop

   MyDB.Close

   '// Modify words based on user options
'''   If vblnSpecial Then
'''      If InStr(1, strPWD, "a") Then strPWD = Replace(strPWD, "a", "@")
'''   End If

   If vblnNumbers Then
      If InStr(1, strPWD, "e") Then strPWD = Replace(strPWD, "e", "3")
      If InStr(1, strPWD, "o") Then strPWD = Replace(strPWD, "o", "0")
   End If

   If vblnUpper Then
      For lngI = 1 To Len(strPWD) Step 3
         Mid$(strPWD, lngI, 1) = UCase$(Mid$(strPWD, lngI, 1))
      Next lngI
   End If

   If vblnSpecial Then strPWD = strPWD & "!"

   '// Return pass phrase
   GetCreatedPassPhrase = strPWD

End Function

Public Function GetCreatedPassword(ByVal vstrSeed As String, _
                                   ByVal vlngLength As Long, _
                                   ByVal vblnUpper As Boolean, _
                                   ByVal vblnNumbers As Boolean, _
                                   ByVal vblnSpecial As Boolean, _
                                   ByVal vblnStartNumber As Boolean, _
                                   ByVal vblnEndNumber As Boolean) As String

  Dim lngX     As Long
  Dim lngI     As Long
  Dim strMap   As String
  Dim strPWD   As String

   '// get seed value
   If LenB(vstrSeed) = 0 Then vstrSeed = App.Title

   For lngX = 1 To Len(vstrSeed)
      lngI = lngI + Asc(Mid$(vstrSeed, lngX, 1))
   Next lngX

   Rnd -1
   Randomize CInt(lngI * C_EKEY)

   '// Build Map and make sure there is at least 1 of every selected option
   If vblnStartNumber Then
      strPWD = Mid$(C_Numbers, GetIndex(C_Numbers), 1)
   End If

   strMap = C_Lower
   strPWD = strPWD & Mid$(C_Lower, GetIndex(C_Lower), 1)

   If vblnUpper Then
      strMap = strMap & C_Upper
      strPWD = strPWD & Mid$(C_Upper, GetIndex(C_Upper), 1)
   End If

   If vblnNumbers Then
      strMap = strMap & C_Numbers & C_Numbers
      strPWD = strPWD & Mid$(C_Numbers, GetIndex(C_Numbers), 1)
   End If

   If vblnSpecial Then
      strMap = strMap & C_Special & C_Special
      strPWD = strPWD & Mid$(C_Special, GetIndex(C_Special), 1)
   End If

   '// Fill in the rest to make the required password length
   For lngX = 1 To vlngLength - Len(strPWD)
      strPWD = strPWD & Mid$(strMap, GetIndex(strMap), 1)
   Next lngX

   '// change the last character to a number if necessary
   If vblnEndNumber Then

      Select Case Right$(strPWD, 1)
      Case Is < "0", Is > "9"
         strPWD = Left$(strPWD, Len(strPWD) - 1) & Mid$(C_Numbers, GetIndex(C_Numbers), 1)
      End Select

   End If

   '// return password
   GetCreatedPassword = strPWD

End Function

Public Function GetIndex(ByRef rstrItems As String) As Integer

   GetIndex = Int((Len(rstrItems) * Rnd + 1))

End Function

Public Sub Main()

   Call IsAppRunning
   Call ManifestWrite
   
   '// set default database name and location
   gstrDB = App.Path & "\PwdData.mdb"
   gstrWordsDB = App.Path & "\CommonWords.mdb"
   
   frmMain.Show

End Sub

Public Sub OpenDB(ByRef rConnection As ADODB.Connection, ByVal vDBName As String)

   Set rConnection = New ADODB.Connection
   rConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vDBName & ";Jet OLEDB:Database Password=;Jet OLEDB:Engine" & _
      " Type=5;"

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

Public Function PWDQuality(ByVal vstrPwd As String, Optional ByVal vblnCheckDic As Boolean = False) As Long
  
  '// Scores a string based on its contents within a range of 0 to 10
  
  Const C_LowerF  As String = C_Lower & "l"
  Const C_SpecF   As String = C_Special & "-_=+[]{};':"""",./<>?\|`~"
  Dim sngPoints   As Single
  Dim sngEntropy  As Single
  Dim lngPwd      As Long
  Dim lngI        As Long
  Dim lngX        As Long
  Dim blnN        As Boolean
  Dim blnS        As Boolean
  Dim blnU        As Boolean
  Dim blnL        As Boolean
  Dim blnA        As Boolean
  Dim aryChar()   As Long
  Dim strChar     As String
  Dim varComm     As Variant
  Dim MyDB        As ADODB.Connection
  Dim MySet       As ADODB.Recordset
 
   lngPwd = Len(vstrPwd)
   
   '// score Character Range
   '// Scores the string based on the range of characters it uses.
   '// A string using only punctuation or numbers
   '// is limited to the number of characters that can be employed.
   '// Better ranges such as alpha numeric are given a higher score
   '// if they mix case than if they are all lowercase with numbers
   '// or all uppercase with numbers.
   '// Ranges which employ all printable characters except the space
   '// are given a higher score than alpha numeric as they are more
   '// complex and the highest of all is given to the range of all
   '// printable characters including the space. So a pass phrase
   '// with spaces would be given a high score.
   
   For lngI = 1 To lngPwd
      If Not blnN Then '// numbers
         If InStr(1, C_Numbers, Mid$(vstrPwd, lngI, 1)) Then
            blnN = True
         End If
      End If
      
      If Not blnS Then '// special characters
         If InStr(1, C_SpecF, Mid$(vstrPwd, lngI, 1)) Then
            blnS = True
         End If
      End If
      
      If Not blnU Then '// upper case letters
         If InStr(1, C_Upper, Mid$(vstrPwd, lngI, 1)) Then
            blnU = True
         End If
      End If
      
      If Not blnL Then '// lower case letters
         If InStr(1, C_LowerF, Mid$(vstrPwd, lngI, 1)) Then
            blnL = True
         End If
      End If
      
      If Not blnA Then '// spaces
         If Mid$(vstrPwd, lngI, 1) = " " Then
            blnA = True
         End If
      End If
   Next lngI
   
   If blnA And Not (blnN Or blnL Or blnU Or blnS) Then '// only [ ]
      sngPoints = 0
      
   Else
      If blnS And Not (blnN Or blnL Or blnU) Then '// only [*$/]
         sngPoints = 1
      ElseIf blnN And Not (blnL Or blnU Or blnS) Then '// only [0-9]
         sngPoints = 0.5
      ElseIf (blnN And blnS) And Not (blnL And blnU) Then '// only [0-9*$/]
         sngPoints = 1.5
      ElseIf (blnU Or blnL) And Not (blnS Or blnN) Then '// only [a-zA-Z]
         sngPoints = 2
         If Not blnU Or Not blnL Then sngPoints = sngPoints - 0.5
      
      Else
         If (blnL And blnN) And Not (blnU Or blnS) Then '//[a-z0-9]
            sngPoints = 3
         ElseIf (blnU And blnN) And Not (blnL Or blnS) Then '//[A-Z0-9]
            sngPoints = 3
         ElseIf (blnL And blnU And blnN) And Not (blnS) Then '// [a-zA-Z0-9]
            sngPoints = 4
         ElseIf blnN And blnL And blnU And blnS Then '//[a-zA-Z0-9*$/]
            sngPoints = 5
         ElseIf (blnN And blnS) And (blnU Or blnL) Then '// [0-9*$/] and ([a-z] or [A-Z])
            sngPoints = 3
         End If
      End If
      If blnA Then sngPoints = sngPoints + 0.5
   End If

   
   '// score Entropy
   '// Measures the entropy of the characters contained within the word
   '// for example a string with 5 out of 6 characters the same would be
   '// less secure than a password with 6 different characters.
   '// Uses Shannon Entropy formula to measure the entropy of the word
   ReDim aryChar(1 To 1)
   strChar = Left$(vstrPwd, 1)
   aryChar(1) = 1
   
   For lngI = 2 To lngPwd
      lngX = InStr(1, strChar, Mid$(vstrPwd, lngI, 1))
      If lngX = 0 Then
         lngX = UBound(aryChar) + 1
         ReDim Preserve aryChar(1 To lngX) As Long
         strChar = strChar & Mid$(vstrPwd, lngI, 1)
      End If
      aryChar(lngX) = aryChar(lngX) + 1
   Next lngI

   For lngI = 1 To UBound(aryChar)
      sngEntropy = sngEntropy + ((aryChar(lngI) / lngPwd) * (Log((aryChar(lngI) / lngPwd) ^ 2)))
   Next lngI
   sngPoints = sngPoints + (sngEntropy * -1)
   Erase aryChar
   
   
   '// score Length
   '// passwords with a length = 8 is neutral
   '// > 8 add score; 8 < deduct score
   sngPoints = sngPoints + ((lngPwd - 8) * 0.05)
   If lngPwd <= 6 Then sngPoints = sngPoints - 2
   
  
   '// deductions for character repetitions (abcabc, aaaaa, 121212, etc.)
   blnN = False
   For lngX = 1 To lngPwd \ 2
      For lngI = lngX + 1 To lngPwd Step lngX
         If Mid$(vstrPwd, 1, lngX) = Mid$(vstrPwd, lngI, lngX) Then
            sngPoints = sngPoints - 0.5
            blnN = True
            Exit For
         End If
      Next lngI
      If blnN Then Exit For
   Next lngX
   
   '// the following checks need to have a common case
   vstrPwd = LCase$(vstrPwd)
   
   '// deductions for use of Common passwords
   varComm = Array("!@#", "@#$", "#$%", "$%^", "%^&", "^&*", "&*(", "*()", "000", "007", "246", "249", "1022", "10sne1", "111", "1212", "1225", "123123", "1234", "123abc", "123go", "1313", "13579", "14430", "1701", "thx1138", "1928", "1951", "1a2b3c", "1p2o3i", "1q2w3e", "1qw23e", "1sanjose", "2112", "21122112", "222", "welcome", "369", "444", "4runner", "5252", "54321", "555", "5683", "654321", "666", "6969", "777", "80486", "8675309", "888", "90210", "911", "92072", "999", "a12345", "a1b2c3", "a1b2c3d4", "aaa", "admin", "aaron", "abby", "abc", "asdf", "charlie", "hammer", "happy", "ib6ub9", "icecream", "kermit", "pass", "password", "qwerty", "letmein", "dragon", "master", "monkey", "mustang", "myspace", "money", "blink", "god", "sex", "love", "soccer", "jordan", "football", "baseball", "princess", "shadow", "slipknot", "liverpool", "link182", "super", "xanadu", "xavier", "xcountry", "xfiles", "xxx", "yaco")
   For lngI = 0 To UBound(varComm)
      If InStr(1, vstrPwd, varComm(lngI)) Then
         sngPoints = sngPoints - 0.5
         Exit For
      End If
   Next lngI
   Set varComm = Nothing
   
   '// deductions for use of Common Words (optional)
   '// This helps to compensate for words or phrases that might appear
   '// secure but are commonly know and therefore less secure.
   '// compaired against a list of 2063 words
   If vblnCheckDic Then
      If LenB(Dir$(gstrWordsDB)) > 0 Then
         '// Open database
         Call OpenDB(MyDB, gstrWordsDB)
         Call OpenRS(MySet, "SELECT CommonWords.Words From CommonWords", MyDB)
         Do
            If InStr(1, vstrPwd, MySet.Fields("Words")) > 0 Then
               sngPoints = sngPoints - 0.5
            End If
            MySet.MoveNext
         Loop Until MySet.EOF
         MySet.Close
         MyDB.Close
      End If
   End If

   PWDQuality = sngPoints * 10
  
End Function
                 
