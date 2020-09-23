VERSION 5.00
Begin VB.Form frmUserCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check the Quality of Your Password"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   9945
   StartUpPosition =   1  'CenterOwner
   Begin GenPwd.CandyButton cmdQuit 
      Cancel          =   -1  'True
      Height          =   750
      Left            =   8895
      TabIndex        =   7
      Top             =   765
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Caption         =   "Exit"
      PicHighLight    =   0   'False
      CaptionHighLight=   0   'False
      CaptionHighLightColor=   0
      ForeColor       =   16777215
      Picture         =   "frmUserCheck.frx":1601A
      PictureAlignment=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   51
      ColorButtonUp   =   0
      ColorButtonDown =   102
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   -1  'True
      ColorScheme     =   8
      CornerRadius    =   26
      UserCornerRadius=   12
      DisabledPicMode =   0
      UseGREY         =   0   'False
      UseMaskColor    =   -1  'True
      MaskColor       =   -2147483633
      ButtonBehaviour =   0
      ShowFocus       =   0   'False
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "Hide"
      Height          =   300
      Left            =   8535
      TabIndex        =   6
      Top             =   345
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin GenPwd.LynxGrid grdResults 
      Height          =   2145
      Left            =   1395
      TabIndex        =   3
      Top             =   1440
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   3784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      BackColorEvenRowsEnabled=   0   'False
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectColor  =   9895934
      GridLines       =   3
      ThemeColor      =   0
      ThemeStyle      =   3
      Appearance      =   0
      ColumnHeaderSmall=   -1  'True
      ScrollBarStyle  =   1
      AllowColumnResizing=   -1  'True
      AllowWordWrap   =   -1  'True
      HotHeaderTracking=   0   'False
   End
   Begin VB.TextBox txtKey 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2130
      MaxLength       =   255
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   285
      Width           =   6240
   End
   Begin GenPwd.CandyButton cmdLoad 
      Default         =   -1  'True
      Height          =   750
      Left            =   7320
      TabIndex        =   2
      Top             =   765
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Caption         =   "Check"
      PicHighLight    =   0   'False
      CaptionHighLight=   0   'False
      CaptionHighLightColor=   0
      Picture         =   "frmUserCheck.frx":16724
      PictureAlignment=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16756583
      ColorButtonUp   =   16761247
      ColorButtonDown =   13743257
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   -1  'True
      ColorScheme     =   0
      CornerRadius    =   12
      UserCornerRadius=   0
      DisabledPicMode =   0
      UseGREY         =   0   'False
      UseMaskColor    =   -1  'True
      MaskColor       =   -2147483633
      ButtonBehaviour =   0
      ShowFocus       =   0   'False
   End
   Begin GenPwd.Frame3D fraQuality 
      Height          =   285
      Left            =   1395
      Top             =   1125
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   503
      BorderType      =   9
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   4
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   -1  'True
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   0
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   285
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmUserCheck.frx":16ABE
      Picture         =   "frmUserCheck.frx":16ADA
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Tahoma"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "0%"
      UseMnemonic     =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   $"frmUserCheck.frx":16AF6
      Height          =   1995
      Left            =   7080
      TabIndex        =   5
      Top             =   1650
      Width           =   2640
   End
   Begin VB.Label lblQual 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   765
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Enter your Password "
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   375
      Width           =   1890
   End
End
Attribute VB_Name = "frmUserCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHide_Click()
   
   If chkHide.Value = vbChecked Then
      txtKey.PasswordChar = "*"
   Else
      txtKey.PasswordChar = ""
   End If
   
End Sub

Private Sub cmdLoad_Click()

   '// Get password quality
   Call GetPWDQuality(txtKey.Text)

End Sub

Private Sub cmdQuit_Click()
   
   Unload Me

End Sub

Private Sub Form_Load()
   
   With grdResults
      .AddColumn "Reason", .VisibleWidth - 800, , , , , , True
      .AddColumn "Score", 800, lgAlignRightCenter, lgNumeric, "0.00"
   End With
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmUserCheck = Nothing

End Sub

Private Function GetPWDQuality(ByVal vstrPwd As String) As Long
  
  '// Scores a string based on its contents within a range of 0 to 10
  
  Const LGREEN    As Long = &HD5FFD1
  Const LRED      As Long = &HD1E5FF
  Const C_Prefix  As String = "Password Quality is "
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
  Dim lngR        As Long
 
   Screen.MousePointer = vbHourglass
   DoEvents
   lngPwd = Len(vstrPwd)
   grdResults.Clear
   
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
            lngR = grdResults.AddItem("Contains Numbers")
            grdResults.RowBackColor(lngR) = LGREEN
         End If
      End If
      
      If Not blnS Then '// special characters
         If InStr(1, C_Special, Mid$(vstrPwd, lngI, 1)) Then
            blnS = True
            lngR = grdResults.AddItem("Contains Specials")
            grdResults.RowBackColor(lngR) = LGREEN
         End If
      End If
      
      If Not blnU Then '// upper case letters
         If InStr(1, C_Upper, Mid$(vstrPwd, lngI, 1)) Then
            blnU = True
            lngR = grdResults.AddItem("Contains Upper Case")
            grdResults.RowBackColor(lngR) = LGREEN
         End If
      End If
      
      If Not blnL Then '// lower case letters
         If InStr(1, C_Lower, Mid$(vstrPwd, lngI, 1)) Then
            blnL = True
            lngR = grdResults.AddItem("Contains Lower Case")
            grdResults.RowBackColor(lngR) = LGREEN
         End If
      End If
      
      If Not blnA Then '// spaces
         If Mid$(vstrPwd, lngI, 1) = " " Then
            blnA = True
            lngR = grdResults.AddItem("Contains Spaces")
            grdResults.RowBackColor(lngR) = LGREEN
         End If
      End If
   Next lngI
   
   If Not blnN Then lngR = grdResults.AddItem("Does Not Contain Numbers"): grdResults.RowBackColor(lngR) = LRED
   If Not blnS Then lngR = grdResults.AddItem("Does Not Contain Specials"): grdResults.RowBackColor(lngR) = LRED
   If Not blnU Then lngR = grdResults.AddItem("Does Not Contain Upper Case"): grdResults.RowBackColor(lngR) = LRED
   If Not blnL Then lngR = grdResults.AddItem("Does Not Contain Lower Case"): grdResults.RowBackColor(lngR) = LRED
   If Not blnA Then lngR = grdResults.AddItem("Does Not Contain Spaces"): grdResults.RowBackColor(lngR) = LRED
   
   
   If blnA And Not (blnN Or blnL Or blnU Or blnS) Then '// only [ ]
      sngPoints = 0
      
   Else
      If blnS And Not (blnN Or blnL Or blnU) Then '// only [*$/]
         sngPoints = 1
      ElseIf blnN And Not (blnL Or blnU Or blnS) Then '// only [0-9]
         sngPoints = 0.5
      ElseIf (blnN And blnS) And Not (blnL Or blnU) Then '// only [0-9*$/]
         sngPoints = 1.25
      ElseIf (blnU Or blnL) And Not (blnS Or blnN) Then '// only [a-zA-Z]
         sngPoints = 2
         If Not blnU Or Not blnL Then sngPoints = sngPoints - 0.5
      
      Else
         If (blnL And blnN) And Not (blnU Or blnS) Then '// [a-z0-9]
            sngPoints = 3
         ElseIf (blnU And blnN) And Not (blnL Or blnS) Then '// [A-Z0-9]
            sngPoints = 3
         ElseIf (blnL And blnU And blnN) And Not (blnS) Then '// [a-zA-Z0-9]
            sngPoints = 4
         ElseIf blnN And blnL And blnU And blnS Then '// [a-zA-Z0-9*$/]
            sngPoints = 5
         ElseIf (blnN And blnS) And (blnU Or blnL) Then '// [0-9*$/] and ([a-z] or [A-Z])
            sngPoints = 3
         End If
      End If
      If blnA Then sngPoints = sngPoints + 0.5
   End If
   
   With grdResults
      lngR = .AddItem
      .RowBackColor(lngR) = vb3DDKShadow '.GridColor
      .RowHeight(lngR) = 4
      .RowLocked(lngR) = True
   End With
   
   grdResults.AddItem "Composition Score" & vbTab & sngPoints
   
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
   grdResults.AddItem "Entropy Score" & vbTab & (sngEntropy * -1)
   Erase aryChar
   
   
   '// score Length
   '// passwords with a length = 8 is neutral
   '// > 8 add score; 8 < deduct score
   sngPoints = ((lngPwd - 8) * 0.05)
   If lngPwd <= 6 Then sngPoints = sngPoints - 2
   grdResults.AddItem "Length Score (length of 8 = 0)" & vbTab & sngPoints
   
  
   '// deductions for character repetitions (abcabc, aaaaa, 121212, etc.)
   blnN = False
   sngPoints = 0
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
   grdResults.AddItem "Deductions for character repetitions" & vbTab & sngPoints
   
   '// the following checks need to have a common case
   vstrPwd = LCase$(vstrPwd)
   
   '// deductions for use of Common passwords
   sngPoints = 0
   varComm = Array("!@#", "@#$", "#$%", "$%^", "%^&", "^&*", "&*(", "*()", "000", "007", "246", "249", "1022", "10sne1", "111", "1212", "1225", "123123", "1234", "123abc", "123go", "1313", "13579", "14430", "1701", "thx1138", "1928", "1951", "1a2b3c", "1p2o3i", "1q2w3e", "1qw23e", "1sanjose", "2112", "21122112", "222", "welcome", "369", "444", "4runner", "5252", "54321", "555", "5683", "654321", "666", "6969", "777", "80486", "8675309", "888", "90210", "911", "92072", "999", "a12345", "a1b2c3", "a1b2c3d4", "aaa", "admin", "aaron", "abby", "abc", "asdf", "charlie", "hammer", "happy", "ib6ub9", "icecream", "kermit", "pass", "password", "qwerty", "letmein", "dragon", "master", "monkey", "mustang", "myspace", "money", "blink", "god", "sex", "love", "soccer", "jordan", "football", "baseball", "princess", "shadow", "slipknot", "liverpool", "link182", "super", "xanadu", "xavier", "xcountry", "xfiles", "xxx", "yaco")
   For lngI = 0 To UBound(varComm)
      If InStr(1, vstrPwd, varComm(lngI)) Then
         sngPoints = sngPoints - 0.5
         Exit For
      End If
   Next lngI
   Set varComm = Nothing
   grdResults.AddItem "Deductions for use of common passwords" & vbTab & sngPoints
   
   '// deductions for use of Common Words
   '// This helps to compensate for words or phrases that might appear
   '// secure but are commonly know and therefore less secure.
   '// compaired against a list of 2063 words.
   sngPoints = 0
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
      grdResults.AddItem "Deductions for use of common words" & vbTab & sngPoints
      
   Else
      grdResults.AddItem "Check Failed; Missing file (CommonWords.mdb)"
   End If


   '// Show results
   grdResults.Redraw = True
   grdResults.ColWidth(0) = grdResults.VisibleWidth - 800
   
   lngI = grdResults.TotalsCol(1) * 10
   Select Case lngI
   Case Is < 20
      lblQual.Caption = C_Prefix & "Weak"
      fraQuality.FloodColor = &H6969FF

   Case 20 To 39
      lblQual.Caption = C_Prefix & "Below Average"
      fraQuality.FloodColor = &H69A6FF

   Case 40 To 59
      lblQual.Caption = C_Prefix & "Average"
      fraQuality.FloodColor = &H69FFD9

   Case 60 To 79
      lblQual.Caption = C_Prefix & "Above Average"
      fraQuality.FloodColor = &H69FFA2

   Case 80 To 94
      lblQual.Caption = C_Prefix & "Strong"
      fraQuality.FloodColor = &H75FF69

   Case Else
      lblQual.Caption = C_Prefix & "Best"
      fraQuality.FloodColor = &H60D257
   End Select
   
   fraQuality.FloodPercent = lngI
   Screen.MousePointer = vbDefault
     
End Function

Private Sub txtKey_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then Call cmdLoad_Click

End Sub
