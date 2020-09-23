VERSION 5.00
Begin VB.Form frmLoad 
   Caption         =   "Load"
   ClientHeight    =   3750
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin GenPwd.LynxGrid grdList 
      Align           =   1  'Align Top
      Height          =   3465
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6112
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
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      BorderStyle     =   0
      FocusRectColor  =   9895934
      Appearance      =   0
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin GenPwd.Frame3D fraToolBar 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   714
      BorderType      =   0
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   0
      CornerDiameter  =   7
      FillColor       =   3342336
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   405
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmLoad.frx":0000
      Picture         =   "frmLoad.frx":001C
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
      Begin GenPwd.CandyButton cmdLoad 
         Height          =   405
         Left            =   105
         TabIndex        =   1
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
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
         Caption         =   "Load"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         CaptionHighLightColor=   0
         ForeColor       =   16777215
         Picture         =   "frmLoad.frx":0038
         PictureAlignment=   2
         Style           =   8
         Checked         =   0   'False
         ColorButtonHover=   3342336
         ColorButtonUp   =   0
         ColorButtonDown =   6684672
         BorderBrightness=   0
         ColorBright     =   16768256
         DisplayHand     =   -1  'True
         ColorScheme     =   7
         CornerRadius    =   12
         UserCornerRadius=   3
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin GenPwd.CandyButton cmdDelete 
         Height          =   405
         Left            =   1605
         TabIndex        =   2
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
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
         Caption         =   "Delete"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         CaptionHighLightColor=   0
         ForeColor       =   16777215
         Picture         =   "frmLoad.frx":03D2
         PictureAlignment=   2
         Style           =   8
         Checked         =   0   'False
         ColorButtonHover=   51
         ColorButtonUp   =   0
         ColorButtonDown =   102
         BorderBrightness=   0
         ColorBright     =   255
         DisplayHand     =   -1  'True
         ColorScheme     =   8
         CornerRadius    =   12
         UserCornerRadius=   3
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   16777215
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin GenPwd.CandyButton cmdPrint 
         Height          =   405
         Left            =   3105
         TabIndex        =   3
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
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
         Caption         =   "Print All"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         CaptionHighLightColor=   0
         ForeColor       =   16777215
         Picture         =   "frmLoad.frx":076C
         PictureAlignment=   2
         Style           =   8
         Checked         =   0   'False
         ColorButtonHover=   3342336
         ColorButtonUp   =   0
         ColorButtonDown =   6684672
         BorderBrightness=   0
         ColorBright     =   16768256
         DisplayHand     =   -1  'True
         ColorScheme     =   7
         CornerRadius    =   12
         UserCornerRadius=   3
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin GenPwd.CandyButton cmdExport 
         Height          =   405
         Left            =   4605
         TabIndex        =   4
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
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
         Caption         =   "Export Grid"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         CaptionHighLightColor=   0
         ForeColor       =   16777215
         Picture         =   "frmLoad.frx":0B06
         PictureAlignment=   2
         Style           =   8
         Checked         =   0   'False
         ColorButtonHover=   3342336
         ColorButtonUp   =   0
         ColorButtonDown =   6684672
         BorderBrightness=   0
         ColorBright     =   16768256
         DisplayHand     =   -1  'True
         ColorScheme     =   7
         CornerRadius    =   12
         UserCornerRadius=   3
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()

  Dim MyDB As ADODB.Connection

   If MsgBox("Are you sure you want to delete " & grdList.CellText(, 0) & "?", vbQuestion + vbOKCancel) = vbOK Then

      Call OpenDB(MyDB, gstrDB)
      MyDB.Execute "DELETE PwdData.* From PwdData WHERE (((PwdData.pID)=" & CStr(grdList.RowData) & "));"
      MyDB.Close
      grdList.RemoveRow

      cmdLoad.Enabled = (grdList.Rows > 0)
      cmdDelete.Enabled = (grdList.Rows > 0)
   End If

End Sub

Private Sub cmdExport_Click()

   grdList.ExportGrid App.Title

End Sub

Private Sub cmdLoad_Click()

   With frmMain
      glngID = grdList.RowData
      .txtKey(0).Text = grdList.CellText(, 0)
      .txtKey(1).Text = grdList.CellText(, 1)
      .cboLength.ListIndex = grdList.CellValue(, 2) - C_StartLength
      .chkUpper.Value = Abs(grdList.CellValue(, 3))
      .chkNumbers.Value = Abs(grdList.CellValue(, 4))
      .chkSpecial.Value = Abs(grdList.CellValue(, 5))
      .chkStartWithNumber.Value = Abs(grdList.CellValue(, 6))
      .chkLastWithNumber.Value = Abs(grdList.CellValue(, 7))
      ''' .txtKey(2).SetFocus
   End With

   Unload Me

End Sub

Private Sub cmdPrint_Click()

  Dim MyDB     As ADODB.Connection
  Dim MySet    As ADODB.Recordset
  Dim strMPW   As String
  Dim strPWD   As String
  Dim lngI     As Long

   strMPW = Trim$(frmMain.txtKey(2).Text)

   If LenB(strMPW) = 0 Then
      If MsgBox("There is no master password entered.  Do you want to print anyway?", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If

   Call OpenDB(MyDB, gstrDB)
   Call OpenRS(MySet, "Select * From PwdData Order By pName;", MyDB)

   If ADORecordCount(MySet) Then
      Printer.FontSize = 11
      Printer.FontName = "Tahoma"
      Printer.Print "";
      Printer.Print "Printed: " & Now
      Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
      Printer.Print " "
      Printer.FontSize = 12
      Printer.FontName = "Courier New" '"Tahoma"
      Printer.Print " "

      With MySet

         Do
            strPWD = GetCreatedPassword(.Fields("pFor") & .Fields("pName") & strMPW, .Fields("pLength"), .Fields("pUppercase"), _
               .Fields("pNumbers"), .Fields("pSpecial"), .Fields("pFirstNumber"), .Fields("pLastNumber"))

            lngI = 55 - Len(.Fields("pFor"))
            If lngI < 2 Then lngI = 2

            Printer.Print " " & .Fields("pFor") & String(lngI, ".") & "  " & strPWD
            Printer.FontSize = 6
            Printer.Print " "
            Printer.FontSize = 12
            .MoveNext
         Loop Until .EOF

         Printer.EndDoc

      End With
   End If

   MySet.Close
   MyDB.Close

End Sub

Private Sub Form_Load()

  Dim MyDB  As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim lngR  As Long

   With grdList
      lngR = (.VisibleWidth - 3000) \ 2
      .AddColumn "Password For", lngR
      .AddColumn "Your Name", lngR
      .AddColumn "Length", 500, lgAlignRightCenter, lgNumeric
      .AddColumn "Uppercase", 500, lgAlignCenterCenter, lgBoolean
      .AddColumn "Numbers", 500, lgAlignCenterCenter, lgBoolean
      .AddColumn "Special", 500, lgAlignCenterCenter, lgBoolean
      .AddColumn "Number First", 500, lgAlignCenterCenter, lgBoolean
      .AddColumn "Number Last", 500, lgAlignCenterCenter, lgBoolean
      .Height = Me.ScaleHeight - fraToolBar.Height
   End With

   Call OpenDB(MyDB, gstrDB)
   Call OpenRS(MySet, "Select * From PwdData Order By pName;", MyDB)

   If ADORecordCount(MySet) Then

      With MySet

         Do
            lngR = grdList.AddItem(.Fields("pFor") & vbTab & .Fields("pName") & vbTab & .Fields("pLength") & vbTab & _
               .Fields("pUppercase") & vbTab & .Fields("pNumbers") & vbTab & .Fields("pSpecial") & vbTab & _
               .Fields("pFirstNumber") & vbTab & .Fields("pLastNumber"))

            grdList.RowData(lngR) = .Fields("pID")
            .MoveNext
         Loop Until .EOF

      End With
      'grdList.ColWidthAutoSize
   End If

   grdList.Redraw = True

   MySet.Close
   MyDB.Close

   If grdList.Rows > 0 Then grdList.RowColSet 0, 0

   cmdExport.Enabled = (grdList.Rows > 0)
   cmdPrint.Enabled = (grdList.Rows > 0)
   cmdLoad.Enabled = (grdList.Rows > 0)
   cmdDelete.Enabled = (grdList.Rows > 0)
   glngID = 0

End Sub

Private Sub Form_Resize()

   grdList.Redraw = False
   grdList.Height = Me.ScaleHeight - fraToolBar.Height - 50
   grdList.Redraw = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmLoad = Nothing

End Sub

Private Sub grdList_DblClick()

   If grdList.Rows Then Call cmdLoad_Click

End Sub

