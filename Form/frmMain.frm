VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00D1B499&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin GenPwd.Frame3D Frame3D1 
      Height          =   3705
      Left            =   30
      Top             =   60
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   6535
      BorderType      =   9
      BevelWidth      =   5
      BevelInner      =   3
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   10
      FillColor       =   16767931
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   1
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   3705
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":302A
      Picture         =   "frmMain.frx":3046
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
      Begin GenPwd.CandyButton cmdCheckMS 
         Height          =   360
         Left            =   1995
         TabIndex        =   19
         Top             =   3165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Caption         =   "Test @ Microsoft"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         PictureAlignment=   0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   -1  'True
         ColorScheme     =   0
         CornerRadius    =   7
         UserCornerRadius=   -1
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin GenPwd.Frame3D fraQuality 
         Height          =   285
         Left            =   1995
         Top             =   2745
         Width           =   5430
         _ExtentX        =   9578
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
         MouseIcon       =   "frmMain.frx":3062
         Picture         =   "frmMain.frx":307E
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
      Begin GenPwd.Frame3D Frame3D2 
         Height          =   3555
         Left            =   7530
         Top             =   90
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6271
         BorderType      =   7
         BevelWidth      =   3
         BevelInner      =   0
         Caption3D       =   0
         CaptionAlignment=   0
         CaptionLocation =   0
         BackColor       =   -2147483633
         CornerDiameter  =   7
         FillColor       =   16761247
         FillStyle       =   1
         DrawStyle       =   0
         DrawWidth       =   1
         FloodPercent    =   0
         FloodShowPct    =   0   'False
         FloodType       =   0
         FloodColor      =   16761247
         FillGradient    =   0
         Collapsible     =   0   'False
         ChevronColor    =   -2147483630
         Collapse        =   0   'False
         FullHeight      =   3555
         ChevronType     =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":309A
         Picture         =   "frmMain.frx":30B6
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
         Caption         =   "Options"
         UseMnemonic     =   0   'False
         Begin GenPwd.CandyButton cmdPassPhrase 
            Height          =   330
            Left            =   105
            TabIndex        =   21
            Top             =   2250
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   582
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
            Caption         =   "Generate Pass Phrase"
            PicHighLight    =   0   'False
            CaptionHighLight=   0   'False
            CaptionHighLightColor=   0
            ForeColor       =   16777215
            PictureAlignment=   0
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
            UserCornerRadius=   5
            DisabledPicMode =   0
            UseGREY         =   0   'False
            UseMaskColor    =   -1  'True
            MaskColor       =   -2147483633
            ButtonBehaviour =   0
            ShowFocus       =   0   'False
         End
         Begin VB.CheckBox chkLastWithNumber 
            Caption         =   "End With Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   195
            TabIndex        =   15
            Top             =   1965
            Width           =   1935
         End
         Begin VB.CheckBox chkStartWithNumber 
            Caption         =   "Start With Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   195
            TabIndex        =   14
            Top             =   1665
            Width           =   1935
         End
         Begin VB.ComboBox cboLength 
            Height          =   360
            ItemData        =   "frmMain.frx":30D2
            Left            =   180
            List            =   "frmMain.frx":30D4
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   315
            Width           =   765
         End
         Begin VB.CheckBox chkSpecial 
            Caption         =   "Special"
            Height          =   255
            Left            =   195
            TabIndex        =   11
            Top             =   1350
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkNumbers 
            Caption         =   "Numbers"
            Height          =   255
            Left            =   195
            TabIndex        =   10
            Top             =   1065
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkUpper 
            Caption         =   "Uppercase"
            Height          =   255
            Left            =   195
            TabIndex        =   9
            Top             =   780
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin GenPwd.CandyButton cmdSave 
            Height          =   375
            Left            =   390
            TabIndex        =   17
            Top             =   2715
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
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
            Caption         =   "Save"
            PicHighLight    =   0   'False
            CaptionHighLight=   0   'False
            CaptionHighLightColor=   0
            Picture         =   "frmMain.frx":30D6
            PictureAlignment=   2
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
         Begin GenPwd.CandyButton cmdLoad 
            Height          =   375
            Left            =   405
            TabIndex        =   18
            Top             =   3105
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
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
            Caption         =   "Load"
            PicHighLight    =   0   'False
            CaptionHighLight=   0   'False
            CaptionHighLightColor=   0
            Picture         =   "frmMain.frx":3470
            PictureAlignment=   2
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
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            X1              =   180
            X2              =   2075
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " Min. Length "
            Height          =   240
            Index           =   4
            Left            =   945
            TabIndex        =   13
            Top             =   360
            Width           =   1110
         End
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1620
         Width           =   5280
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
         Index           =   2
         Left            =   1995
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1170
         Width           =   5460
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
         Index           =   1
         Left            =   1995
         MaxLength       =   255
         TabIndex        =   2
         Top             =   750
         Width           =   5460
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
         Index           =   0
         Left            =   1995
         MaxLength       =   255
         TabIndex        =   0
         Top             =   300
         Width           =   5460
      End
      Begin GenPwd.CandyButton cmdStrongPwd 
         Height          =   360
         Left            =   3772
         TabIndex        =   20
         Top             =   3165
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Caption         =   "Strong Password?"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         PictureAlignment=   0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   -1  'True
         ColorScheme     =   0
         CornerRadius    =   7
         UserCornerRadius=   -1
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin GenPwd.CandyButton cmdQualityTest 
         Height          =   360
         Left            =   5730
         TabIndex        =   23
         Top             =   3165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Caption         =   "Quality Test"
         PicHighLight    =   -1  'True
         CaptionHighLight=   0   'False
         PictureAlignment=   0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   -1  'True
         ColorScheme     =   0
         CornerRadius    =   7
         UserCornerRadius=   -1
         DisabledPicMode =   0
         UseGREY         =   0   'False
         UseMaskColor    =   -1  'True
         MaskColor       =   -2147483633
         ButtonBehaviour =   0
         ShowFocus       =   0   'False
      End
      Begin VB.Label lblRemberAs 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1230
         TabIndex        =   22
         Top             =   1935
         Width           =   6255
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   165
         Picture         =   "frmMain.frx":380A
         Top             =   2220
         Width           =   1725
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
         Left            =   4200
         TabIndex        =   8
         Top             =   2355
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Password: "
         Height          =   240
         Index           =   3
         Left            =   930
         TabIndex        =   7
         Top             =   1665
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Master Password: "
         Height          =   240
         Index           =   2
         Left            =   285
         TabIndex        =   5
         Top             =   1215
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Your Name: "
         Height          =   240
         Index           =   1
         Left            =   810
         TabIndex        =   3
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Password For: "
         Height          =   240
         Index           =   0
         Left            =   585
         TabIndex        =   1
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "  "
         Height          =   315
         Index           =   5
         Left            =   1995
         TabIndex        =   16
         Top             =   1620
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hWnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Private cPlay As clsPlaySound
Private mblnLoadData As Boolean

Private Sub cmdQualityTest_Click()
   
   frmUserCheck.Show , Me

End Sub

Private Sub cboLength_Click()

   Call CreatePassword

End Sub

Private Sub chkLastWithNumber_Click()

   Call CreatePassword

End Sub

Private Sub chkNumbers_Click()

   Call CreatePassword

End Sub

Private Sub chkSpecial_Click()

   Call CreatePassword

End Sub

Private Sub chkStartWithNumber_Click()

   Call CreatePassword

End Sub

Private Sub chkUpper_Click()

   Call CreatePassword

End Sub

Private Sub cmdCheckMS_Click()

   ShellExecute Me.hWnd, "open", "https://www.microsoft.com/protect/yourself/password/checker.mspx", vbNullString, "C:\", 5

End Sub

Private Sub cmdLoad_Click()

   mblnLoadData = True
   frmLoad.Show vbModal, Me
   mblnLoadData = False
   Call CreatePassword
   txtKey(2).SetFocus

End Sub

Private Sub cmdPassPhrase_Click()

  Dim lngI     As Long
  Dim strPWD   As String
  Dim strPWDO  As String
  Const C_Prefix As String = "Password Quality is "

   strPWD = GetCreatedPassPhrase(Trim$(txtKey(0).Text) & _
                                 Trim$(txtKey(1).Text) & _
                                 Trim$(txtKey(2).Text), _
                                 cboLength.ListIndex + C_StartLength, _
                                 chkUpper.Value, _
                                 chkNumbers.Value, _
                                 chkSpecial.Value, _
                                 strPWDO)

   '// return password
   txtPassword.Text = strPWD
   lblRemberAs.Caption = "Remeber:" & strPWDO

   '// Get password quality
   Call GetPasswordQuality(strPWD, True)

End Sub

Private Sub cmdSave_Click()

  Dim MyDB  As ADODB.Connection

   If LenB(Dir$(gstrDB)) = 0 Then
      Call CreateMDB(gstrDB)
   End If

   Call OpenDB(MyDB, gstrDB)

   If glngID Then
      If MsgBox("Do you want to overwrite the last record you loaded?", vbQuestion + vbYesNo) = vbYes Then

         MyDB.Execute "UPDATE PwdData SET PwdData.pFor = '" & txtKey(0).Text & "', PwdData.pName = '" & txtKey(1).Text & "'," & _
            " PwdData.pLength = " & cboLength.Text & ", PwdData.pUppercase = " & CStr(chkUpper.Value) & ", PwdData.pNumbers = " & _
            CStr(chkNumbers.Value) & ", PwdData.pSpecial = " & CStr(chkSpecial.Value) & ", PwdData.pFirstNumber = " & _
            CStr(chkStartWithNumber.Value) & ", PwdData.pLastNumber = " & CStr(chkLastWithNumber.Value) & _
            " WHERE (((PwdData.pID)=" & CStr(glngID) & "));"

         GoTo Exit_Proc
      End If

   End If

   glngID = 0

   MyDB.Execute "INSERT INTO PwdData (pFor, pName, pLength, pUppercase, pNumbers, pSpecial, pFirstNumber, pLastNumber) VALUES ('" _
      & txtKey(0).Text & "', '" & txtKey(1).Text & "', " & cboLength.Text & ", " & CStr(chkUpper.Value) & ", " & _
      CStr(chkNumbers.Value) & ", " & CStr(chkSpecial.Value) & ", " & CStr(chkStartWithNumber.Value) & ", " & _
      CStr(chkLastWithNumber.Value) & ")"

Exit_Proc:
   MyDB.Close
   cPlay.PlaySoundResource 101
   cmdLoad.Enabled = (LenB(Dir$(gstrDB)) > 0)

End Sub

Private Sub cmdStrongPwd_Click()

   ShellExecute Me.hWnd, "open", "https://www.microsoft.com/protect/yourself/password/create.mspx", vbNullString, "C:\", 5

End Sub

Private Sub CreatePassword()

  Dim lngI        As Long
  Dim strPWD      As String

   If Not mblnLoadData Then
      strPWD = GetCreatedPassword(Trim$(txtKey(0).Text) & _
                                  Trim$(txtKey(1).Text) & _
                                  Trim$(txtKey(2).Text), _
                                  cboLength.ListIndex + C_StartLength, _
                                  chkUpper.Value, _
                                  chkNumbers.Value, _
                                  chkSpecial.Value, _
                                  chkStartWithNumber.Value, _
                                  chkLastWithNumber.Value)
   
      '// return password
      txtPassword.Text = strPWD
      lblRemberAs.Caption = vbNullString
   
      '// Get password quality
      Call GetPasswordQuality(strPWD)
   End If
   
End Sub

Private Sub Form_Load()

  Dim lngI As Long

   Set cPlay = New clsPlaySound
   mblnLoadData = True

   '// set password length selection

   For lngI = C_StartLength To 20
      cboLength.AddItem lngI
   Next lngI

   '// set default length to 10
   cboLength.ListIndex = 4

   mblnLoadData = False
   '// set user name
   txtKey(1).Text = Environ$("USERNAME")

   cmdLoad.Enabled = (LenB(Dir$(gstrDB)) > 0)
   cmdPassPhrase.Enabled = (LenB(Dir$(gstrWordsDB)) > 0)
   

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set cPlay = Nothing

   Call EndApp(Me)
   Set frmMain = Nothing

End Sub

Public Sub GetPasswordQuality(ByVal strPWD As String, Optional ByVal vblnCheckDic As Boolean)

  Dim lngI        As Long
  Const C_Prefix  As String = "Password Quality is "
  
   '// Get password quality
   lngI = PWDQuality(strPWD, vblnCheckDic)

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

End Sub

Private Sub Image1_Click()

   frmAbout.Show , Me

End Sub

Private Sub txtKey_Change(Index As Integer)

   Call CreatePassword

End Sub

