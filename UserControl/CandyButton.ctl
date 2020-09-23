VERSION 5.00
Begin VB.UserControl CandyButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ClipBehavior    =   0  'None
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   ToolboxBitmap   =   "CandyButton.ctx":0000
End
Attribute VB_Name = "CandyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// By: Mario Villanueva
'// Submitted on: 2/18/2007 7:16:52 AM
'// http://www.planet-source-code.com/vb/scripts/showcode.asp?lngWId=1&txtCodeId=64969

'-------------------------------------------------------------------------------------------------
' Modifications by: Morgan Haueisen
' Date: 1/27/2009
'  * Removed XP style buttons
'  * Fixed several PropertyChanged where it referred to variable name and not property name in Read/Write Properties
'  * Fixed Show Hand Icon
'  * Fixed spelling of HighLite to HighLight
'  * Changed how text is drawn
'  * Changed how Picture is drawn (added MaskColor, Use Mask color, Use Grey colors)
'  * Changed IconHighLite
'  * Changed default Font to "Tahoma"
'  * Added User defined Corner Radius which will override the default setting
'  * Added hWnd
'  * Added Option Buttons style
'  * Added LeftEdge and RightEdge to Alignments
'  * Added Default/Cancel
'  * Added Key Events
'  * Added optional Focus Rectangle
'  * Added GlowButton button and 4 color schemes
'  * Changed sub-classing
'  * Various code changes to improve readability and speed

'-------------------------------------------------------------------------------------------------
' Sub-Classing code
' Author: Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
' v1.7 Changed zAddressOf, removed zProbe, and added Subs GetMem1 and GetMem4............20080422

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                           'When to callback
   MSG_BEFORE = 1                               'Callback before the original WndProc
   MSG_AFTER = 2                                'Callback after the original WndProc
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER   'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1        'All messages callback
Private Const MSG_ENTRIES   As Long = 32        'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38      'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4        'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1         'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2         'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9         'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11        'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12        'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13        'Thunk data index of the User-defined callback parameter data index
Private z_ScMem             As Long             'Thunk base address
Private z_Sc(64)            As Long             'Thunk machine-code initialised here
Private z_Funk              As Collection       'hWnd/thunk-address collection

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Function CallWindowProcA Lib "user32" ( _
      ByVal lpPrevWndFunc As Long, _
      ByVal hWnd As Long, _
      ByVal Msg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" ( _
      ByVal lpAddress As Long, _
      ByVal dwSize As Long, _
      ByVal flAllocationType As Long, _
      ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
      ByVal lpAddress As Long, _
      ByVal dwSize As Long, _
      ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Const WM_MOUSEMOVE    As Long = &H200
Private Const WM_MOUSELEAVE   As Long = &H2A3
Private Const WM_MOVING       As Long = &H216
Private Const WM_SIZING       As Long = &H214
Private Const WM_EXITSIZEMOVE As Long = &H232
'''Private Const WM_PAINT        As Long = &HF

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize            As Long
   dwFlags           As TRACKMOUSEEVENT_FLAGS
   hwndTrack         As Long
   dwHoverTime       As Long
End Type

Private bTrack       As Boolean
Private bTrackUser32 As Boolean
Private IsHover      As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" _
      Alias "_TrackMouseEvent" ( _
      ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long


'// Candy Button declarations----------------------------------------------------------------------------
Private Type typCrystalParam
   Ref_Intensity     As Long
   Ref_Left          As Long
   Ref_Top           As Long
   Ref_Radius        As Long
   Ref_Height        As Long
   Ref_Width         As Long
   RadialGXPercent   As Long
   RadialGYPercent   As Long
   RadialGOffsetX    As Long
   RadialGOffsetY    As Long
   RadialGIntensity  As Long
End Type

Private Type typBITMAPINFOHEADER
   biSize            As Long
   biWidth           As Long
   biHeight          As Long
   biPlanes          As Integer
   biBitCount        As Integer
   biCompression     As Long
   biSizeImage       As Long
   biXPelsPerMeter   As Long
   biYPelsPerMeter   As Long
   biClrUsed         As Long
   biClrImportant    As Long
End Type

Private Type typRGBTRIPLE
   rgbBlue  As Byte
   rgbGreen As Byte
   rgbRed   As Byte
End Type

Private Type typBITMAP
   bmType         As Long
   bmWidth        As Long
   bmHeight       As Long
   bmWidthBytes   As Long
   bmPlanes       As Integer
   bmBitsPixel    As Integer
   bmBits         As Long
End Type

Private Type typBITMAPINFO
   bmiHeader As typBITMAPINFOHEADER
   bmiColors As typRGBTRIPLE
End Type

Private Const BI_RGB                As Long = 0&
Private Const DIB_RGB_COLORS        As Long = 0&
Private Const DST_TEXT              As Long = &H1
Private Const DST_PREFIXTEXT        As Long = &H2
Private Const DST_COMPLEX           As Long = &H0
Private Const DST_ICON              As Long = &H3
Private Const DST_typBITMAP         As Long = &H4
Private Const DSS_NORMAL            As Long = &H0
Private Const DSS_UNION             As Long = &H10
Private Const DSS_DISABLED          As Long = &H20
Private Const DSS_MONO              As Long = &H80
Private Const DSS_RIGHT             As Long = &H8000
Private Const RGN_XOR               As Long = 3
Private Const MK_LBUTTON            As Long = &H1

Private Const C_DT_CALCRECT         As Long = &H400&
Private Const C_DT_WORDBREAK        As Long = &H10&
Private Const C_DT_CENTER           As Long = &H1& Or C_DT_WORDBREAK Or &H4&

Private Type typPOINTAPI
   X As Long
   Y As Long
End Type

Private Type typRECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32" ( _
      ByVal hDestRgn As Long, _
      ByVal hSrcRgn1 As Long, _
      ByVal hSrcRgn2 As Long, _
      ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long, _
      ByVal X3 As Long, _
      ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal crColor As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As typRECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Ellipse Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function SetRect Lib "user32" ( _
      ByRef lpRect As typRECT, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As typRECT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As typRECT, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
      ByVal lOleColor As Long, _
      ByVal lHPalette As Long, _
      ByRef lColorRef As Long) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
      ByVal hDestDC As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" ( _
      ByVal hdc As Long, _
      ByVal xLeft As Long, _
      ByVal yTop As Long, _
      ByVal hIcon As Long, _
      ByVal cxWidth As Long, _
      ByVal cyWidth As Long, _
      ByVal istepIfAniCur As Long, _
      ByVal hbrFlickerFreeDraw As Long, _
      ByVal diFlags As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" ( _
      ByVal aHDC As Long, _
      ByVal hBitmap As Long, _
      ByVal nStartScan As Long, _
      ByVal nNumScans As Long, _
      ByRef lpBits As Any, _
      ByRef lpbi As typBITMAPINFO, _
      ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal SrcX As Long, _
      ByVal SrcY As Long, _
      ByVal Scan As Long, _
      ByVal NumScans As Long, _
      ByRef Bits As Any, _
      ByRef BitsInfo As typBITMAPINFO, _
      ByVal wUsage As Long) As Long
      
'''Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, ByRef lpRect As typRECT) As Long
'''Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As typRECT) As Long
Private Declare Function RoundRect Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal Left As Long, _
      ByVal Top As Long, _
      ByVal Right As Long, _
      ByVal Bottom As Long, _
      ByVal EllipseWidth As Long, _
      ByVal EllipseHeight As Long) As Long
Private Declare Function OffsetRect Lib "user32" (ByRef lpRect As typRECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (ByRef lpDestRect As typRECT, ByRef lpSourceRect As typRECT) As Long
Private Declare Function DrawText Lib "user32" _
      Alias "DrawTextA" ( _
      ByVal hdc As Long, _
      ByVal lpStr As String, _
      ByVal nCount As Long, _
      ByRef lpRect As typRECT, _
      ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'// Hand Cursor ---------------------------------------------
Private Type typPICTDESC
    cbSize     As Long
    pictType   As Long
    hIcon      As Long
    hPal       As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      ByRef lpPictDesc As typPICTDESC, _
      ByRef riid As Any, _
      ByVal fOwn As Long, _
      ByRef ipic As IPicture) As Long
Private Declare Function LoadCursor Lib "user32.dll" _
      Alias "LoadCursorA" ( _
      ByVal hInstance As Long, _
      ByVal lpCursorName As Long) As Long

'// Public Enum ------------------------------------------------------------------------------------
Public Enum enuCandy_Alignment
   PIC_TOP = 0
   PIC_BOTTOM = 1
   PIC_LEFT = 2
   PIC_RIGHT = 3
   PIC_LeftEdge = 4
   PIC_RIGHTEdge = 5
End Enum

Public Enum enuCandy_Style
   Crystal = 2
   MAC = 3
   MAC_Variation = 4
   WMP = 5
   Plastic = 6
   Iceblock = 7
   GlowButton = 8
End Enum
#If False Then
   Private Crystal, MAC, MAC_Variation, WMP, Plastic, Iceblock, GlowButton
#End If

Public Enum enuCandy_ColorScheme
   Custom = 0
   Aqua = 1
   WMP10 = 2
   DeepBlue = 3
   DeepRed = 4
   DeepGreen = 5
   DeepYellow = 6
   BlackBlue = 7
   BlackRed = 8
   BlackGreen = 9
   BlackYellow = 10
End Enum

Public Enum enuCandy_State
   eNormal = 0
   ePressed = 1
   eHover = 3
End Enum

Public Enum enuCandy_DisablePicMode
   [Grayed] = 0
   [Blended] = 1
End Enum

Public Enum enuCandy_ButtonBehaviour
   [Standard Button] = 0
   [Check Box] = -1
   [Option Button] = 1
End Enum

'// private variables ----------------------------------------------------------------------
Private mlngButtonRegion         As Long
Private mudtCrystalParam         As typCrystalParam

Private mudtRC                   As typRECT
Private mudtRC2                  As typRECT
'''Private mudtRC3               As typRECT '// Focus Rect
Private mudtPicPt                As typPOINTAPI
Private mudtPicSz                As typPOINTAPI '// picture Position & Size

Private mudtDisabledPicMode      As enuCandy_DisablePicMode
Private mudtPictureAlignment     As enuCandy_Alignment
Private mblnUseMask              As Boolean
Private mblnUseGrey              As Boolean     '// use only grey colors for pictures
Private mlngMaskColor            As Long        '// mask color
Private mpicButtonPic            As StdPicture
Private mblnPicHighLight         As Boolean

Private mudtStyle                As enuCandy_Style
Private mudtColorScheme          As enuCandy_ColorScheme
Private mudtButtonBehaviour      As enuCandy_ButtonBehaviour
Private mblnDisplayHand          As Boolean
Private mlngHandCursor           As Long
Private mlngCornerRadius         As Long
Private mlngUserCornerRadius     As Long
Private mlngBorderBrightness     As Long

Private mblnIsEnabled            As Boolean
Private mblnIsChecked            As Boolean
Private mblnIsOver               As Boolean
Private mblnHasFocus             As Boolean
Private mblnShowFocus            As Boolean

Private mclrButtonHover          As OLE_COLOR
Private mclrButtonUp             As OLE_COLOR
Private mclrButtonDown           As OLE_COLOR
Private mclrButtonBright         As OLE_COLOR

Private mstrCaption              As String
Private mclrCapForecolor         As OLE_COLOR
Private mclrCapHighLightColor    As OLE_COLOR
Private mblnCapHighLight         As Boolean

'// public events -------------------------------------------------------------------------
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Status(ByVal sStatus As String)

Private Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Percentage As Long) As Long

  Dim lngR(2) As Long
  Dim lngG(2) As Long
  Dim lngB(2) As Long

   Percentage = SetBound(Percentage, 0, 100)

   GetRGB lngR(0), lngG(0), lngB(0), Color1
   GetRGB lngR(1), lngG(1), lngB(1), Color2

   lngR(2) = lngR(0) + (lngR(1) - lngR(0)) * Percentage \ 100
   lngG(2) = lngG(0) + (lngG(1) - lngG(0)) * Percentage \ 100
   lngB(2) = lngB(0) + (lngB(1) - lngB(0)) * Percentage \ 100

   BlendColors = RGB(lngR(2), lngG(2), lngB(2))

End Function

Public Property Get BorderBrightness() As Long

   BorderBrightness = mlngBorderBrightness

End Property

Public Property Let BorderBrightness(ByVal vNewValue As Long)

   mlngBorderBrightness = SetBound(vNewValue, -100, 100)
   PropertyChanged "BorderBrightness"
   Call DrawButton

End Property

Public Property Get Caption() As String

   Caption = mstrCaption

End Property

Public Property Let Caption(ByVal vNewValue As String)

   mstrCaption = vNewValue
   PropertyChanged "Caption"
   Call CalcTextRects
   Call DrawButton

End Property

Public Property Get CaptionHighLight() As Boolean

   CaptionHighLight = mblnCapHighLight

End Property

Public Property Let CaptionHighLight(ByVal vNewValue As Boolean)

   mblnCapHighLight = vNewValue
   PropertyChanged "CaptionHighLite"

End Property

Public Property Get CaptionHighLightColor() As OLE_COLOR

   CaptionHighLightColor = mclrCapHighLightColor

End Property

Public Property Let CaptionHighLightColor(ByVal vNewValue As OLE_COLOR)

   mclrCapHighLightColor = vNewValue
   PropertyChanged "CaptionHighLiteColor"

End Property

Public Property Get Checked() As Boolean

   '// Here for backward compatibly (same as Value property)
   Checked = mblnIsChecked

End Property

Public Property Let Checked(ByVal vNewValue As Boolean)

   '// Here for backward compatibly (same as Value property)
   Value = vNewValue

End Property

Public Property Get ColorBright() As OLE_COLOR

   ColorBright = mclrButtonBright

End Property

Public Property Let ColorBright(ByVal vNewValue As OLE_COLOR)

   mclrButtonBright = ConvertFromSystemColor(vNewValue)
   
   If Not mudtColorScheme = Custom Then
      mudtColorScheme = Custom
      PropertyChanged "ColorScheme"
   End If
   
   PropertyChanged "ColorBright"
   Call DrawButton

End Property

Public Property Get ColorButtonDown() As OLE_COLOR

   ColorButtonDown = mclrButtonDown

End Property

Public Property Let ColorButtonDown(ByVal vNewValue As OLE_COLOR)

   mclrButtonDown = ConvertFromSystemColor(vNewValue)
   
   If Not mudtColorScheme = Custom Then
      mudtColorScheme = Custom
      PropertyChanged "ColorScheme"
   End If
   
   PropertyChanged "ColorButtonDown"
   Call DrawButton
   
End Property

Public Property Get ColorButtonHover() As OLE_COLOR

   ColorButtonHover = mclrButtonHover

End Property

Public Property Let ColorButtonHover(ByVal vNewValue As OLE_COLOR)

   mclrButtonHover = ConvertFromSystemColor(vNewValue)
   
   If Not mudtColorScheme = Custom Then
      mudtColorScheme = Custom
      PropertyChanged "ColorScheme"
   End If
   
   PropertyChanged "ColorButtonHover"
   Call DrawButton

End Property

Public Property Get ColorButtonUp() As OLE_COLOR

   ColorButtonUp = mclrButtonUp

End Property

Public Property Let ColorButtonUp(ByVal vNewValue As OLE_COLOR)

   mclrButtonUp = ConvertFromSystemColor(vNewValue)
   
   If mudtColorScheme <> Custom Then
      mudtColorScheme = Custom
      PropertyChanged "ColorScheme"
   End If
   
   PropertyChanged "ColorButtonUp"
   Call DrawButton

End Property

Public Property Get ColorScheme() As enuCandy_ColorScheme

   ColorScheme = mudtColorScheme

End Property

Public Property Let ColorScheme(ByVal vNewValue As enuCandy_ColorScheme)
   
   mudtColorScheme = vNewValue
   Call ColorSchemeSet
   PropertyChanged "ColorScheme"
   Call DrawButton

End Property

Private Sub ColorSchemeSet()

   Select Case mudtColorScheme
   Case Aqua
      mclrButtonUp = &HD06720
      mclrButtonHover = &HE99950
      mclrButtonDown = &HA06710
      mclrButtonBright = &HFFEDB0

   Case WMP10
      mclrButtonUp = &HD09060
      mclrButtonHover = &HE06000
      mclrButtonDown = &HA98050
      mclrButtonBright = &HFFFAFA

   Case DeepBlue
      mclrButtonUp = &H800000
      mclrButtonHover = &HA00000
      mclrButtonDown = &HF00000
      mclrButtonBright = &HFF0000

   Case DeepRed
      mclrButtonUp = &H80&
      mclrButtonHover = &HA0&
      mclrButtonDown = &HF0&
      mclrButtonBright = &HFF&

   Case DeepGreen
      mclrButtonUp = &H8000&
      mclrButtonHover = &HA000&
      mclrButtonDown = &HC000&
      mclrButtonBright = &HFF00&

   Case DeepYellow
      mclrButtonUp = &H8080&
      mclrButtonHover = &HA0A0&
      mclrButtonDown = &HC0C0&
      mclrButtonBright = &HFFFF&
      
   Case BlackRed
      mclrButtonUp = &H0
      mclrButtonHover = &H33&
      mclrButtonDown = &H66&
      mclrButtonBright = &HFF&
   
   Case BlackBlue
      mclrButtonUp = &H0
      mclrButtonHover = &H330000
      mclrButtonDown = &H660000
      mclrButtonBright = &HFFDD00
   
   Case BlackGreen
      mclrButtonUp = &H0&
      mclrButtonHover = &H3300&
      mclrButtonDown = &H6600&
      mclrButtonBright = &HFF00&
   
   Case BlackYellow
      mclrButtonUp = &H0&
      mclrButtonHover = &H3333&
      mclrButtonDown = &H4B4C&
      mclrButtonBright = &HFFFF&
   
   End Select

End Sub

Private Function ConvertFromSystemColor(ByVal vColor As Long) As Long

   Call OleTranslateColor(vColor, 0&, ConvertFromSystemColor)

End Function

Public Property Get CornerRadius() As Long

   CornerRadius = mlngUserCornerRadius

End Property

Public Property Let CornerRadius(ByVal vNewValue As Long)

   mlngUserCornerRadius = vNewValue
   Call Init_Style
   Call DrawButton
   PropertyChanged "CornerRadius"

End Property

Private Sub CalcButtonParam()
  
  Dim lngCR       As Long
  Dim lngWidth    As Long
  Dim lngHeight   As Long

   lngHeight = UserControl.ScaleHeight
   lngWidth = UserControl.ScaleWidth
   lngCR = GetCornerRadius
   
   With mudtCrystalParam
   
      Select Case mudtStyle
      Case Plastic
         '// None required
      Case MAC
         .Ref_Intensity = 70
         .Ref_Left = (lngCR \ 3)
         .Ref_Top = 0
         .Ref_Height = 12
         .Ref_Width = lngWidth + 2 * lngCR
         .Ref_Radius = 10
         .RadialGXPercent = 200
         .RadialGYPercent = 100 - (7 * 100 \ lngHeight)
         If .RadialGYPercent > 80 Then .RadialGYPercent = 80
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight
         .RadialGIntensity = 130
         
      Case WMP
         .Ref_Intensity = 40
         .Ref_Left = -lngCR \ 2 - 1
         .Ref_Top = -lngCR
         .Ref_Height = (lngCR) + 1
         .Ref_Width = lngWidth + 2 * lngCR
         .Ref_Radius = lngCR
         .RadialGXPercent = 60
         .RadialGYPercent = 60
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight
         .RadialGIntensity = 130
         
      Case GlowButton
         .Ref_Intensity = 70
         .Ref_Left = -lngCR \ 2 - 1
         .Ref_Top = -lngCR
         .Ref_Height = lngCR
         .Ref_Width = lngWidth + 2 * lngCR
         .Ref_Radius = lngCR
         .RadialGXPercent = 50
         .RadialGYPercent = 40
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight
         .RadialGIntensity = 130
      
      Case MAC_Variation
         .Ref_Intensity = 70
         .Ref_Left = (lngCR \ 3) - 1
         .Ref_Height = lngCR
         .Ref_Width = lngWidth + 2 * lngCR
         .Ref_Top = 0
         .Ref_Radius = (lngCR \ 2)
         .RadialGXPercent = 200
         .RadialGYPercent = 70
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight
         .RadialGIntensity = 130
         
      Case Crystal
         .Ref_Intensity = 50
         .Ref_Left = lngCR \ 2
         .Ref_Height = lngCR * 1.1
         .Ref_Width = lngWidth + 2 * lngCR
         .Ref_Top = 1
         .Ref_Radius = lngCR \ 2
         .RadialGXPercent = 300
         .RadialGYPercent = 60
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight
         .RadialGIntensity = 120
         
      Case Iceblock
         .Ref_Intensity = 50
         .Ref_Left = lngCR \ 2
         .Ref_Top = 2
         .Ref_Height = lngCR + 1
         .Ref_Width = lngWidth - lngCR
         .Ref_Radius = lngCR \ 2
         .RadialGXPercent = 60
         .RadialGYPercent = 60
         .RadialGOffsetX = lngWidth \ 2
         .RadialGOffsetY = lngHeight \ 2
         .RadialGIntensity = 100
      End Select
   End With '// mudtCrystalParam
   
End Sub

Private Sub CalcTextRects()

  Dim lngWidth    As Long
  Dim lngHeight   As Long
  Dim lngCR       As Long
  
   With UserControl
      lngWidth = .ScaleWidth
      lngHeight = .ScaleHeight
      
      If Not mpicButtonPic Is Nothing Then
         mudtPicSz.X = .ScaleX(mpicButtonPic.Width, 8, .ScaleMode)
         mudtPicSz.Y = .ScaleY(mpicButtonPic.Height, 8, .ScaleMode)
      Else
         mudtPicSz.X = 0
         mudtPicSz.Y = 0
      End If
   End With

   lngCR = GetCornerRadius
      
   '// calculate the rects required to draw the text
   With mudtRC2
      Select Case mudtPictureAlignment
      Case PIC_LEFT, PIC_LeftEdge
         .Left = 1 + mudtPicSz.X
         .Right = lngWidth - 2 - lngCR ''' \ 2
         .Top = 1
         .Bottom = lngHeight - 2

      Case PIC_RIGHT
         .Left = 1
         .Right = lngWidth - 2 - mudtPicSz.X
         .Top = 1
         .Bottom = lngHeight - 2
         
      Case PIC_RIGHTEdge
         .Left = 1
         .Right = lngWidth - mudtPicSz.X - lngCR
         .Top = 1
         .Bottom = lngHeight - 2

      Case PIC_TOP
         .Left = 1
         .Right = lngWidth - 2
         .Top = 1 + mudtPicSz.Y
         .Bottom = lngHeight - 2

      Case PIC_BOTTOM
         .Left = 1
         .Right = lngWidth - 2
         .Top = 1
         .Bottom = lngHeight - 2 - mudtPicSz.Y
      End Select
   End With '// mudtRC2

   Call DrawText(UserControl.hdc, mstrCaption, Len(mstrCaption), mudtRC2, C_DT_CALCRECT Or C_DT_WORDBREAK)
   Call CopyRect(mudtRC, mudtRC2)

   Select Case mudtPictureAlignment
   Case PIC_TOP, PIC_LEFT '//Left, Top
      Call OffsetRect(mudtRC, (lngWidth - mudtRC.Right) \ 2, (lngHeight - mudtRC.Bottom) \ 2)
   Case PIC_RIGHTEdge
      Call OffsetRect(mudtRC, (lngWidth - mudtRC.Right - mudtPicSz.X - lngCR - lngCR), (lngHeight - mudtRC.Bottom) \ 2)
   Case PIC_RIGHT '// Right
      Call OffsetRect(mudtRC, (lngWidth - mudtRC.Right - mudtPicSz.X - 4) \ 2, (lngHeight - mudtRC.Bottom) \ 2)
   Case PIC_BOTTOM '// Bottom
      Call OffsetRect(mudtRC, (lngWidth - mudtRC.Right) \ 2, (lngHeight - mudtRC.Bottom - mudtPicSz.Y - 4) \ 2)
   Case PIC_LeftEdge '// Left Edge
      Call OffsetRect(mudtRC, lngCR, (lngHeight - mudtRC.Bottom) \ 2)
   End Select

   Call CopyRect(mudtRC2, mudtRC)
   Call OffsetRect(mudtRC2, 1, 1)
   
   '// once we have the text position we are able to calculate the pic position
   If Not mpicButtonPic Is Nothing Then
      '// if there is no caption, or we have the picture as background
      '// then we put the picture at the center of the button
      If LenB(Trim$(mstrCaption)) > 0 Then
   
         With mudtPicPt
            Select Case mudtPictureAlignment
            Case PIC_LEFT '// left
               .X = mudtRC.Left - mudtPicSz.X - 4
               .Y = (lngHeight - mudtPicSz.Y) \ 2
   
            Case PIC_RIGHT '// right
               .X = mudtRC.Right + 4
               .Y = (lngHeight - mudtPicSz.Y) \ 2
   
            Case PIC_RIGHTEdge '// right Edge
               .X = mudtRC.Right + lngCR
               .Y = (lngHeight - mudtPicSz.Y) \ 2
   
            Case PIC_TOP '// top
               .X = (lngWidth - mudtPicSz.X) \ 2
               .Y = mudtRC.Top - mudtPicSz.Y - 2
   
            Case PIC_BOTTOM '// bottom
               .X = (lngWidth - mudtPicSz.X) \ 2
               .Y = mudtRC.Bottom + 2
   
            Case PIC_LeftEdge '// left edge
               .X = lngCR
               .Y = (lngHeight - mudtPicSz.Y) \ 2
            End Select
         End With
   
      Else '// center the picture
         mudtPicPt.X = (lngWidth - mudtPicSz.X) \ 2
         mudtPicPt.Y = (lngHeight - mudtPicSz.Y) \ 2
      End If
   End If
   
End Sub

Private Function CreateRoundedRegion(ByVal X As Long, _
                                     ByVal Y As Long, _
                                     ByVal lWidth As Long, _
                                     ByVal lHeight As Long) As Long
'/----------------------------------------------------------------------------------/
'/ Description:                                                                     /
'/                                                                                  /
'/ CreateRoundedRegion returns a rounded region based on a given Width, Height      /
'/ and a CornerRadius. We will use this function instead of normal CreateRoundRect  /
'/ because this will give us a better rounded rectangle for our purposes.           /
'/----------------------------------------------------------------------------------/

  Dim lngI     As Long
  Dim lngJ     As Long
  Dim lngI2    As Long
  Dim lngJ2    As Long
  Dim hRgn     As Long
  Dim lngCR    As Long

   lngCR = GetCornerRadius
   
   '// Create initial region
   hRgn = CreateRectRgn(0, 0, X + lWidth, Y + lHeight)

   For lngJ = 0 To Y + lHeight
      For lngI = 0 To (X + lWidth) \ 2

         If Not IsInRoundRect(lngI, lngJ, X, Y, lWidth, lHeight, lngCR) Then
            '// substract the pixels outside of the rounded rectangle (it doesn't exclude the border)
            If Not lngJ = lngJ2 Then
               ExcludePixelsFromRegion hRgn, X + lWidth - lngI2, lngJ2, lWidth - lngI, lngJ

               If Not 2 * lngI2 = X + lWidth Then
                  lngI2 = lngI2 + 1
               End If

               ExcludePixelsFromRegion hRgn, lngI, lngJ, lngI2, lngJ2
            End If

            lngI2 = lngI
            lngJ2 = lngJ
         End If

      Next lngI
   Next lngJ
   
   CreateRoundedRegion = hRgn

End Function

Public Property Get DisplayHand() As Boolean

   DisplayHand = mblnDisplayHand

End Property

Public Property Let DisplayHand(ByVal vNewValue As Boolean)

   mblnDisplayHand = vNewValue
   PropertyChanged "DisplayHand"
   
   Call HandCursorVisible

End Property

Private Sub DrawButton(Optional ByVal vState As enuCandy_State = eNormal)

   UserControl.Cls

   If mudtButtonBehaviour And mblnIsChecked Then
      vState = ePressed
   End If

   Select Case mudtStyle
   Case Plastic
      Call DrawPlasticButton(vState)
   Case Else
      Call DrawCrystalButton(vState)
   End Select

   Call DrawCaption(vState)
   Call DrawPictures(vState)
   Call DrawFocusR

End Sub

Private Sub DrawCrystalButton(ByVal vState As enuCandy_State)

  Dim lWidth      As Long
  Dim lHeight     As Long
  Dim lngI        As Long
  Dim lngJ        As Long
  Dim ptColor     As Long
  '''Dim RGXPercent  As Single
  '''Dim RGYPercent  As Single
  Dim hHlRgn      As Long
  Dim Bordercolor As Long
  Dim nBrush      As Long
  Dim ClientRct   As typRECT
  Dim lngColorB   As Long
  Dim lngButtonC  As Long
  Dim lngCR       As Long

   With UserControl
      lWidth = .ScaleWidth
      lHeight = .ScaleHeight
   End With
   
   If Not mblnIsEnabled Then
      lngColorB = vbWhite
      lngButtonC = 11583680
   
   Else
      lngColorB = mclrButtonBright
      Select Case vState
      Case eHover
         lngButtonC = mclrButtonHover
   
      Case ePressed
         lngButtonC = mclrButtonDown
   
      Case eNormal
         lngButtonC = mclrButtonUp
      End Select
   End If
   
   '''RGYPercent = (100 - mudtCrystalParam.RadialGYPercent) / (lHeight * 2)
   '''RGXPercent = (100 - mudtCrystalParam.RadialGXPercent) / lWidth

   '// Get Border color
   If mlngBorderBrightness >= 0 Then
      Bordercolor = BlendColors(lngButtonC, vbWhite, mlngBorderBrightness)
   Else
      Bordercolor = BlendColors(lngButtonC, vbBlack, -mlngBorderBrightness)
   End If
   '// Get corner radius
   lngCR = GetCornerRadius
   
   '// Create Highlight region (hHlRgn), we will use PtInRegion to
   '// check if we are inside the highlight Rounded rectangle
   '// you could simply use IsInRoundRect(lngI ,lngJ ,mudtCrystalParam.Ref_Left, mudtCrystalParam.Ref_Top, mudtCrystalParam.Ref_Width, mudtCrystalParam.Ref_Height, mudtCrystalParam.Ref_Radius * 2, mudtCrystalParam.Ref_Radius * 2)
   '// instead of PtInRegion and remove these lines, but will be slower.
   hHlRgn = CreateRoundRectRgn(mudtCrystalParam.Ref_Left, mudtCrystalParam.Ref_Top, mudtCrystalParam.Ref_Width, mudtCrystalParam.Ref_Height, mudtCrystalParam.Ref_Radius * 2, mudtCrystalParam.Ref_Radius * 2)
   
   '// Paint the Background
   SetRect ClientRct, 0, 0, lWidth, lHeight
   nBrush = CreateSolidBrush(lngButtonC)
   FillRect hdc, ClientRct, nBrush
   DeleteObject nBrush
   
   '// Draw a radial Gradient
   If mudtStyle = GlowButton Then
      If Not vState = eNormal Or Not Ambient.UserMode Then '// Hover state or design mode
         DrawElipse UserControl.hdc, mudtCrystalParam, lWidth, lHeight, lngButtonC, lngColorB
      Else
         DrawElipse UserControl.hdc, mudtCrystalParam, lWidth, lHeight, lngButtonC, BlendColors(lngButtonC, mclrButtonBright, 50)
      End If
   Else
      DrawElipse UserControl.hdc, mudtCrystalParam, lWidth, lHeight, lngButtonC, lngColorB
   End If

   For lngJ = 0 To lHeight
      For lngI = 0 To lWidth \ 2

         If PtInRegion(mlngButtonRegion, lngI, lngJ) Then
            '// We are inside the button
            If PtInRegion(hHlRgn, lngI, lngJ) Then
               ptColor = BlendColors(vbWhite, lngButtonC, lngJ * mudtCrystalParam.Ref_Intensity \ lngCR)
               Line (lngI, lngJ)-(lWidth - lngI + 1, lngJ), ptColor
               lngI = 0
               lngJ = lngJ + 1
            End If

         Else
            '// this draw a thin border
            SetPixelV hdc, lngI, lngJ, Bordercolor
            SetPixelV hdc, lWidth - lngI, lngJ, Bordercolor
         End If

      Next lngI
   Next lngJ
         
   DeleteObject hHlRgn

End Sub

Private Sub DrawElipse(ByVal lHDC As Long, ByRef CrystalParam As typCrystalParam, ByVal lWidth As Long, _
                       ByVal lHeight As Long, ByVal FromColor As Long, ByVal ToColor As Long)

  Dim oldBrush As Long
  Dim newBrush As Long
  Dim newPen   As Long
  Dim oldPen   As Long
  Dim incX     As Single
  Dim incY     As Single
  Dim RadX     As Long
  Dim RadY     As Long
  Dim klr      As Long
  Dim Rc       As typRECT

   On Error Resume Next
   klr = 1
   RadX = CrystalParam.RadialGXPercent * lWidth / 100
   RadY = CrystalParam.RadialGYPercent * lHeight / 100
   SetRect Rc, CrystalParam.RadialGOffsetX - RadX, CrystalParam.RadialGOffsetY - RadY, CrystalParam.RadialGOffsetX + RadX, CrystalParam.RadialGOffsetY + RadY
   
   incX = 1
   incY = 1

   If RadX > RadY Then
      incX = (RadX / RadY)
   Else
      incY = (RadY / RadX)
   End If

   newBrush = CreateSolidBrush(FromColor)
   oldBrush = SelectObject(lHDC, newBrush)
   newPen = CreatePen(5, 0, FromColor)
   oldPen = SelectObject(lHDC, newPen)

   Do Until Not IsRectEmpty(Rc) = 0
      Ellipse lHDC, Rc.Left, Rc.Top, Rc.Right, Rc.Bottom
      InflateRect Rc, -incX, -incY
      klr = klr + 1
      newBrush = CreateSolidBrush(BlendColors(FromColor, ToColor, klr * CrystalParam.RadialGIntensity / RadY))
      DeleteObject SelectObject(lHDC, newBrush)
   Loop

   DeleteObject SelectObject(lHDC, oldBrush)
   DeleteObject SelectObject(lHDC, oldPen)

End Sub

Private Sub DrawCaption(ByVal vState As enuCandy_State)
   
   If LenB(mstrCaption) Then
      If Not mblnIsEnabled Then
         Call SetTextColor(UserControl.hdc, ConvertFromSystemColor(vbGrayText))
         Call DrawText(UserControl.hdc, mstrCaption, Len(mstrCaption), mudtRC, C_DT_CENTER)
         
      Else
         Select Case vState
         Case ePressed
            Call SetTextColor(UserControl.hdc, ConvertFromSystemColor(mclrCapForecolor))
            Call DrawText(UserControl.hdc, mstrCaption, Len(mstrCaption), mudtRC2, C_DT_CENTER)
         
         Case eHover
            If mblnCapHighLight Then
               Call SetTextColor(UserControl.hdc, ConvertFromSystemColor(mclrCapHighLightColor))
            Else
               Call SetTextColor(UserControl.hdc, ConvertFromSystemColor(mclrCapForecolor))
            End If
            Call DrawText(UserControl.hdc, mstrCaption, Len(mstrCaption), mudtRC, C_DT_CENTER)
         
         Case Else '// Normal
            Call SetTextColor(UserControl.hdc, ConvertFromSystemColor(mclrCapForecolor))
            Call DrawText(UserControl.hdc, mstrCaption, Len(mstrCaption), mudtRC, C_DT_CENTER)
         End Select
      End If
   End If
   
End Sub

Private Sub DrawFocusR()

   With UserControl
      If mblnShowFocus Then
         If mblnHasFocus Then
            Dim lngO  As Long
            Dim lngCR As Long
            ''' Call SetTextColor(UserControl.hdc, mclrCapForecolor)
            ''' Call DrawFocusRect(UserControl.hdc, mudtRC3)
            .DrawStyle = 2
            lngCR = GetCornerRadius
            lngO = lngCR \ 2
            Call RoundRect(.hdc, lngO, lngO, .ScaleWidth - lngO, .ScaleHeight - lngO, lngCR, lngCR)
            .DrawStyle = 0
         End If
      End If
   End With
   
End Sub

Private Sub DrawPictures(ByVal vbytState As enuCandy_State)

  Dim lngClrMask As Long
   '// check if there is a main picture, if not then exit
   If Not mpicButtonPic Is Nothing Then
   
      lngClrMask = ConvertFromSystemColor(mlngMaskColor)

      With UserControl
         If Not mblnIsEnabled Then
            Call TransBlt(.hdc, mudtPicPt.X + 1, mudtPicPt.Y + 1, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask)
         
         Else
            Select Case vbytState
            Case eHover
               If mblnPicHighLight Then
                  Call TransBlt(.hdc, mudtPicPt.X + 1, mudtPicPt.Y + 1, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask, mclrButtonUp)
                  Call TransBlt(.hdc, mudtPicPt.X - 1, mudtPicPt.Y - 1, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask)
               Else
                  Call TransBlt(.hdc, mudtPicPt.X, mudtPicPt.Y, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask)
               End If
   
            Case ePressed '// down
               Call TransBlt(.hdc, mudtPicPt.X + 1, mudtPicPt.Y + 1, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask)
   
            Case Else '// Normal
               Call TransBlt(.hdc, mudtPicPt.X, mudtPicPt.Y, mudtPicSz.X, mudtPicSz.Y, mpicButtonPic, lngClrMask, , , mblnUseGrey)
            End Select
         End If
      End With '// UserControl

   End If '// Not mpicButtonPic Is Nothing

End Sub

Private Sub DrawPlasticButton(ByVal vState As enuCandy_State)

  Dim lngI           As Long
  Dim lngJ           As Long
  Dim HighlightColor As Long
  Dim ShadowColor    As Long
  Dim ptColor        As Long
  Dim LinearGPercent As Long
  Dim lngBaseColor   As Long
  Dim lWidth         As Long
  Dim lHeight        As Long
  Dim lngCR          As Long

   lngCR = GetCornerRadius
   
   With UserControl
      lWidth = .ScaleWidth '- 1
      lHeight = .ScaleHeight '- 1
   End With
   
   Select Case vState
   Case eHover
      lngBaseColor = mclrButtonHover
   Case ePressed
      lngBaseColor = mclrButtonDown
   Case eNormal
      lngBaseColor = mclrButtonUp
   End Select
      
   ShadowColor = BlendColors(vbBlack, lngBaseColor, 50)
   For lngJ = 0 To lHeight

      If lngJ < lngCR Then
         HighlightColor = BlendColors(vbWhite, lngBaseColor, lngJ * 30 \ lngCR)
      End If

      LinearGPercent = Abs((2 * lngJ - lHeight) * 100 \ lHeight)

      For lngI = 0 To lWidth \ 2

         If IsInRoundRect(lngI, lngJ, 1, 1, lWidth - 2, lHeight - 2, lngCR) Then
            
            '// Drawing the button properly
            If IsInRoundRect(lngI, lngJ, 4, 2, lWidth - lngCR, 2 * lngCR - 1, 2 * lngCR \ 3) And Not _
               IsInRoundRect(lngI, lngJ, 4, lngCR \ 2, lWidth - lngCR, 2 * lngCR - 1, 2 * lngCR \ 3) Then
               ptColor = HighlightColor '// draw reflected highlight
               
            Else
               ptColor = BlendColors(lngBaseColor, mclrButtonBright, LinearGPercent)
            End If

            SetPixelV hdc, lngI, lngJ, ptColor
            SetPixelV hdc, lWidth - lngI, lngJ, ptColor
            
         ElseIf IsInRoundRect(lngI, lngJ, 0, 0, lWidth, lHeight, lngCR) Then
            '// this draw a thin border
            SetPixelV hdc, lngI, lngJ, ShadowColor
            SetPixelV hdc, lWidth - lngI, lngJ, ShadowColor
         End If

      Next lngI
   Next lngJ

End Sub

Public Property Get Enabled() As Boolean

   Enabled = mblnIsEnabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)

   mblnIsEnabled = vNewValue
   PropertyChanged "Enabled"
   Call DrawButton
   UserControl.Enabled = mblnIsEnabled

End Property

Private Sub ExcludePixelsFromRegion(hRgn As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)

  Dim hRgnTemp As Long

   hRgnTemp = CreateRectRgn(X1, Y1, X2, Y2)
   CombineRgn hRgn, hRgn, hRgnTemp, RGN_XOR
   DeleteObject hRgnTemp

End Sub

Public Property Get Font() As StdFont

   Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal newFont As StdFont)

   Set UserControl.Font = newFont
   PropertyChanged "Font"
   Call CalcTextRects
   Call DrawButton

End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = mclrCapForecolor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   mclrCapForecolor = NewForeColor
   UserControl.ForeColor = mclrCapForecolor
   PropertyChanged "ForeColor"
   Call DrawButton

End Property

Private Function GetCornerRadius() As Long
   
   If mlngUserCornerRadius < 1 Then
      GetCornerRadius = mlngCornerRadius
   Else
      GetCornerRadius = mlngUserCornerRadius
   End If
   If GetCornerRadius < 1 Then GetCornerRadius = 1

End Function

Private Sub GetRGB(ByRef rlngR As Long, ByRef rlngG As Long, ByRef rlngB As Long, ByVal vColor As Long)

   vColor = ConvertFromSystemColor(vColor)
   
   rlngR = vColor And &HFF&
   rlngG = (vColor And &HFF00&) \ &H100&
   rlngB = (vColor And &HFF0000) \ &H10000

End Sub

Private Sub HandCursorVisible()

  Dim lngHandle      As Long
  Dim picHandPointer As StdPicture
  Const IDC_HAND     As Long = 32649
  
   If Ambient.UserMode Then '// If we're not in design mode
      If mblnDisplayHand Then
         '// Get handle to Hand Pointer icon
         lngHandle = LoadCursor(0, IDC_HAND)
         
         If Not lngHandle = 0 Then
            '// use function to convert memory handle to stdPicture
            '// so we can apply it to the MouseIcon
            Set picHandPointer = HandCursorHandleToPicture(lngHandle, False)
         End If
   
         If Not picHandPointer Is Nothing Then
            UserControl.MouseIcon = picHandPointer
            UserControl.MousePointer = vbCustom
         End If
         
         Set picHandPointer = Nothing
      
      Else
         Set UserControl.MouseIcon = Nothing
         UserControl.MousePointer = vbDefault
      End If
   End If
   
End Sub

Private Function HandCursorHandleToPicture(ByVal hHandle As Long, ByVal isBitmap As Boolean) As IPicture

  Dim udtPIC         As typPICTDESC
  Dim guid(0 To 3)   As Long
    
   '// Convert an icon/bitmap handle to a Picture object
   On Error GoTo ExitRoutine
   
   '// initialize the udtPIC structure
   With udtPIC
      .cbSize = Len(udtPIC)
      .hIcon = hHandle
      If isBitmap Then
         .pictType = vbPicTypeBitmap
      Else
         .pictType = vbPicTypeIcon
      End If
   End With

   '// this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
   '// we use an array of Long to initialize it faster
   guid(0) = &H7BF80980
   guid(1) = &H101ABF32
   guid(2) = &HAA00BB8B
   guid(3) = &HAB0C3000
   '// create the picture,
   '// return an object reference right into the function result
   OleCreatePictureIndirect udtPIC, guid(0), True, HandCursorHandleToPicture
   Erase guid

ExitRoutine:
End Function

Public Property Get hWnd() As Long

   hWnd = UserControl.hWnd

End Property

Public Property Get PictureHighLight() As Boolean
   
   '// enable picture highlighting when mouse over
   PictureHighLight = mblnPicHighLight

End Property

Public Property Let PictureHighLight(ByVal vNewValue As Boolean)

   mblnPicHighLight = vNewValue
   PropertyChanged "PicHighLight"

End Property

Private Sub Init_Style()
   '/----------------------------------------------------------------------------------/
   '/ Description:                                                                     /
   '/                                                                                  /
   '/ Init_Style will create the window region according to the button style           /
   '/ and will be responsible of storing the same region (but without the border)      /
   '/ in mlngButtonRegion. This will be used later to determine if a point             /
   '/ is inside the button region.                                                     /
   '/----------------------------------------------------------------------------------/

   '// Remove the older Region
   If mlngButtonRegion Then DeleteObject mlngButtonRegion
   
   With UserControl
   
      'If mlngCornerRadius < 1 Then
         Select Case mudtStyle
         Case Crystal, WMP, MAC_Variation
            mlngCornerRadius = SetBound(.ScaleHeight \ 2 + 1, 1, .ScaleWidth \ 2)
         Case MAC, GlowButton
            mlngCornerRadius = 12
         Case Iceblock
            mlngCornerRadius = SetBound(.ScaleHeight \ 4 + 1, 1, .ScaleWidth \ 4)
         Case Plastic
            mlngCornerRadius = SetBound(.ScaleHeight \ 3, 1, .ScaleWidth \ 3)
         End Select
      'End If
      
      mlngButtonRegion = CreateRoundedRegion(0, 0, ScaleWidth, .ScaleHeight)
      
      '// Set the Button Region
      Call SetWindowRgn(.hWnd, mlngButtonRegion, True)
      DeleteObject mlngButtonRegion
      
      '// Store the region but exclude the border
      mlngButtonRegion = CreateRoundedRegion(1, 1, .ScaleWidth - 2, .ScaleHeight - 2)
      
      '''// Focus Rect
      '''GetClientRect .hWnd, mudtRC3
      '''InflateRect mudtRC3, -mlngCornerRadius \ 2, -4
      
   End With
   
   Call CalcButtonParam
   
End Sub

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  
  '// Determine if the passed function is supported
  Dim hmod        As Long
  Dim bLibLoaded  As Boolean

   hmod = GetModuleHandleA(sModule)

   If hmod = 0 Then
      hmod = LoadLibraryA(sModule)
      If hmod Then
         bLibLoaded = True
      End If
   End If

   If hmod Then
      If GetProcAddress(hmod, sFunction) Then
         IsFunctionExported = True
      End If
   End If

   If bLibLoaded Then
      FreeLibrary hmod
   End If

End Function

Private Function IsInCircle(ByVal vlngX As Long, ByVal vlngY As Long, ByVal vlngR As Long) As Boolean

  Dim lResult As Long

   '// this detect a circunference centered on y=-vlngR and x=0
   lResult = (vlngR * vlngR) - (vlngX * vlngX)
   If lResult >= 0 Then
      lResult = Sqr(lResult)
      IsInCircle = (Abs(vlngY - vlngR) < lResult)
   End If

End Function

Private Function IsInRoundRect(ByVal vlngI As Long, _
                               ByVal vlngJ As Long, _
                               ByVal vlngX As Long, _
                               ByVal vlngY As Long, _
                               ByVal vlngWidth As Long, _
                               ByVal vlngHeight As Long, _
                               ByVal vlngRadius As Long) As Boolean

  Dim lngOffX As Long
  Dim lngOffY As Long

   lngOffX = vlngI - vlngX
   lngOffY = vlngJ - vlngY

   If lngOffY > vlngRadius And lngOffY + vlngRadius < vlngHeight And lngOffX > vlngRadius And lngOffX + vlngRadius < vlngWidth Then
      IsInRoundRect = True '// This is to catch early most cases
   ElseIf lngOffX < vlngRadius And lngOffY <= vlngRadius Then
      IsInRoundRect = IsInCircle(lngOffX - vlngRadius, lngOffY, vlngRadius)
   ElseIf lngOffX + vlngRadius > vlngWidth And lngOffY <= vlngRadius Then
      IsInRoundRect = IsInCircle(lngOffX - vlngWidth + vlngRadius, lngOffY, vlngRadius)
   ElseIf lngOffX < vlngRadius And lngOffY + vlngRadius >= vlngHeight Then
      IsInRoundRect = IsInCircle(lngOffX - vlngRadius, lngOffY - vlngHeight + vlngRadius * 2, vlngRadius)
   ElseIf lngOffX + vlngRadius > vlngWidth And lngOffY + vlngRadius >= vlngHeight Then
      IsInRoundRect = IsInCircle(lngOffX - vlngWidth + vlngRadius, lngOffY - vlngHeight + vlngRadius * 2, vlngRadius)
   Else
      IsInRoundRect = (lngOffX > 0 And lngOffX < vlngWidth And lngOffY > 0 And lngOffY < vlngHeight)
   End If

End Function

Public Property Get MaskColor() As OLE_COLOR

   MaskColor = mlngMaskColor

End Property

Public Property Let MaskColor(ByVal vNewValue As OLE_COLOR)

   mlngMaskColor = vNewValue
   Call DrawButton
   PropertyChanged "MaskColor"

End Property

Public Property Get Picture() As StdPicture

   Set Picture = mpicButtonPic

End Property

Public Property Set Picture(Value As StdPicture)

   Set mpicButtonPic = Value
   PropertyChanged "Picture"
   Call CalcTextRects
   Call DrawButton

End Property

Public Property Get PictureAlignment() As enuCandy_Alignment

   PictureAlignment = mudtPictureAlignment

End Property

Public Property Let PictureAlignment(ByVal vNewValue As enuCandy_Alignment)

   If Not vNewValue = mudtPictureAlignment Then
      mudtPictureAlignment = vNewValue
      PropertyChanged "PictureAlignment"
      Call CalcTextRects
      Call DrawButton
   End If

End Property

Public Property Get PictureModeDisabled() As enuCandy_DisablePicMode

   PictureModeDisabled = mudtDisabledPicMode
    
End Property

Public Property Let PictureModeDisabled(ByVal vNewValue As enuCandy_DisablePicMode)
 
   mudtDisabledPicMode = vNewValue
   Call DrawButton
   PropertyChanged "DisabledPicMode"
   
End Property

Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal uMsgWhen As eMsgWhen = eMsgWhen.MSG_AFTER)

   'Add the message value to the window handle's specified callback table
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its  memory
      If uMsgWhen And MSG_BEFORE Then              'If the message is to be added to the before original WndProc table...
         zAddMsg uMsg, IDX_BTABLE                  'Add the message to the before table
      End If

      If uMsgWhen And MSG_AFTER Then               'If message is to be added to the after original WndProc table...
         zAddMsg uMsg, IDX_ATABLE                  'Add the message to the after table
      End If
   End If

End Sub

Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long
   'Call the original WndProc
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
   End If

End Function

Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

   'Delete the message value from the window handle's specified callback table
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its  memory
      If When And MSG_BEFORE Then                  'If the message is to be deleted from the before original WndProc table...
         zDelMsg uMsg, IDX_BTABLE                  'Delete the message from the before table
      End If

      If When And MSG_AFTER Then                   'If the message is to be deleted from the after original WndProc table...
         zDelMsg uMsg, IDX_ATABLE                  'Delete the message from the after table
      End If

   End If

End Sub

Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long

   'Get the subclasser lParamUser callback parameter
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      sc_lParamUser = zData(IDX_PARM_USER)         'Get the lParamUser callback parameter
   End If

End Property

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal vNewValue As Long)
   
   'Let the subclasser lParamUser callback parameter
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its  memory
      zData(IDX_PARM_USER) = vNewValue              'Set the lParamUser callback parameter
   End If

End Property

Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'-SelfSub code------------------------------------------------------------------------------------

   '*************************************************************************************************
   '* lng_hWnd   - Handle of the window to subclass
   '* lParamUser - Optional, user-defined callback parameter
   '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method,
   '   etc.
   '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
   '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl
   '   for design-time subclassing
   '*************************************************************************************************
  Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
  Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg  tables
  Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
  Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
  Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
  Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
  Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
  Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
  Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
  Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
  Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
  Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
  Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
  Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long

   If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
      zError SUB_NAME, "Invalid window handle"
      Exit Function
   End If

   nMyID = GetCurrentProcessId                                               'Get this process's ID
   GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
   If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
      zError SUB_NAME, "Window handle belongs to another process"
      Exit Function
   End If

   If oCallback Is Nothing Then                                               'If the user hasn't specified the callback owner
      Set oCallback = Me                                                      'Then it is me
   End If

   nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
   If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
   End If

   If z_Funk Is Nothing Then              'If this is the first time through, do the one-time initialization
      Set z_Funk = New Collection         'Create the hWnd/thunk-address collection
      z_Sc(14) = &HD231C031
      z_Sc(15) = &HBBE58960
      z_Sc(17) = &H4339F631
      z_Sc(18) = &H4A21750C
      z_Sc(19) = &HE82C7B8B
      z_Sc(20) = &H74&
      z_Sc(21) = &H75147539
      z_Sc(22) = &H21E80F
      z_Sc(23) = &HD2310000
      z_Sc(24) = &HE8307B8B
      z_Sc(25) = &H60&
      z_Sc(26) = &H10C261
      z_Sc(27) = &H830C53FF
      z_Sc(28) = &HD77401F8
      z_Sc(29) = &H2874C085
      z_Sc(30) = &H2E8&
      z_Sc(31) = -1447168
      z_Sc(32) = &H75FF3075
      z_Sc(33) = &H2875FF2C
      z_Sc(34) = &HFF2475FF
      z_Sc(35) = &H3FF2473
      z_Sc(36) = -1995418625
      z_Sc(37) = &HBFF1C45
      z_Sc(38) = &H73396775
      z_Sc(39) = &H58627404
      z_Sc(40) = &H6A2473FF
      z_Sc(41) = &H873FFFC
      z_Sc(42) = &H891453FF
      z_Sc(43) = &H7589285D
      z_Sc(44) = &H3045C72C
      z_Sc(45) = &H8000&
      z_Sc(46) = &H8920458B
      z_Sc(47) = &H4589145D
      z_Sc(48) = &HC4836124
      z_Sc(49) = &H1862FF04
      z_Sc(50) = 904073099
      z_Sc(51) = &HA78C985
      z_Sc(52) = &H8B04C783
      z_Sc(53) = &HAFF22845
      z_Sc(54) = &H73FF2775
      z_Sc(55) = 475266856
      z_Sc(56) = &H438D1F75
      z_Sc(57) = &H144D8D34
      z_Sc(58) = &H1C458D50
      z_Sc(59) = &HFF3075FF
      z_Sc(60) = 1979657333
      z_Sc(61) = &H873FF28
      z_Sc(62) = &HFF525150
      z_Sc(63) = &H53FF2073
      z_Sc(64) = &HC328&

      z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk  data
      z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
   End If

   z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

   If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
      On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                     'Add the hWnd/thunk-address to the collection
      On Error GoTo 0

      If bIdeSafety Then                                                      'If the user wants IDE protection
         z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                         'Store the EbMode function address in the thunk data
      End If

      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data

      nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
      If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
         zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
         GoTo ReleaseMemory
      End If

      z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
      sc_Subclass = True                                                      'Indicate success
   Else
      zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
   End If

   Exit Function                                               'Exit sc_Subclass

CatchDoubleSub:
   zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
   VirtualFree z_ScMem, 0, MEM_RELEASE                         'sc_Subclass has failed after memory allocation, so release the memory

End Function

Private Sub sc_Terminate()
  
  Dim lngI As Long

   'Terminate all subclassing
   If Not (z_Funk Is Nothing) Then                 'Ensure that subclassing has been started

      With z_Funk

         For lngI = .Count To 1 Step -1            'Loop through the collection of window handles in reverse order
            z_ScMem = .Item(lngI)                  'Get the thunk address
            If IsBadCodePtr(z_ScMem) = 0 Then      'Ensure that the thunk hasn't already released its memory
               sc_UnSubclass zData(IDX_HWND)       'UnSubclass
            End If

         Next lngI                                 'Next member of the collection
      End With

      Set z_Funk = Nothing                         'Destroy the hWnd/thunk-address collection
   End If

End Sub

Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)

   'UnSubclass the specified window handle
   If z_Funk Is Nothing Then                                   'Ensure that subclassing has been started
      zError "sc_UnSubclass", "Window handle isn't subclassed"
   
   Else
      If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then            'Ensure that the thunk hasn't already released its memory
         zData(IDX_SHUTDOWN) = -1                              'Set the shutdown indicator
         zDelMsg ALL_MESSAGES, IDX_BTABLE                      'Delete all before messages
         zDelMsg ALL_MESSAGES, IDX_ATABLE                      'Delete all after messages
      End If

      z_Funk.Remove "h" & lng_hWnd                             'Remove the specified window handle from the collection
   End If

End Sub

Private Function SetBound(ByVal Num As Long, ByVal MinNum As Long, ByVal MaxNum As Long) As Long

   '// make sure Num is within allowable limits
   Select Case Num
   Case Is < MinNum
      SetBound = MinNum
   Case Is > MaxNum
      SetBound = MaxNum
   Case Else
      SetBound = Num
   End Select

End Function

Public Property Get ButtonBehaviour() As enuCandy_ButtonBehaviour

   ButtonBehaviour = mudtButtonBehaviour

End Property

Public Property Let ButtonBehaviour(ByVal vNewValue As enuCandy_ButtonBehaviour)

   mudtButtonBehaviour = vNewValue
   Call DrawButton
   PropertyChanged "ButtonBehaviour"

End Property

Public Property Get ButtonStyle() As enuCandy_Style

   ButtonStyle = mudtStyle

End Property

Public Property Let ButtonStyle(ByVal vNewValue As enuCandy_Style)

   If Not vNewValue = mudtStyle Then
      mudtStyle = vNewValue
      PropertyChanged "Style"
      Call Init_Style
      Call DrawButton
   End If

End Property

Private Sub TransBlt(ByVal vlDstDC As Long, _
                     ByVal vlDstX As Long, _
                     ByVal vlDstY As Long, _
                     ByVal vlDstW As Long, _
                     ByVal vlDstH As Long, _
                     ByVal vSrcPic As StdPicture, _
                     Optional ByVal vlTransColor As Long = -1, _
                     Optional ByVal vlBrushColor As Long = -1, _
                     Optional ByVal vbMonoMask As Boolean = False, _
                     Optional ByVal vbGreyscale As Boolean = False)

   Const DI_NORMAL      As Long = &H3
   Const C_MAX          As Long = &HFF
   Const C_MAXP         As Long = &H100
   
   Dim lngB             As Long
   Dim lngH             As Long
   Dim lngF             As Long
   Dim lngI             As Long
   Dim TmpDC            As Long
   Dim TmpBmp           As Long
   Dim TmpObj           As Long
   Dim Sr2DC            As Long
   Dim Sr2Bmp           As Long
   Dim Sr2Obj           As Long
   Dim DataDest()       As typRGBTRIPLE
   Dim DataSrc()        As typRGBTRIPLE
   Dim udtInfo          As typBITMAPINFO
   Dim BrushRGB         As typRGBTRIPLE
   Dim lngCLR           As Long
   Dim SrcDC            As Long
   Dim tObj             As Long
   Dim lngTemp          As Long
   Dim lngOpacity       As Long
   Dim aLighten(C_MAX)  As Byte

   '// make transparent and grayscale images
   
   '// prevent errors
   If vlDstW = 0 Or vlDstH = 0 Then Exit Sub
   If vSrcPic Is Nothing Then Exit Sub

   If Not mblnIsEnabled Then
      vbGreyscale = (mudtDisabledPicMode = Grayed)
      If Not vbGreyscale Then
         '// Buid Lighten array
         For lngI = 0 To C_MAX
            lngTemp = (&H125 * lngI) \ C_MAX
            If lngTemp > C_MAX Then
               lngTemp = C_MAX
            End If
            aLighten(lngI) = lngTemp
         Next lngI
      End If
      
      lngOpacity = 50
      
   Else
      lngOpacity = C_MAX
   End If
   
   SrcDC = CreateCompatibleDC(hdc)

   If vlDstW < 0 Then vlDstW = UserControl.ScaleX(vSrcPic.Width, 8, UserControl.ScaleMode)
   If vlDstH < 0 Then vlDstH = UserControl.ScaleY(vSrcPic.Height, 8, UserControl.ScaleMode)

   If vSrcPic.Type = vbPicTypeBitmap Then '// icon or typBITMAP ?
      tObj = SelectObject(SrcDC, vSrcPic)
      
   Else '// Icon
      Dim hBrush  As Long
      tObj = SelectObject(SrcDC, CreateCompatibleBitmap(vlDstDC, vlDstW, vlDstH))
      hBrush = CreateSolidBrush(vlTransColor)
      DrawIconEx SrcDC, 0, 0, vSrcPic.Handle, vlDstW, vlDstH, 0, hBrush, DI_NORMAL
      DeleteObject hBrush
   End If

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)
   TmpBmp = CreateCompatibleBitmap(vlDstDC, vlDstW, vlDstH)
   Sr2Bmp = CreateCompatibleBitmap(vlDstDC, vlDstW, vlDstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
   
   ReDim DataDest(vlDstW * vlDstH * 3 - 1)
   ReDim DataSrc(UBound(DataDest))
   
   With udtInfo.bmiHeader
      .biSize = Len(udtInfo.bmiHeader)
      .biWidth = vlDstW
      .biHeight = vlDstH
      .biPlanes = 1
      .biBitCount = 24
   End With

   BitBlt TmpDC, 0, 0, vlDstW, vlDstH, vlDstDC, vlDstX, vlDstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, vlDstW, vlDstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, vlDstH, DataDest(0), udtInfo, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, vlDstH, DataSrc(0), udtInfo, 0

   If vlBrushColor > 0 Then
      BrushRGB.rgbBlue = (vlBrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (vlBrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = vlBrushColor And &HFF
   End If

   '// No Maskcolor to use
   If Not mblnUseMask Then vlTransColor = -1

   For lngH = 0 To vlDstH - 1
      lngF = lngH * vlDstW
      
      For lngB = 0 To vlDstW - 1
         lngI = lngF + lngB
         lngTemp = C_MAX - lngOpacity
         
         If GetNearestColor(hdc, CLng(DataSrc(lngI).rgbRed) + C_MAXP& * DataSrc(lngI).rgbGreen + &H10000 * DataSrc(lngI).rgbBlue) <> vlTransColor Then
            
            With DataDest(lngI)
               If vlBrushColor > -1 Then
                  If vbMonoMask Then
                     If (CLng(DataSrc(lngI).rgbRed) + DataSrc(lngI).rgbGreen + DataSrc(lngI).rgbBlue) <= &H180 Then
                        DataDest(lngI) = BrushRGB
                     End If
                  
                  Else
                     If lngOpacity = C_MAX Then
                        DataDest(lngI) = BrushRGB
                     Else
                        .rgbRed = (lngTemp * .rgbRed + lngOpacity * BrushRGB.rgbRed) \ C_MAXP
                        .rgbGreen = (lngTemp * .rgbGreen + lngOpacity * BrushRGB.rgbGreen) \ C_MAXP
                        .rgbBlue = (lngTemp * .rgbBlue + lngOpacity * BrushRGB.rgbBlue) \ C_MAXP
                     End If
                  End If
               
               Else
                  If vbGreyscale Then
                     lngCLR = CLng(DataSrc(lngI).rgbRed * 0.3) + DataSrc(lngI).rgbGreen * 0.59 + DataSrc(lngI).rgbBlue * 0.11
                     If lngOpacity = C_MAX Then
                        .rgbRed = lngCLR
                        .rgbGreen = lngCLR
                        .rgbBlue = lngCLR
                        
                     Else
                        .rgbRed = (lngTemp * .rgbRed + lngOpacity * lngCLR) \ C_MAXP
                        .rgbGreen = (lngTemp * .rgbGreen + lngOpacity * lngCLR) \ C_MAXP
                        .rgbBlue = (lngTemp * .rgbBlue + lngOpacity * lngCLR) \ C_MAXP
                     End If
                     
                  Else
                     If lngOpacity = C_MAX Then
                        DataDest(lngI) = DataSrc(lngI)
                        
                     Else
                        .rgbRed = (lngTemp * .rgbRed + lngOpacity * aLighten(DataSrc(lngI).rgbRed)) \ C_MAXP
                        .rgbGreen = (lngTemp * .rgbGreen + lngOpacity * aLighten(DataSrc(lngI).rgbGreen)) \ C_MAXP
                        .rgbBlue = (lngTemp * .rgbBlue + lngOpacity * aLighten(DataSrc(lngI).rgbBlue)) \ C_MAXP
                     End If
                  End If
               End If
            End With '// DataDest(lngI)
         End If '// GetNearestColor
      Next lngB
   Next lngH

   SetDIBitsToDevice vlDstDC, vlDstX, vlDstY, vlDstW, vlDstH, 0, 0, 0, vlDstH, DataDest(0), udtInfo, 0

   Erase DataDest
   Erase DataSrc
   Erase aLighten
   
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   
   If vSrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
   
   DeleteDC TmpDC
   DeleteDC Sr2DC
   DeleteObject tObj
   DeleteDC SrcDC

End Sub

Public Property Get ShowFocusRect() As Boolean

   ShowFocusRect = mblnShowFocus

End Property

Public Property Let ShowFocusRect(ByVal vNewValue As Boolean)

   mblnShowFocus = vNewValue
   Call DrawButton
   PropertyChanged "ShowFocus"

End Property

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  
  '//Track the mouse leaving the indicated window
  Dim tme As TRACKMOUSEEVENT_STRUCT

   If bTrack Then

      With tme
         .cbSize = Len(tme)
         .dwFlags = TME_LEAVE
         .hwndTrack = lng_hWnd
      End With

      If bTrackUser32 Then
         TrackMouseEvent tme
      Else
         TrackMouseEventComCtl tme
      End If

   End If

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
   
   UserControl_MouseDown vbLeftButton, 0, 0, 0
   RaiseEvent Click
   
End Sub

Private Sub UserControl_Click()

   If mblnIsEnabled Then RaiseEvent Click
   
End Sub

Private Sub UserControl_DblClick()
   
   If mblnIsEnabled Then RaiseEvent DblClick

End Sub

Private Sub UserControl_GotFocus()
   
   mblnHasFocus = True
   If mblnShowFocus Then Call DrawButton

End Sub

Private Sub UserControl_InitProperties()

   UserControl.FontName = "Tahoma"
   UserControl.FontSize = 8
   mblnUseMask = True
   mblnIsEnabled = True
   mclrButtonHover = &HFFC090
   mclrButtonUp = &HE99950
   mclrButtonBright = &HFFEDB0
   mclrButtonDown = &HE99950
   mstrCaption = Ambient.DisplayName
   mlngUserCornerRadius = 0
   mudtStyle = MAC
   mlngMaskColor = vbButtonFace
   mblnPicHighLight = False

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   If mblnIsEnabled Then

      RaiseEvent KeyDown(KeyCode, Shift)
   
      Select Case KeyCode
      Case vbKeySpace '// spacebar pressed
         UserControl_MouseDown vbLeftButton, Shift, 0, 0
   
      Case vbKeyRight, vbKeyDown '// right and down arrows
         SendKeys "{Tab}"
   
      Case vbKeyLeft, vbKeyUp '// left and up arrows
         SendKeys "+{Tab}"
      End Select
   End If
   
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   
   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If mblnIsEnabled Then
      If KeyCode = vbKeySpace Then
         UserControl_MouseUp vbLeftButton, Shift, 0, 0
      End If
      RaiseEvent KeyUp(KeyCode, Shift)
   End If

End Sub

Private Sub UserControl_LostFocus()
   
   mblnHasFocus = False
   If mblnShowFocus Then Call DrawButton
   
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If mblnIsEnabled Then

      If Not Button = vbRightButton Then
      
         Select Case mudtButtonBehaviour
         Case [Check Box]
            mblnIsChecked = Not mblnIsChecked
            Call DrawButton
            
         Case [Option Button]
            If Not mblnIsChecked Then
               mblnIsChecked = True
               Call DrawButton(ePressed)
               Call UncheckAllValues
            End If
         
         Case Else '// Standard
            Call DrawButton(ePressed)
         End Select
      End If
      
      RaiseEvent MouseDown(Button, Shift, X, Y)
   End If
   
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If mblnIsEnabled Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
   End If
   
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If mblnIsEnabled Then
      
      If Not Button = vbRightButton Then
         If mblnIsOver Then
            Call DrawButton(eHover)
         Else
            Call DrawButton
         End If
      End If
      
      RaiseEvent MouseUp(Button, Shift, X, Y)
   End If
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
      mblnIsEnabled = .ReadProperty("Enabled", True)
      mstrCaption = .ReadProperty("Caption", UserControl.Name)
      mblnCapHighLight = .ReadProperty("CaptionHighLight", False)
      mclrCapHighLightColor = .ReadProperty("CaptionHighLightColor", &HFF00&)
      mblnPicHighLight = .ReadProperty("PicHighLight", True)
      mclrCapForecolor = .ReadProperty("ForeColor", vbBlack)
      Set mpicButtonPic = .ReadProperty("Picture", Nothing)
      mudtPictureAlignment = .ReadProperty("PictureAlignment", 0)
      mudtStyle = .ReadProperty("Style", 0)
      mblnIsChecked = .ReadProperty("Checked", mblnIsChecked)
      mclrButtonHover = .ReadProperty("ColorButtonHover", &HFFC090)
      mclrButtonUp = .ReadProperty("ColorButtonUp", &HE99950)
      mclrButtonDown = .ReadProperty("ColorButtonDown", &HE99950)
      mclrButtonBright = .ReadProperty("ColorBright", &HFFEDB0)
      mlngBorderBrightness = .ReadProperty("BorderBrightness", 0)
      mblnDisplayHand = .ReadProperty("DisplayHand", False)
      mudtColorScheme = .ReadProperty("ColorScheme", 0)
      mlngCornerRadius = .ReadProperty("CornerRadius", -1)
      mlngUserCornerRadius = .ReadProperty("UserCornerRadius", -1)
      mudtDisabledPicMode = .ReadProperty("DisabledPicMode", 0)
      mblnUseGrey = .ReadProperty("UseGREY", False)
      mblnUseMask = .ReadProperty("UseMaskColor", True)
      mlngMaskColor = .ReadProperty("MaskColor", vbButtonFace)
      mudtButtonBehaviour = .ReadProperty("ButtonBehaviour", 0)
      mblnShowFocus = .ReadProperty("ShowFocus", False)
   End With
   
   Call CalcTextRects
   Call ColorSchemeSet
   Call HandCursorVisible
   UserControl.ForeColor = mclrCapForecolor
   
   If Ambient.UserMode Then '// If we're not in design mode
      bTrack = True
      bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

      If Not bTrackUser32 Then
         If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
            bTrack = False
         End If
      End If

      If bTrack Then
         '// OS supports mouse leave, so let's subclass for it
         With UserControl
            '// Subclass the UserControl
            sc_Subclass .hWnd
            '''sc_AddMsg .hWnd, WM_PAINT, MSG_BEFORE
            sc_AddMsg .hWnd, WM_MOUSEMOVE
            sc_AddMsg .hWnd, WM_MOUSELEAVE
         End With
      End If
   End If

End Sub

Private Sub UserControl_Resize()

   If UserControl.ScaleHeight > 0 And UserControl.ScaleWidth > 0 Then
      Call Init_Style
      Call CalcTextRects
      Call DrawButton
   End If

End Sub

Private Sub UserControl_Show()

   Call Init_Style
   Call CalcTextRects
   DoEvents
   Call DrawButton

End Sub

Private Sub UserControl_Terminate()
   
   sc_Terminate
   If mlngButtonRegion Then DeleteObject mlngButtonRegion

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Font", UserControl.Font
      .WriteProperty "Enabled", mblnIsEnabled
      .WriteProperty "Caption", mstrCaption
      .WriteProperty "PicHighLight", mblnPicHighLight
      .WriteProperty "CaptionHighLight", mblnCapHighLight
      .WriteProperty "CaptionHighLightColor", mclrCapHighLightColor, &HFF00&
      .WriteProperty "ForeColor", mclrCapForecolor, vbBlack
      .WriteProperty "Picture", mpicButtonPic, Nothing
      .WriteProperty "PictureAlignment", mudtPictureAlignment
      .WriteProperty "Style", mudtStyle
      .WriteProperty "Checked", mblnIsChecked
      .WriteProperty "ColorButtonHover", mclrButtonHover
      .WriteProperty "ColorButtonUp", mclrButtonUp
      .WriteProperty "ColorButtonDown", mclrButtonDown
      .WriteProperty "BorderBrightness", mlngBorderBrightness
      .WriteProperty "ColorBright", mclrButtonBright
      .WriteProperty "DisplayHand", mblnDisplayHand
      .WriteProperty "ColorScheme", mudtColorScheme
      .WriteProperty "CornerRadius", mlngCornerRadius
      .WriteProperty "UserCornerRadius", mlngUserCornerRadius
      .WriteProperty "DisabledPicMode", mudtDisabledPicMode
      .WriteProperty "UseGREY", mblnUseGrey
      .WriteProperty "UseMaskColor", mblnUseMask
      .WriteProperty "MaskColor", mlngMaskColor
      .WriteProperty "ButtonBehaviour", mudtButtonBehaviour
      .WriteProperty "ShowFocus", mblnShowFocus
   End With

End Sub

Public Property Get Value() As Boolean

   Value = mblnIsChecked

End Property

Public Property Let Value(ByVal vNewValue As Boolean)

   mblnIsChecked = vNewValue
   If mudtButtonBehaviour Then '// Option or Check Button?
      Call DrawButton
   End If
   PropertyChanged "VALUE"

End Property

Private Sub UncheckAllValues()

  Dim objCtl As Object
   
   '// Check all controls in parent
   For Each objCtl In Parent.Controls
   
      '// Is it a CandyButton?
      If TypeOf objCtl Is CandyButton Then
         '// is it not this button
         If Not objCtl.hWnd = UserControl.hWnd Then
            '// is the button type Option?
            If objCtl.ButtonBehaviour = [Option Button] Then
               '// Is the button in the same container as this button?
               If objCtl.Container.hWnd = UserControl.ContainerHwnd Then
                  objCtl.Value = False
               End If
            End If
         End If
      End If
   
   Next objCtl

End Sub

Public Property Get UseGreyscale() As Boolean

   UseGreyscale = mblnUseGrey

End Property

Public Property Let UseGreyscale(ByVal vNewValue As Boolean)

   mblnUseGrey = vNewValue

   If Not mpicButtonPic Is Nothing Then
      Call DrawButton
   End If

   PropertyChanged "UseGREY"

End Property

Public Property Get UseMaskColor() As Boolean

   UseMaskColor = mblnUseMask

End Property

Public Property Let UseMaskColor(ByVal vNewValue As Boolean)

   mblnUseMask = vNewValue

   If Not mpicButtonPic Is Nothing Then
      Call DrawButton
   End If

   PropertyChanged "UseMaskColor"

End Property

Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
   '-The following routines are exclusively for the sc_ subclass routines----------------------------
   'Add the message to the specified table of the window handle

  Dim nCount As Long                      'Table entry count
  Dim nBase  As Long                      'Remember z_ScMem
  Dim i      As Long                      'Loop index

   nBase = z_ScMem                        'Remember z_ScMem so that we can restore its value  on exit
   z_ScMem = zData(nTable)                'Map zData() to the specified table

   If uMsg = ALL_MESSAGES Then            'If ALL_MESSAGES are being added to the table...
      nCount = ALL_MESSAGES               'Set the table entry count to ALL_MESSAGES
   Else
      nCount = zData(0)                   'Get the current table entry count
      If nCount >= MSG_ENTRIES Then       'Check for message table overflow
      
         zError "zAddMsg", "Message table overflow. Either increase the" & _
            " value of Const MSG_ENTRIES or use ALL_MESSAGES instead of" & _
            " specific message values"
         GoTo Bail
      End If

      For i = 1 To nCount                 'Loop through the table entries
         If zData(i) = 0 Then             'If the element is free...
            zData(i) = uMsg               'Use this element
            GoTo Bail                     'Bail
         ElseIf zData(i) = uMsg Then      'If the message is already in the table...
            GoTo Bail                     'Bail
         End If
      Next i                              'Next message table entry

      nCount = i                          'On drop through: i = nCount + 1, the new table entry count
      zData(nCount) = uMsg                'Store the message in the appended table entry
   End If

   zData(0) = nCount                      'Store the new table entry count
Bail:
   z_ScMem = nBase                        'Restore the value of z_ScMem

End Sub

Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long

  Dim o As Long   'Object pointer
  Dim i As Long   'vTable entry counter
  Dim j As Long   'vTable address
  Dim n As Long   'Method pointer
  Dim b As Byte   'First method byte
  Dim m As Byte   'Known good first method byte

   o = ObjPtr(oCallback)   'Get the callback object's address
   GetMem4 o, j    'Get the address of the callback object's vTable
   j = j + &H7A4   'Increment to the the first user entry for a usercontrol
   GetMem4 j, n    'Get the method pointer
   GetMem1 n, m    'Get the first method byte... &H33 if pseudo-code, &HE9 if native
   j = j + 4       'Bump to the next vtable entry

   For i = 1 To 511      'Loop through a 'sane' number of vtable entries
      GetMem4 j, n       'Get the method pointer

      If IsBadCodePtr(n) Then   'If the method pointer is an invalid code address
         GoTo vTableEnd         'We've reached the end of the vTable, exit the for loop
      End If

      GetMem1 n, b           'Get the first method byte

      If b <> m Then         'If the method byte doesn't matche the known good value
         GoTo vTableEnd      'We've reached the end of the vTable, exit the for loop
      End If

      j = j + 4              'Bump to the next vTable entry
   Next i                    'Bump counter

   Debug.Assert False                                             'Halt if running under the VB IDE
   Err.Raise vbObjectError, "zAddressOf", "Ordinal not found"     'Raise error if running compiled

vTableEnd:
   'We've hit the end of the vTable
   GetMem4 j - (nOrdinal * 4), n   'Get the method pointer for the specified ordinal
   zAddressOf = n                  'Address of the callback ordinal

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

   RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)

   RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)

  'Delete the message from the specified table of the window handle
  Dim nCount As Long                   'Table entry count
  Dim nBase  As Long                   'Remember z_ScMem
  Dim i      As Long                   'Loop index

   nBase = z_ScMem                     'Remember z_ScMem so that we can restore its value on exit
   z_ScMem = zData(nTable)             'Map zData() to the specified table

   If uMsg = ALL_MESSAGES Then         'If ALL_MESSAGES are being deleted from the table...
      zData(0) = 0                     'Zero the table entry count
      
   Else
      nCount = zData(0)                'Get the table entry count
      For i = 1 To nCount              'Loop through the table entries
         If zData(i) = uMsg Then       'If the message is found...
            zData(i) = 0               'Null the msg value -- also frees the element for re-use
            GoTo Bail                  'Bail
         End If
      Next i                           'Next message table entry

      zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
   End If

Bail:
   z_ScMem = nBase                     'Restore the value of z_ScMem

End Sub

Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
'Error handler

   App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
   MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long

  'Return the address of the specified DLL/procedure
   zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)        'Get the specified procedure address
   Debug.Assert zFnAddr                                           'In the IDE, validate that the procedure address was located

End Function

Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

  'Map zData() to the thunk address for the specified window handle
   If z_Funk Is Nothing Then                                      'Ensure that subclassing has been started
      zError "zMap_hWnd", "Subclassing hasn't been started"
   Else
      On Error GoTo Catch                                         'Catch unsubclassed window handles
      z_ScMem = z_Funk("h" & lng_hWnd)                            'Get the thunk address
      zMap_hWnd = z_ScMem
   End If

   Exit Function                                                  'Exit returning the thunk address

Catch:
   zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
   
   '-Subclass callback, usually ordinal #1, the last method in this source file----------------------
   '*************************************************************************************************
   '* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
   '*              you will know unless the callback for the uMsg value is specified as
   '*              MSG_BEFORE_AFTER (both before and after the original WndProc).
   '* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
   '*              message being passed to the original WndProc and (if set to do so) the after
   '*              original WndProc callback.
   '* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
   '*              and/or, in an after the original WndProc callback, act on the return value as set
   '*              by the original WndProc.
   '* lng_hWnd   - Window handle.
   '* uMsg       - Message value.
   '* wParam     - Message related data.
   '* lParam     - Message related data.
   '* lParamUser - User-defined callback parameter
   '*************************************************************************************************
  Dim X As Long
  Dim Y As Long

   Select Case uMsg
   '''Case WM_PAINT
   '''   Call Init_Style
      
   Case WM_MOUSEMOVE
      If Not mblnIsOver Then
         mblnIsOver = True
         TrackMouseLeave lng_hWnd
         Call DrawButton(eHover)
         RaiseEvent MouseEnter
      End If
   
   Case WM_MOUSELEAVE
      mblnIsOver = False
      Call DrawButton
      RaiseEvent MouseLeave
   End Select

End Sub

