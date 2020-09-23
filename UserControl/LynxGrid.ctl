VERSION 5.00
Begin VB.UserControl LynxGrid 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   ToolboxBitmap   =   "LynxGrid.ctx":0000
End
Attribute VB_Name = "LynxGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Function:  Owner-drawn editable Grid
'*************************************************************************************************
'// Title:    LynxGrid version 2
'// Author:   Morgan Haueisen
'// Modified: 7/12/07
'// Version:  2.15.4a (11 May 2009)(see HistoryLog file for details)
'// MOST CHANGES MADE BY MORGAN HAUEISEN ARE MARKED AS '// MLH
'// Downloaded from: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=70425&lngWId=1

'*************************************************************************************************
' Title:     Original LynxGrid
' Author:    Richard Mewett
' Created:   01/08/05
' Version:   1.89 (10 May 2007)
' Copyright © 2005 Richard Mewett. All rights reserved.
' Provides a combination of MSFlexGrid and ListView (Report Style) functionality.
' Downloaded from: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=65559&lngWId=1
'*************************************************************************************************
' Sub-Classing code
' Author: Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
' v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
' v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
' v1.3 Data array scrapped in favour of property accessors.............................. 20060405
' v1.4 Optional IDE protection added
'      User-defined callback parameter added
'      All user routines that pass in a hWnd get additional validation
'      End removed from zError.......................................................... 20060411
' v1.5 Added nOrdinal parameter to sc_Subclass
'      Switched machine-code array from Currency to Long................................ 20060412
' v1.6 Added an optional callback target object
'      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
' v1.7 Changed zAddressOf, removed zProbe, and added Subs GetMem1 and GetMem4............20080422

'*************************************************************************************************
' Credits:  Gary Noble (Phantom Man)- API Scroll Bar Code
'           Heriberto Mantilla Santamaría - XP Theme API + Alpha Blend
'           Matthew R. Usner - DrawArrow + Beta testing
'           LaVolpe - Bug fixes & numerous suggestions
'           Riccardo Cohen - Bug reports & ownerdrawn XP/Office ThemeStyles
'           Thierry Calu - ComboBox Height adjustment & automatic DropDown
'           John Underhill (Steppenwolfe) - Unicode suggestions / ReturnAddr patch
'           Mario Wallerius (Wallonso) - numerous suggestions & testing
'*************************************************************************************************

' This software is provided "as-is," without any express or implied warranty.
' In no event shall the author be held liable for any damages arising from the
' use of this software.
' If you do not agree with these terms, do not use "LynxGrid". Use of
' the program implicitly means you have agreed to these terms.

' Permission is granted to anyone to use this software for any purpose,
' including commercial use, and to alter and redistribute it, provided that
' the following conditions are met:

' 1. All redistributions of source code files must retain all copyright
'    notices that are currently in place, and this list of conditions without
'    any modification.
' 2. All redistributions in binary form must retain all occurrences of the
'    above copyright notice and web site addresses that are currently in
'    place (for example, in the About boxes).
' 3. Modified versions in source or binary form must be plainly marked as
'    such, and must not be misrepresented as being the original software.
'*************************************************************************************************

Option Explicit
'// Option Compare Text - Can't be used because Find has the option to case match.
Private mblnDrwGrid As Boolean

'-Selfsub-class declarations----------------------------------------------------------------------------

Private Enum eMsgWhen                                    'When to callback
   MSG_BEFORE = 1                                        'Callback before the original WndProc
   MSG_AFTER = 2                                         'Callback after the original WndProc
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER            'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                 'All messages callback
Private Const MSG_ENTRIES   As Long = 32                 'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38               'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                 'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                  'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                  'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                  'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                 'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                 'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                 'Thunk data index of the User-defined callback
'   parameter data index
Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEHOVER     As Long = &H2A1
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_THEMECHANGED   As Long = &H31A

Private z_ScMem             As Long                      'Thunk base address
Private z_Sc(64)            As Long                      'Thunk machine-code initialised here
Private z_Funk              As Collection                'hWnd/thunk-address collection

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Function CallWindowProcA Lib "user32" ( _
      ByVal lpPrevWndFunc As Long, _
      ByVal hWnd As Long, _
      ByVal Msg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "Kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "Kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'''Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
'''Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function LoadLibraryA Lib "Kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLongA Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "Kernel32" ( _
      ByVal lpAddress As Long, _
      ByVal dwSize As Long, _
      ByVal flAllocationType As Long, _
      ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "Kernel32" ( _
      ByVal lpAddress As Long, _
      ByVal dwSize As Long, _
      ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-End of Selfsub-class declarations----------------------------------------------------------------------------

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize          As Long
   dwFlags         As TRACKMOUSEEVENT_FLAGS
   hwndTrack       As Long
   dwHoverTime     As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" _
      Alias "_TrackMouseEvent" ( _
      ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

'// MLH - ExportGrid (used to find the desktop folder and open CSV) -----------
Private Type typSHITEMID
   cb    As Long
   abID  As Byte
End Type

Private Type typITEMIDLIST
   mkid  As typSHITEMID
End Type

Private Type typSHELLEXECUTEINFO
   cbSize       As Long
   fMask        As Long
   hWnd         As Long
   lpVerb       As String
   lpFile       As String
   lpParameters As String
   lpDirectory  As String
   nShow        As Long
   hInstApp     As Long
   lpIDList     As Long
   lpClass      As String
   hkeyClass    As Long
   dwHotKey     As Long
   hIcon        As Long
   hProcess     As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef SEI As typSHELLEXECUTEINFO) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" ( _
      ByVal hWndOwner As Long, _
      ByVal nFolder As Long, _
      ByRef Pidl As typITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
      Alias "SHGetPathFromIDListA" ( _
      ByVal Pidl As Long, _
      ByVal pszPath As String) As Long

'API Type Declarations ------------------------------------------------------
Private Declare Function GetLocaleInfo Lib "Kernel32" _
      Alias "GetLocaleInfoA" ( _
      ByVal Locale As Long, _
      ByVal LCType As Long, _
      ByVal lpLCData As String, _
      ByVal cchData As Long) As Long

Private Type OSVersionInfo
   dwOSVersionInfoSize As Long
   dwMajorVersion      As Long
   dwMinorVersion      As Long
   dwBuildNumber       As Long
   dwPlatformId        As Long
   szCSDVersion        As String * 128 'Maintenance string for PSS usage
End Type

Private Type Rect
   Left   As Long
   Top    As Long
   Right  As Long
   Bottom As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type TRIVERTEX
   X     As Long
   Y     As Long
   Red   As Integer
   Green As Integer
   Blue  As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UpperLeft  As Long
   LowerRight As Long
End Type

'API Declarations
'Private Declare Function GetCursorPos Lib "user32" ( ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetWindowLongW Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandleW Lib "Kernel32" (ByVal lpModuleName As Long) As Long

Private Declare Function IsCharAlphaNumeric Lib "user32" _
      Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "Kernel32" _
      Alias "GetVersionExA" ( _
      ByRef lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetParent Lib "user32" ( _
      ByVal hWndChild As Long, _
      ByVal hWndNewParent As Long) As Long
      
Private Declare Function SendMessageAsLong Lib "user32" _
      Alias "SendMessageA" ( _
      ByVal hWnd As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetRect Lib "user32" ( _
      ByRef lpRect As Rect, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function SetRectRgn Lib "gdi32" ( _
      ByVal hRgn As Long, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Declare Function DrawTextA Lib "user32" ( _
      ByVal hdc As Long, _
      ByVal lpStr As String, _
      ByVal nCount As Long, _
      ByRef lpRect As Rect, _
      ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" ( _
      ByVal hdc As Long, _
      ByVal lpStr As Long, _
      ByVal nCount As Long, _
      ByRef lpRect As Rect, _
      ByVal wFormat As Long) As Long
Private Declare Function MoveTo Lib "gdi32" _
      Alias "MoveToEx" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByRef lpPoint As Any) As Long
Private Declare Function MoveToEx Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByRef lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" ( _
      ByVal nPenStyle As Long, _
      ByVal nWidth As Long, _
      ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" ( _
      ByVal OLE_COLOR As Long, _
      ByVal hPalette As Long, _
      ByRef pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function DrawFrameControl Lib "user32" ( _
      ByVal hdc As Long, _
      ByRef lpRect As Rect, _
      ByVal un1 As Long, _
      ByVal un2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" _
      Alias "GradientFill" ( _
      ByVal hdc As Long, _
      ByRef pVertex As TRIVERTEX, _
      ByVal dwNumVertex As Long, _
      ByRef pMesh As GRADIENT_RECT, _
      ByVal dwNumMesh As Long, _
      ByVal dwMode As Long) As Long '// MLH - new
'Private Declare Function GradientFill Lib "msimg32" ( ByVal hDC As Long, ByRef pVertex As Any, ByVal dwNumVertex As Long,
'   ByRef pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long '// MLH replaced with GradientFillRect
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function RoundRect Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal Left As Long, _
      ByVal Top As Long, _
      ByVal Right As Long, _
      ByVal Bottom As Long, _
      ByVal EllipseWidth As Long, _
      ByVal EllipseHeight As Long) As Long '// MLH - for Heavy Focus Rect
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long '// MLH - new

'---------------------------------------------------------------------------------------------------
'XP Theme
Public Enum lgThemeConst '// MLH
   Blue = 0
   silver = 1
   Olive = 2
   Storm = 3
   Earth = 4
   CustomTheme = 5
   Autodetect = 6
End Enum

Private muThemeColor       As lgThemeConst
Private mlngCustomColorFrom  As Long
Private mlngCustomColorTo    As Long
Private mstrCurSysThemeName  As String

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" ( _
      ByVal mhTheme As Long, _
      ByVal lHDC As Long, _
      ByVal iPartId As Long, _
      ByVal iStateId As Long, _
      ByRef pRect As Rect, _
      ByRef pClipRect As Rect) As Long

'''Private Declare Function DrawThemeEdge Lib "uxtheme.dll" ( ByVal mhTheme As Long, ByVal hDC As Long, ByVal iPartId As
'   Long, ByVal iStateId As Long, ByRef pDestRect As Rect, ByVal uEdge As Long, ByVal uFlags As Long, ByRef pContentRect As
'   Rect) As Long  '// MLH - Not Used

Private Declare Function OpenThemeData Lib "uxtheme.dll" ( _
      ByVal hWnd As Long, _
      ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal mhTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
      ByVal pszThemeFileName As Long, _
      ByVal dwMaxNameChars As Long, _
      ByVal pszColorBuff As Long, _
      ByVal cchMaxColorChars As Long, _
      ByVal pszSizeBuff As Long, _
      ByVal cchMaxSizeChars As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Private Const CLR_INVALID = &HFFFF

Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_SHOWDROPDOWN = &H14F
'''Private Const CB_GETDROPPEDSTATE = &H157 '// MLH - Not Used

Private Const GWL_EXSTYLE      As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&

'// DrawText- Unicode support
Private Const DT_BOTTOM        As Long = &H8
Private Const DT_CENTER        As Long = &H1
Private Const DT_LEFT          As Long = &H0
Private Const DT_RIGHT         As Long = &H2
Private Const DT_TOP           As Long = &H0
Private Const DT_VCENTER       As Long = &H4
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_SINGLELINE    As Long = &H20
Private Const DT_WORDBREAK     As Long = &H10
Private Const DT_CALCRECT      As Long = &H400
Private Const DT_NOPREFIX      As Long = &H800 '// MLH - New

'''Private Const BDR_RAISEDINNER = &H4 '// MLH - Not Used
'''Private Const BDR_RAISEDOUTER = &H1 '// MLH - Not Used
'''Private Const BDR_SUNKENINNER = &H8 '// MLH - Not Used
'''Private Const BDR_SUNKENOUTER = &H2 '// MLH - Not Used
'''Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)'// MLH - Not Used
'''Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)'// MLH - Not Used

'''Private Const BF_BOTTOM = &H8       '// MLH - Not Used
'''Private Const BF_LEFT = &H1         '// MLH - Not Used
'''Private Const BF_RIGHT = &H4        '// MLH - Not Used
'''Private Const BF_TOP = &H2          '// MLH - Not Used
'''Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM) '// MLH - Not Used

Private Const DFC_BUTTON        As Long = &H4
Private Const DFCS_FLAT         As Long = &H4000
Private Const DFCS_BUTTONCHECK  As Long = &H0
Private Const DFCS_BUTTONPUSH   As Long = &H10
Private Const DFCS_CHECKED      As Long = &H400
Private Const DFCS_PUSHED       As Long = &H200
'''Private Const DFCS_TRANSPARENT  As Long = &H800 ' Win98/2000 only '// MLH - Not Used
Private Const DFCS_HOT          As Long = &H1000

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const GRADIENT_FILL_RECT_H    As Long = &H0
Private Const GRADIENT_FILL_RECT_V    As Long = &H1
'''Private Const GRADIENT_FILL_TRIANGLE  As Long = &H2 '// MLH - Not Used
'''Private GRADIENT_FILL_RECT_DIRECTION  As Long '// MLH - Not Used

Private Const GWL_STYLE = (-16)
Private Const ES_UPPERCASE As Long = &H8&
Private Const ES_LOWERCASE As Long = &H10&

'-----------------------------------------------------------------------------------------------------
'API Scroll Bars

Private Type SCROLLINFO
   cbSize    As Long
   fMask     As Long
   nMin      As Long
   nMax      As Long
   nPage     As Long
   nPos      As Long
   nTrackPos As Long
End Type

Private Declare Function InitialiseFlatSB Lib "comctl32.dll" _
      Alias "InitializeFlatSB" ( _
      ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal n As Long, _
      ByRef lpcScrollInfo As SCROLLINFO, _
      ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal n As Long, _
      ByRef LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal wSBflags As Long, _
      ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal wBar As Long, _
      ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" ( _
      ByVal hWnd As Long, _
      ByVal int2 As Long, _
      ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" ( _
      ByVal hWnd As Long, _
      ByVal Code As Long, _
      ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" ( _
      ByVal hWnd As Long, _
      ByVal Code As Long, _
      ByRef LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" ( _
      ByVal hWnd As Long, _
      ByVal Code As Long, _
      ByRef LPSCROLLINFO As SCROLLINFO, _
      ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" ( _
      ByVal hWnd As Long, _
      ByVal Index As Long, _
      ByVal vNewValue As Long, _
      ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hWnd As Long) As Long

Public Enum ScrollBarOrienationEnum
   Scroll_Horizontal
   Scroll_Vertical
   Scroll_Both
   Scroll_None
End Enum

Public Enum ScrollBarStyleEnum
   Style_Regular = 1& 'FSB_REGULAR_MODE
   Style_Flat = 0&    'FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
   efsHorizontal = 0 'SB_HORZ
   efsVertical = 1   'SB_VERT
End Enum

'''Private Const SB_HORZ = 0 '// MLH - Not Used
'''Private Const SB_VERT = 1 '// MLH - Not Used
Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

'// MLH - the following are Not Used
'''Private Const ESB_DISABLE_BOTH = &H3
'''Private Const ESB_ENABLE_BOTH = &H0
'''Private Const FSB_FLAT_MODE = 1&
'''Private Const FSB_REGULAR_MODE = 0&
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&

'Private mbSBInitialised      As Boolean '// MLH - Not Used
Private muSBOrienation       As ScrollBarOrienationEnum
Private muSBStyle            As ScrollBarStyleEnum

'// MLH - Changed mSBSmallChangeHorz to Const
'''Private mSBSmallChangeHorz   As Long
'''Private mSBSmallChangeVert   As Long
'''Private Const C_SBSmallChangeHorz   As Long = 1
'''Private Const C_SBSmallChangeVert   As Long = 1
'''Private m_bEnabledHorz      As Boolean '// MLH - Not Used
'''Private m_bEnabledVert      As Boolean '// MLH - Not Used
Private mbSBVisibleHorz       As Boolean
Private mbSBVisibleVert       As Boolean
Private mbSBNoFlatScrollBars  As Boolean
Private mSBhWnd               As Long

'-----------------------------------------------------------------------------------
'Private Enum Statements

Private Enum lgFlagsEnum
   lgFLChecked = 2
   lgFLSelected = 4
   lgFLChanged = 8
   lgFLFontBold = 16
   lgFLFontItalic = 32
   lgFLFontUnderline = 64
   lgFLWordWrap = 128
   lgFLNewRow = 256
   lgFLlocked = 512
End Enum

Private Enum lgHeaderStateEnum
   lgNormal = 1
   lghot = 2
   lgDOWN = 3
End Enum

Private Enum lgRectTypeEnum
   lgRTColumn = 0
   lgRTCheckBox = 1
   lgRTImage = 2
End Enum

'// Public Enum Statements

'// MLH - this is for columns not rows
'''Public Enum lgAllowColumnResizingEnum
'''   lgResizeNone = 0
'''   lgResizeCol = 1
'''   lgResizeRow = 2
'''   lgResizeBoth = 4
'''End Enum

Public Enum lgMultiSelectEnum '// MLH - new
   lgSingleSelect = 0
   lgMultiStandard = 1
   lgMultiLatch = 2
End Enum

Public Enum lgCaptionAlignmentEnum '// MLH - new
   lgAlignLeft = DT_LEFT Or DT_VCENTER
   lgAlignCenter = DT_CENTER Or DT_VCENTER
   lgAlignRight = DT_RIGHT Or DT_VCENTER
End Enum

Public Enum lgAlignmentEnum
   lgAlignLeftTop = DT_LEFT Or DT_TOP
   lgAlignLeftCenter = DT_LEFT Or DT_VCENTER
   lgAlignLeftBottom = DT_LEFT Or DT_BOTTOM
   lgAlignCenterTop = DT_CENTER Or DT_TOP
   lgAlignCenterCenter = DT_CENTER Or DT_VCENTER
   lgAlignCenterBottom = DT_CENTER Or DT_BOTTOM
   lgAlignRightTop = DT_RIGHT Or DT_TOP
   lgAlignRightCenter = DT_RIGHT Or DT_VCENTER
   lgAlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

Public Enum lgBorderStyleEnum
   lgNone = 0
   lgSingle = 1
End Enum

Public Enum lgCellFormatEnum
   lgCFBackColor = 1
   lgCFForeColor = 2
   lgCFImage = 2
   lgCFFontName = 3
   lgCFFontBold = 4
   lgCFFontItalic = 5
   lgCFFontUnderline = 6
End Enum

Public Enum lgDataTypeEnum
   lgString = 0
   lgNumeric = 1
   lgDate = 2
   lgBoolean = 3
   lgProgressBar = 4
   lgCustom = 5
   lgButton = 6
End Enum

Public Enum lgEditTriggerEnum
   lgNone = 0
   lgEnterKey = 2
   lgF2Key = 4
   lgMouseClick = 8
   lgMouseDblClick = 16
   lgAnyKey = 32
   lgAnyF2DblCk = 52 '// MLH - new
End Enum

Public Enum lgFocusRectModeEnum
   lgNone = 0
   lgRow = 1
   lgCol = 2
End Enum

Public Enum lgFocusRectStyleEnum
   lgFRLight = 0
   lgFRHeavy = 1
   lgFRMedium = 2 '// MLH - new
End Enum

Public Enum lgFocusRowHighlightStyle '// MLH - new
   [Solid] = 0
   [Gradient_V]
   [Gradient_H]
End Enum

Public Enum lgMoveControlEnum
   lgBCNone = 0
   lgBCHeight = 1
   lgBCWidth = 2
   lgBCLeft = 4
   lgBCTop = 8
End Enum

Public Enum lgSearchModeEnum
   lgSMEqual = 0
   lgSMGreaterEqual = 1
   lgSMLike = 2
   lgSMNavigate = 4
   lgSMBeginsWith = 5
End Enum

Public Enum lgSortTypeEnum
   lgSTAscending = 1    '// MLH - Changed
   lgSTDescending = 2   '// MLH - Changed
   lgSTNormal = 0       '// MLH - New
End Enum

Public Enum lgThemeColorEnum
   lgTCCustom = 0
   lgTCDefault = 1
   lgTCBlue = 2
   lgTCGreen = 3
End Enum

Public Enum lgThemeStyleEnum
   lgTSWindows3D = 0
   lgTSWindowsFlat = 1
   lgTSWindowsXP = 2
   lgTSOfficeXP = 3
End Enum

'// Used for the "Appearance" property
Public Enum lgAppearanceEnum
   Appear_Flat = 0
   Appear_3D = 1
End Enum

Public Enum lgGridLinesEnum
   lgGrid_None
   lgGrid_Both
   lgGrid_Vertical
   lgGrid_Horizontal
End Enum

Public Enum lgEditMoveEnum '// MLH - New
   lgDontNone = 0
   lgMoveRight = 1
   lgMoveDown = 2
End Enum

'// User Defined Types

Private Type udtColumn
   EditCtrl          As Object
   dCustomWidth      As Single
   lWidth            As Long
   lX                As Long
   nAlignment        As lgAlignmentEnum
   nImageAlignment   As lgAlignmentEnum
   nSortOrder        As lgSortTypeEnum
   nType             As Integer
   nFlags            As Integer
   MoveControl       As Integer
   bVisible          As Boolean
   sCaption          As String
   sFormat           As String
   sInputFilter      As String
   sTag              As String
   bLocked           As Boolean '// MLH - new
End Type

Private Type udtCell
   nAlignment  As Integer
   nFormat     As Integer
   nFlags      As Integer
   sValue      As String
   lImage      As Integer     '// MLH Cell Image (for all columns but 0)
   sTag        As String      '// MLH Cell Tag
   pPic        As StdPicture  '// MLH - new
End Type

Private Type udtItem
   lHeight     As Long
   lImage      As Long     '// Row Image (displayed in column 0 only)
   lItemData   As Long
   nFlags      As Integer
   sTag        As String
   bGroupRow   As Boolean  '// MLH - New
   bVisible    As Boolean  '// MLH - New
   Cell()      As udtCell
End Type

Private Type udtFormat
   lBackColor  As Long
   lForeColor  As Long
   'nImage      As Integer '// MLH - moved to Type udtCell - did not work here
   sFontName   As String
   lRefCount   As Long
End Type

Private Type udtRender
   DTFlag         As Long
   CheckBoxSize   As Long
   ImageSpace     As Boolean  '// MLH - changed was long - alignment handled in Drawgrid
   ImageHeight    As Long
   ImageWidth     As Long
   LeftImage      As Long
   LeftText       As Long
   HeaderHeight   As Long
   TextHeight     As Long
   CaptionHeight  As Long     '// MLH - new
End Type

Private Type typTotals
   'dValue   As Double
   bAvg     As Boolean
   sCaption As String
End Type

'// User Defined Settings
'// MLH removed all setting constants with VB constants
Private Const DEF_BORDERSTYLE           As Long = lgBorderStyleEnum.lgSingle
Private Const DEF_CACHEINCREMENT        As Long = 10
Private Const DEF_FOCUSRECTMODE         As Long = lgFocusRectModeEnum.lgNone '// MLH - Changed Default
Private Const DEF_FOCUSRECTSTYLE        As Long = lgFocusRectStyleEnum.lgFRHeavy
Private Const DEF_GRIDCOLOR             As Long = &HC0C0C0
Private Const DEF_GRIDLINEWIDTH         As Long = 1
Private Const DEF_PROGRESSBARCOLOR      As Long = &H8080FF
Private Const DEF_MaxLineCount          As Long = 0
Private Const DEF_MinRowHeight          As Long = 0 'DEF_ROWHEIGHT * 15
Private Const DEF_SEARCHCOLUMN          As Long = 0
Private Const DEF_THEMESTYLE            As Long = lgThemeStyleEnum.lgTSWindowsXP
Private Const DEF_MinVerticalOffset     As Long = 2 '// MLH - new
'// MLH - the following are Not Used
'''Private Const DEF_ROWHEIGHT             As Long = 18
'''Private Const DEF_AllowColumnResizing   As Long = lgAllowColumnResizingEnum.lgResizeBoth
'''Private Const DEF_FORECOLORTOTALS       As Long = vbRed
'''Private Const DEF_FORMATSTRING          As String = vbNullString

Private Const C_NULL_RESULT         As Long = -1
Private Const C_AUTOSCROLL_TIMEOUT  As Long = 25
Private Const C_SIZE_VARIANCE       As Long = 4
Private Const C_SCROLL_NONE         As Long = 0
Private Const C_SCROLL_UP           As Long = 1
Private Const C_SCROLL_DOWN         As Long = 2
'// For Rendering
Private Const C_MAX_CHECKBOXSIZE    As Long = 16
Private Const C_SIZE_SORTARROW      As Long = 8
Private Const C_TEXT_SPACE          As Long = 3
Private Const C_ARROW_SPACE         As Long = 5
Private Const C_RIGHT_CHECKBOX      As Long = 15
Private Const C_CHECKTEXT           As String = "ABCDWXYZ"

'// VB Controls
Private WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1
Private picTooltip         As PictureBox

'// Data & Columns arrays
Private mCols()         As udtColumn
Private mItems()        As udtItem
Private mColPtr()       As Long
Private mudtTotals()    As typTotals   '// MLH - New
Private mudtTotalsVal() As Double      '// MLH - New
'Private mColReal()      As Long '// MLH - not used
Private mRowPtr()       As Long
Private mCF()           As udtFormat

Private mbTotalsLineShow As Boolean
Private mRowCount        As Long
'Private mRowsVisible    As Long
Private mSortColumn      As Long
Private mSortSubColumn   As Long
Private mSwapCol         As Long
Private mDragCol         As Long
Private mResizeCol       As Long
Private mResizeRow       As Long '// MLH - New
Private mbIgnoreMove     As Boolean

Private mEditCol        As Long
Private mEditRow        As Long
Private mbEditPending   As Boolean
Private mEditParent     As Long
Private muEditMove      As lgEditMoveEnum '// MLH - New

Private mCol            As Long
Private mRow            As Long
Private mSelectedRow    As Long
Private mLRLocked       As Boolean
Private mLCLocked       As Boolean

Private mMouseCol       As Long
Private mMouseRow       As Long
Private mMouseDownCol   As Long
Private mMouseDownRow   As Long
Private mMouseDownX     As Long
Private mbMouseDown     As Boolean

Private mR              As udtRender

'------------------------------------------------------------------------
'// Appearance Properties
Private mbApplySelectionToImages  As Boolean
Private mBackColor                As Long
Private mBackColorBkg             As Long
Private mBackColorEdit            As Long
'''Private mBackColorFixed           As Long'// MLH - Not Needed
Private mBackColorSel             As Long
Private mForeColor                As Long
Private mForeColorEdit            As Long
'''Private mForeColorFixed           As Long'// MLH - Not Used
Private mForeColorHdr             As Long
Private mForeColorSel             As Long
Private mblnColumnHeaderSmall     As Boolean
'''Private mForeColorTotals         As Long  '// MLH - Not Used
Private mbBackColorEvenRowsE      As Boolean '// MLH - New
Private mBackColorEvenRows        As Long    '// MLH - New
Private mFocusRectColor           As Long
Private mGridColor                As Long
Private mProgressBarColor         As Long
Private mbAlphaBlendSelection     As Boolean
Private muBorderStyle             As lgBorderStyleEnum
Private mbDisplayEllipsis         As Boolean
Private muFocusRectMode           As lgFocusRectModeEnum
Private muFocusRectStyle          As lgFocusRectStyleEnum
Private mFont                     As Font
Private muGridLines               As lgGridLinesEnum
Private mGridLineWidth            As Long
Private muThemeStyle              As lgThemeStyleEnum '// MLH - changed
Private mbColumnHeaders           As Boolean          '// MLH - new
Private mbCenterRowImage          As Boolean          '// MLH - new
Private mColumnHeaderLines        As Integer          '// MLH - new
Private msCaption                 As String           '// MLH - new
Private muCaptionAlignment        As lgCaptionAlignmentEnum '// MLH - new
Private muScrollBarStyle          As ScrollBarStyleEnum '// MLH - New
Private mblnKeepForeColor         As Boolean '// MLH - New

Private ucFontBold                As Boolean
Private ucFontItalic              As Boolean
Private ucFontName                As String

'------------------------------------------------------------------------
'// Behaviour Properties
Private mbAllowColumnResizing    As Boolean '// MLH
Private mbAllowRowResizing       As Boolean '// MLH - new
Private mbAllowWordWrap          As Boolean '// MLH - new
Private mbAllowDelete            As Boolean '// MLH - new
Private mbAllowInsert            As Boolean '// MLH - new
Private mbCheckboxes             As Boolean
Private mbAllowColumnSwap        As Boolean '// MLH - renamed
Private mbAllowColumnDrag        As Boolean '// MLH - new
Private mbAllowColumnSort        As Boolean
Private mbAllowEdit              As Boolean
Private muEditTrigger            As lgEditTriggerEnum
Private mbFullRowSelect          As Boolean
Private muFocusRowHighlightStyle As lgFocusRowHighlightStyle '// MLH - new
Private mbHideSelection          As Boolean
Private mbAllowColumnHover       As Boolean
'Private mbMultiSelect            As Boolean
Private muMultiSelect            As lgMultiSelectEnum '// MLH - Added New Options
Private mbRedraw                 As Boolean
Private mbUserRedraw             As Boolean '// MLH - Changed
Private mbScrollTrack            As Boolean
Private mbAutoToolTips           As Boolean
Private mlngFreezeAtCol          As Long '// MLH - New

'------------------------------------------------------------------------
'// Miscellaneous Properties
Private mCacheIncrement      As Long
Private mbEnabled            As Boolean
Private mExpandRowImage      As Integer
'''Private mFormatString       As String '// MLH - Not Used
'''Private mLocked             As Boolean'// MLH - Not Used
Private mMaxLineCount        As Long

Private mMinRowHeightUser    As Long      '// MLH - new
Private mMinRowHeight        As Long
Private mMinVerticalOffset   As Long      '// MLH - new

Private muScaleUnits         As ScaleModeConstants
Private mSearchColumn        As Long

Private moImageList          As Object
Private mImageListScaleMode  As Integer
Private mbCellsChanged       As Boolean '// MLH - Have any cells been edited
Private miKeyCode            As Integer '// MLH - Used in Drawgrid (when workwrap is on) and UserControl_KeyDown for Edit mode.

'------------------------------------------------------------------------
'// Control State Variables
Private mbInCtrl            As Boolean
'''Private mbInFocus           As Boolean '// MLH
Private mbWinNT             As Boolean
Private mbWinXP             As Boolean
Private mbLockFocusDraw     As Boolean

Private mbPendingRedraw     As Boolean
Private mbPendingScrollBar  As Boolean

Private mTextBoxStyle       As Long
Private mClipRgn            As Long
Private mhTheme             As Long
Private mScrollAction       As Long
Private mScrollTick         As Long
Private mHotColumn          As Long
Private mbIgnoreKeyPress    As Boolean

'------------------------------------------------------------------------
'// Static Variables - '// MLH - Moved statics here
'// Private sbDrawLocked     As Boolean '// SetRedrawState
Private mlTime           As Long     '// Usercontrol_Keypress
Private msCode           As String   '// Usercontrol_Keypress
Private mlResizeX        As Long     '// UserControl_MouseMove
Private mlResizeY        As Long     '// UserControl_MouseMove'// MLH - New
Private mlTickCount      As Long     '// ShowCompleteCell '// MLH - new
Private mbWorking        As Boolean  '// ShowCompleteCell '// MLH - new
Private mbCancelShow     As Boolean  '// ShowCompleteCell '// MLH - new

'------------------------------------------------------------------------
'// Events - Standard VB
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'// Events - Control Specific
Public Event CellButtonClick(ByVal Row As Long, ByVal Col As Long) '// MLH - New
Public Event CellImageClick(ByVal Row As Long, ByVal Col As Long)
Public Event ColumnClick(ByVal Col As Long)
Public Event ColumnSizeChanged(ByVal Col As Long, ByVal MoveControl As lgMoveControlEnum)
Public Event CustomSort(Ascending As Boolean, Col As Long, Value1 As String, Value2 As String, Swap As Boolean)
Public Event RowChecked(ByVal Row As Long)
Public Event RowCountChanged()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event RowColChanged()
Public Event Scroll()
Public Event SelectionChanged()
Public Event SortComplete()
Public Event ThemeChanged()
Public Event EditKeyPress(ByVal Col As Long, ByRef KeyAscii As Integer)
'''Public Event EnterCell()'// MLH - Not Used
Public Event BeforeEdit(ByVal Row As Long, ByVal Col As Long, ByRef Cancel As Boolean) '// MLH - Changed Name
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long, ByRef vNewValue As String, ByRef Cancel As Boolean) '// MLH - Changed Name
Public Event BeforeDelete(ByVal Row As Long, ByRef Cancel As Boolean) '// MLH - New
Public Event AfterDelete()                                            '// MLH - New
Public Event BeforeInsert(ByVal Row As Long, ByRef Cancel As Boolean) '// MLH - New
Public Event AfterInsert(ByVal Row As Long)                           '// MLH - New

'''''''// This code needs reference to ADO in the project
'''Public Sub FillGridFromQuery(ByRef rActiveRecordset As ADODB.Recordset, _
'''                             Optional ByVal bAddColumns As Boolean = False, _
'''                             Optional ByVal bFitColWidth As Boolean = False)
'''
'''  Dim fd      As ADODB.Field
'''  Dim BkMark  As Variant
'''  Dim lCol    As Long
'''  Dim lRow    As Long
'''  Dim lngI    As Long
'''  Dim strTemp As String
'''  Dim pPic    As StdPicture
'''
'''  '// MLH - New
'''
'''   If bAddColumns Then
'''      Me.ClearAll
'''   Else
'''      Me.Clear
'''   End If
'''
'''   If Not (rActiveRecordset.EOF And rActiveRecordset.BOF) Then '// Empty Query ?
'''
'''      BkMark = rActiveRecordset.Bookmark '// save current index
'''
'''      rActiveRecordset.MoveLast
'''      lRow = rActiveRecordset.RecordCount
'''
'''      If bAddColumns Then
'''         For Each fd In rActiveRecordset.Fields
'''             Me.AddColumn fd.Name
'''         Next
'''      End If
'''
'''      rActiveRecordset.MoveFirst
'''      For lngI = 1 To lRow
'''         strTemp = vbNullString
'''         For lCol = 0 To rActiveRecordset.Fields.Count - 1
'''             strTemp = strTemp & rActiveRecordset.Fields(lCol).Value & vbTab
'''         Next
'''
'''         Me.AddItem strTemp
'''
'''         rActiveRecordset.MoveNext
'''      Next lngI
'''
'''      rActiveRecordset.Bookmark = BkMark '// restore saved index
'''
'''      If bFitColWidth Or bAddColumns Then
'''         Call ColForceFit
'''      End If
'''
'''   End If
'''
'''End Sub

Public Function AddColumn(Optional ByVal Caption As String, _
                          Optional ByVal Width As Single, _
                          Optional ByVal Alignment As lgAlignmentEnum = lgAlignLeftCenter, _
                          Optional ByVal DataType As lgDataTypeEnum = lgString, _
                          Optional ByVal Format As String, _
                          Optional ByVal InputFilter As String, _
                          Optional ByVal ImageAlignment As lgAlignmentEnum = lgAlignLeftCenter, _
                          Optional ByVal WordWrap As Boolean = False, _
                          Optional ByVal Index As Long = 0, _
                          Optional ByVal bVisible As Boolean = True, _
                          Optional ByVal bLocked As Boolean = False) As Long

   '// MLH - Added Optional ByVal bVisible and bLocked
   '-----------------------------------------------------------------------------------
   '// Purpose: Add a Column to the Grid

   '// Caption        - The text that appears on the Header
   '// Width          - The Width!
   '// Alignment      - The Alignment!
   '// DataType       - Allows the control to determine proper Sort Sequence when Sorting
   '// Format         - Format Mask applied to Cell data before it is displayed (i.e. "#.00")
   '// InputFilter    - Characters allowed in TextBox entry
   '// ImageAlignment - Image Alignment!
   '// WordWrap       - Enable Word-Wrap for this col
   '// Index          - Allows a new Column to be Inserted before an existing one
   '// bVisible       - Make Col visible/invisible
   '// bLocked        - Prevent column from getting the focus

   '// mColPtr() is used as an Index to the Columns (a bit like an array of "pointers")
   '-----------------------------------------------------------------------------------

  Dim lCount  As Long
  Dim lNewCol As Long

   If Not (mCols(0).nAlignment = 0) Then
      lNewCol = UBound(mCols) + 1
      ReDim Preserve mCols(lNewCol) As udtColumn
      ReDim Preserve mColPtr(lNewCol) As Long
      ReDim Preserve mudtTotals(lNewCol) As typTotals    '// MLH - New
      ReDim Preserve mudtTotalsVal(lNewCol) As Double    '// MLH - New
      'ReDim Preserve mColReal(lNewCol) As Long'// MLH - not used
   End If

   If (Index > 0) And (Index < lNewCol) Then
      If lNewCol > 1 Then

         For lCount = lNewCol To Index + 1 Step -1
            mColPtr(lCount) = mColPtr(lCount - 1)
         Next lCount

         mColPtr(Index) = lNewCol
      End If

      AddColumn = Index

   Else
      mColPtr(lNewCol) = lNewCol
      AddColumn = lNewCol
   End If

   With mCols(lNewCol)
      .sCaption = Caption
      .dCustomWidth = Width

      '// lWidth is always Pixels (because thats what API functions require) and
      '// is calculated to prevent repeated Width Scaling calculations
      .lWidth = ScaleX(.dCustomWidth, muScaleUnits, vbPixels)

      .nAlignment = Alignment
      .nImageAlignment = ImageAlignment
      .nSortOrder = lgSTNormal '// MLH
      .nType = DataType
      .sFormat = Format

      If LenB(InputFilter) = 0 Then

         Select Case DataType
         Case lgNumeric
            .sInputFilter = "1234567890.-,"  '// MLH rVal fix for ,

         Case lgDate
            .sInputFilter = "1234567890./-"
         End Select

      Else
         .sInputFilter = InputFilter
      End If

      If WordWrap Then
         .nFlags = lgFLWordWrap
      End If

      .bVisible = bVisible '// MLH - new
      .bLocked = bLocked   '// MLH - new
   End With

   DisplayChange

End Function

Public Function AddItem(Optional ByVal vstrItem As String, _
                        Optional ByVal vRow As Long = C_NULL_RESULT, _
                        Optional ByVal vbRowChecked As Boolean = False, _
                        Optional ByVal vbMarkAsNew As Boolean = False, _
                        Optional ByVal vbRowLocked As Boolean = False, _
                        Optional ByVal vbRowVisible As Boolean = True, _
                        Optional ByVal vRowData As Long) As Long

   '---------------------------------------------------------------------------------
   '// Purpose: Add an vstrItem (new Row) to the Grid

   '// vstrItem     - This contains the data for the Cells in the new Row. You can pass multiple
   '//                Cells by using a Delimiter (vbTab) between Cell data
   '// vRow         - Allows a new Row to be Inserted before an existing one
   '// vbRowChecked - Default Checked state of the new row
   '// vbMarkAsNew  - Mark row as new/inserted by user
   '// vbRowLocked  - Mark row as Locked (Prevent row selection)
   '// vbRowVisible - Hide/Show Row

   '// mItems() is an array of the Items in the Grid
   '// mRowPtr() is used as an vRow to the Items (a bit like an array of "pointers")

   '// The Row technique is used to allow faster Inserts & Sorts since we only need to swap a Long (4 bytes)
   '// rather than a large data structure (a UDT in this case)

   '// The mItems() is resized incrementally to reduce the Redim Preserve overhead. The default mCacheIncrement
   '// is 10 but this can be increased to a higher value to increase performance if adding thousands of Items

   '// MLH - Changed Optional vRow As Long = 0, Added vbMarkAsNew, vbRowLocked, vbRowVisible
   '---------------------------------------------------------------------------------

  Dim lCol    As Long
  Dim lCount  As Long
  Dim sText() As String

   mRowCount = mRowCount + 1

   If mRowCount > UBound(mItems) Then
      ReDim Preserve mItems(mRowCount + mCacheIncrement) As udtItem
      ReDim Preserve mRowPtr(mRowCount + mCacheIncrement) As Long
   End If

   If (vRow >= 0) And (vRow < mRowCount) Then
      If mRowCount Then '// MLH
         For lCount = mRowCount To vRow + 1 Step -1
            mRowPtr(lCount) = mRowPtr(lCount - 1)
         Next lCount

         mRowPtr(vRow) = mRowCount
      End If

      AddItem = vRow

   Else
      mRowPtr(mRowCount) = mRowCount
      AddItem = mRowCount
   End If

   mItems(mRowCount).lItemData = vRowData '// MLH
   
   '// MLH - Is never =< 0
   'If mMinRowHeight > 0 Then
   mItems(mRowCount).lHeight = mMinRowHeight
   'Else
   '   mItems(mRowCount).lHeight = DEF_ROWHEIGHT
   'End If
   
   ReDim mItems(mRowCount).Cell(UBound(mCols))
   
   For lCount = LBound(mCols) To UBound(mCols)

      With mItems(mRowCount).Cell(lCount)
         .nAlignment = mCols(lCount).nAlignment
         .nFormat = C_NULL_RESULT
         .nFlags = mCols(lCount).nFlags
      End With

      ApplyCellFormat mRowCount, lCount, lgCFBackColor, mBackColor
      ApplyCellFormat mRowCount, lCount, lgCFForeColor, mForeColor
      ApplyCellFormat mRowCount, lCount, lgCFFontName, mFont.Name
   Next lCount

   If UBound(mCols) > 0 Then
      lCol = 0
      sText() = Split(vstrItem, vbTab)

      For lCount = LBound(sText) To UBound(sText)

         mItems(mRowCount).Cell(lCol).sValue = sText(lCount)

         If mCols(lCol).nType = lgBoolean Then '// MLH - If Boolean set CellChecked = True
            If LenB(sText(lCount)) Then '// MLH added to prevent error when a null string is passed
               SetFlag mItems(mRowPtr(mRowCount)).Cell(lCol).nFlags, lgFLChecked, CBool(sText(lCount))
            End If
         End If
         
         mudtTotalsVal(lCol) = mudtTotalsVal(lCol) + rVal(sText(lCount)) '// MLH - New

         lCol = lCol + 1

         If lCol > UBound(mCols) Then
            Exit For
         End If

      Next lCount

   Else '// adding Item data with only 1 column
      mItems(mRowCount).Cell(0).sValue = vstrItem
      If mCols(0).nType = lgBoolean Then '// MLH - If Boolean set CellChecked = True
         If LenB(sText(lCount)) Then '// MLH added to prevent error when a null string is passed
            SetFlag mItems(mRowPtr(mRowCount)).Cell(0).nFlags, lgFLChecked, CBool(vstrItem)
         End If
      End If
      
      mudtTotalsVal(lCol) = mudtTotalsVal(lCol) + rVal(vstrItem) '// MLH - New

   End If

   If vbRowChecked Then
      SetFlag mItems(mRowCount).nFlags, lgFLChecked, True '// Set RowChecked = True
   End If

   If vbRowLocked Then
      SetFlag mItems(mRowCount).nFlags, lgFLlocked, True  '// Set RowLocked = True
   End If

   If vbMarkAsNew Then '// MLH - New row inserted
      SetFlag mItems(mRowCount).nFlags, lgFLNewRow, True
   End If
   
   mItems(mRowCount).bVisible = vbRowVisible '// MLH - new

   If mbAllowWordWrap Then SetRowSize (mRowCount) '// MLH
   DisplayChange

   RaiseEvent RowCountChanged

End Function

Public Property Get AllowColumnDrag() As Boolean
Attribute AllowColumnDrag.VB_Description = "Returns/sets a value that determines whether the user is allow to change the Column order by Dragging"

   '// MLH - new
   AllowColumnDrag = mbAllowColumnDrag

End Property

Public Property Let AllowColumnDrag(ByVal vNewValue As Boolean)

   mbAllowColumnDrag = vNewValue
   If vNewValue Then mbAllowColumnSwap = False
   PropertyChanged "ColumnDrag"

End Property

Public Property Get AllowColumnHover() As Boolean
Attribute AllowColumnHover.VB_Description = "Returns/sets a value that determines whether the column is Highlighted when the mouse moves over it"

   AllowColumnHover = mbAllowColumnHover

End Property

Public Property Let AllowColumnHover(ByVal vNewValue As Boolean)

   '// MLH - renamed
   mbAllowColumnHover = vNewValue
   PropertyChanged "HotHeaderTracking"

   If Not vNewValue Then
      Call DrawHeaderRow
   End If

End Property

Public Property Get AllowColumnResizing() As Boolean
Attribute AllowColumnResizing.VB_Description = "Returns/sets a value that determines whether the user is allow to resize columns"

   '// MLH - new
   AllowColumnResizing = mbAllowColumnResizing

End Property

Public Property Let AllowColumnResizing(ByVal vNewValue As Boolean)

   mbAllowColumnResizing = vNewValue
   PropertyChanged "AllowColumnResizing"

End Property

Public Property Get AllowColumnSort() As Boolean
Attribute AllowColumnSort.VB_Description = "Returns/sets a value that determines whether the user is allow to sort columns"

   '// MLH - Changed Name
   AllowColumnSort = mbAllowColumnSort

End Property

Public Property Let AllowColumnSort(ByVal vNewValue As Boolean)

   mbAllowColumnSort = vNewValue
   PropertyChanged "ColumnSort"

End Property

Public Property Get AllowColumnSwap() As Boolean
Attribute AllowColumnSwap.VB_Description = "Returns/sets a value that determines whether the user is allow to change the Column order by swapping two columns"

   '// MLH - Changed Name
   AllowColumnSwap = mbAllowColumnSwap

End Property

Public Property Let AllowColumnSwap(ByVal vNewValue As Boolean)

   mbAllowColumnSwap = vNewValue
   If vNewValue Then mbAllowColumnDrag = False
   PropertyChanged "ColumnSwap"

End Property

Public Property Get AllowDelete() As Boolean
Attribute AllowDelete.VB_Description = "Returns/sets a value that determines whether the user is allow to delete cell info or the entire row"

   '// MLH
   AllowDelete = mbAllowDelete

End Property

Public Property Let AllowDelete(ByVal vNewValue As Boolean)

   '// MLH - new sub
   mbAllowDelete = vNewValue
   PropertyChanged "AllowDelete"

End Property

Public Property Get AllowEdit() As Boolean
Attribute AllowEdit.VB_Description = "Returns/sets a value that determines whether the user is allow to edit a cell"

   '// MLH - Changed Name
   AllowEdit = mbAllowEdit

End Property

Public Property Let AllowEdit(ByVal vNewValue As Boolean)

   mbAllowEdit = vNewValue
   PropertyChanged "Editable"

   If vNewValue Then '// MLH
      If muFocusRectMode = lgFocusRectModeEnum.lgNone Then
         muFocusRectMode = lgFocusRectModeEnum.lgCol
      End If
   End If

End Property

Public Property Get AllowInsert() As Boolean
Attribute AllowInsert.VB_Description = "Returns/sets a value that determines whether the user is allow to insert a row"

   AllowInsert = mbAllowInsert

End Property

Public Property Let AllowInsert(ByVal vNewValue As Boolean)

   '// MLH - new sub
   mbAllowInsert = vNewValue
   PropertyChanged "AllowInsert"

End Property

Public Property Get AllowRowResizing() As Boolean
Attribute AllowRowResizing.VB_Description = "Returns/sets a value that determines whether the user is allow to resize rows"

   AllowRowResizing = mbAllowRowResizing

End Property

Public Property Let AllowRowResizing(ByVal vNewValue As Boolean)

   mbAllowRowResizing = vNewValue
   PropertyChanged "AllowRowResizing"

End Property

Public Property Get AllowWordWrap() As Boolean
Attribute AllowWordWrap.VB_Description = "Returns/sets a value that determines whether truncated cells (that have CellWordWrap set to True) are word wrap."

   AllowWordWrap = mbAllowWordWrap

End Property

Public Property Let AllowWordWrap(ByVal vNewValue As Boolean)

   '// MLH - new sub
   mbAllowWordWrap = vNewValue
   DrawGrid mbRedraw
   PropertyChanged "AllowWordWrap"

End Property

Public Property Get AlphaBlendSelection() As Boolean
Attribute AlphaBlendSelection.VB_Description = "Returns/sets a value that determines whether the full row focus bar is Soften"

   AlphaBlendSelection = mbAlphaBlendSelection

End Property

Public Property Let AlphaBlendSelection(ByVal vNewValue As Boolean)

   mbAlphaBlendSelection = vNewValue
   DisplayChange
   PropertyChanged "AlphaBlendSelection"

End Property

Public Property Get Appearance() As lgAppearanceEnum
Attribute Appearance.VB_Description = "Returns/sets a value that determines appearance of the grid (Flat or 3D)"

   Appearance = UserControl.Appearance

End Property

Public Property Let Appearance(ByVal udtValue As lgAppearanceEnum)

   UserControl.Appearance = udtValue
   PropertyChanged "Appearance"

End Property

Private Sub ApplyCellFormat(ByVal vRow As Long, _
                            ByVal vCol As Long, _
                            ByVal Apply As lgCellFormatEnum, _
                            ByVal vNewValue As Variant)

   '---------------------------------------------------------------------------------
   '// Purpose: Apply formatting to a Cell. Attempts to find a matching entry in the
   '// Format Table and creates a new entry if a match is not found.

   '// In any "normal" use the grid will only have a few specifically formatted cells
   '// (such as Red forecolor in a financial column to indicate negative). It is therefore
   '// wasteful for each cell to store these properties. This system significantly reduces
   '// the memory used by the cells in a large Grid at the cost of slightly reduced perfomance.

   '// The Format element is an Integer allowing 32767 combinations. It could be a
   '// long for more combinations - however the aim is to keep the Cell UDT as small as possible!
   '---------------------------------------------------------------------------------

  Dim lBackColor  As Long
  Dim lForeColor  As Long
   '''Dim nImage      As Integer '// MLH
  Dim sFontName   As String
  Dim lCount      As Long
  Dim nIndex      As Integer
  Dim nFreeIndex  As Integer
  Dim nNewIndex   As Integer
  Dim bMatch      As Boolean

   If mRowCount = C_NULL_RESULT Then '// MLH - prevent errors

      If IsInIDE Then
         MsgBox "No Rows Added", vbExclamation, "ERROR - ApplyCellFormat"
      End If

      Exit Sub
   End If

   nIndex = mItems(vRow).Cell(vCol).nFormat

   If nIndex >= 0 Then
      '// Get current properties
      With mCF(nIndex)
         lBackColor = .lBackColor
         lForeColor = .lForeColor
         'nImage = .nImage '// MLH
         sFontName = .sFontName
      End With

   Else
      '// Set default properties
      lBackColor = mBackColor
      lForeColor = mForeColor
      sFontName = mFont.Name
   End If

   Select Case Apply
   Case lgCFBackColor
      lBackColor = vNewValue

   Case lgCFForeColor
      lForeColor = vNewValue

   'Case lgCFImage '// MLH
   '   nImage = vNewValue

   Case lgCFFontName
      sFontName = vNewValue

   End Select

   '// Search Format Table for matching entry
   nFreeIndex = C_NULL_RESULT

   For lCount = 0 To UBound(mCF)

      'If (mCF(lCount).lBackColor = lBackColor) And (mCF(lCount).lForeColor = lForeColor) And (mCF(lCount).nImage = nImage) And (mCF(lCount).sFontName = sFontName) Then '// MLH
      If (mCF(lCount).lBackColor = lBackColor) And (mCF(lCount).lForeColor = lForeColor) And (mCF(lCount).sFontName = sFontName) Then

         '// Existing Entry matches what we required
         bMatch = True
         nNewIndex = lCount
         Exit For

      ElseIf (mCF(lCount).lRefCount = 0) And (nFreeIndex = C_NULL_RESULT) Then
         '// An unused entry
         nFreeIndex = lCount
      End If

   Next lCount

   '// No existing matches
   If Not bMatch Then
      '// Is there an unused Entry?
      If nFreeIndex >= 0 Then
         nNewIndex = nFreeIndex
      Else
         nNewIndex = UBound(mCF) + 1
         ReDim Preserve mCF(nNewIndex + 9) As udtFormat
      End If

      With mCF(nNewIndex)
         .lBackColor = lBackColor
         .lForeColor = lForeColor
         '.nImage = nImage'// MLH
         .sFontName = sFontName
      End With

   End If

   '// Has the Format Entry Index changed?
   If Not (nIndex = nNewIndex) Then
      '// Increment reference count for new entry
      mCF(nNewIndex).lRefCount = mCF(nNewIndex).lRefCount + 1

      If nIndex >= 0 Then
         '// Decrement reference count for previous entry
         mCF(nIndex).lRefCount = mCF(nIndex).lRefCount - 1
      End If

   End If

   mItems(vRow).Cell(vCol).nFormat = nNewIndex

End Sub

Public Property Get ApplySelectionToImages() As Boolean
Attribute ApplySelectionToImages.VB_Description = "Returns/sets a value that determines whether the Focus bar color fills image background"

   ApplySelectionToImages = mbApplySelectionToImages

End Property

Public Property Let ApplySelectionToImages(ByVal vNewValue As Boolean)

   mbApplySelectionToImages = vNewValue
   PropertyChanged "ApplySelectionToImages"
   DrawGrid mbRedraw

End Property

Private Function AppThemed() As Boolean

   '---------------------------------------------------------------------------------
   '// MLH - new sub
   '// Purpose: Determines If The Current Window is Themed
   '---------------------------------------------------------------------------------

   On Error Resume Next
   AppThemed = IsAppThemed()
   On Error GoTo 0

End Function

Public Property Get AutoToolTips() As Boolean
Attribute AutoToolTips.VB_Description = "Returns/sets a value that determines whether truncated cells show tooltip"

   AutoToolTips = mbAutoToolTips

End Property

Public Property Let AutoToolTips(ByVal vNewValue As Boolean)

   mbAutoToolTips = vNewValue
   PropertyChanged "AutoToolTips"

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets a value that determines the Grid background color"

   BackColor = mBackColor

End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)

   mBackColor = vNewValue
   PropertyChanged "BackColor"
   DrawGrid mbRedraw

End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets a value that determines the Application background color of the grid"

   BackColorBkg = mBackColorBkg

End Property

Public Property Let BackColorBkg(ByVal vNewValue As OLE_COLOR)

   mBackColorBkg = vNewValue
   UserControl.BackColor = mBackColorBkg
   PropertyChanged "BackColorBkg"
   DisplayChange

End Property

Public Property Get BackColorEdit() As OLE_COLOR
Attribute BackColorEdit.VB_Description = "Returns/sets a value that determines the Background color of the Edit box"

   BackColorEdit = mBackColorEdit

End Property

Public Property Let BackColorEdit(ByVal vNewValue As OLE_COLOR)

   mBackColorEdit = vNewValue
   PropertyChanged "BackColorEdit"

End Property

Public Property Get BackColorEvenRows() As OLE_COLOR
Attribute BackColorEvenRows.VB_Description = "Returns/sets a value that determines the Color of even rows (see BackColorEvenRowsEnabled)"

   BackColorEvenRows = mBackColorEvenRows '// MLH - new sub

End Property

Public Property Let BackColorEvenRows(ByVal vNewValue As OLE_COLOR)

   mBackColorEvenRows = vNewValue '// MLH - new sub
   PropertyChanged "BackColorEvenRows"

End Property

Public Property Get BackColorEvenRowsEnabled() As Boolean
Attribute BackColorEvenRowsEnabled.VB_Description = "Returns/sets a value that determines whether Odd/Even rows have different background colors"

   BackColorEvenRowsEnabled = mbBackColorEvenRowsE '// MLH - new sub

End Property

Public Property Let BackColorEvenRowsEnabled(ByVal vNewValue As Boolean)

   mbBackColorEvenRowsE = vNewValue '// MLH - new sub
   PropertyChanged "BackColorEvenRowsEnabled"

End Property

Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_Description = "Returns/sets a value that determines the Background color of selected rows"

   BackColorSel = mBackColorSel

End Property

Public Property Let BackColorSel(ByVal vNewValue As OLE_COLOR)

   mBackColorSel = vNewValue
   PropertyChanged "BackColorSel"
   DisplayChange

End Property

Public Sub BindControl(ByVal vCol As Long, _
                       ByRef Ctrl As Object, _
                       Optional ByVal MoveControl As lgMoveControlEnum = lgBCHeight Or lgBCLeft Or lgBCTop Or lgBCWidth)

   '---------------------------------------------------------------------------------
   '// Purpose: Bind an external Control to a Column

   '// Col    - Column Index
   '// Ctrl   - The Control!
   '// Resize - Specify how the Control Size should be modified
   '---------------------------------------------------------------------------------

   Set mCols(vCol).EditCtrl = Ctrl
   mCols(vCol).MoveControl = MoveControl

End Sub

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, _
                            ByVal oColorTo As OLE_COLOR, _
                            Optional ByVal Alpha As Long = 128) As Long

  Dim lCFrom As Long
  Dim lCTo   As Long
  Dim lSrcR  As Long
  Dim lSrcG  As Long
  Dim lSrcB  As Long
  Dim lDstR  As Long
  Dim lDstG  As Long
  Dim lDstB  As Long

   lCFrom = oColorFrom
   lCTo = oColorTo
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000

   BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
               ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
               ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Function

Public Property Get BorderStyle() As lgBorderStyleEnum
Attribute BorderStyle.VB_Description = "Returns/sets a value that determines the border style of the grid (None or Single)"

   BorderStyle = muBorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewValue As lgBorderStyleEnum)

   muBorderStyle = vNewValue
   UserControl.BorderStyle = muBorderStyle
   PropertyChanged "BorderStyle"

End Property

Public Property Get CacheIncrement() As Long
Attribute CacheIncrement.VB_Description = "Increase arrays by ? - larger number makes it run faster but uses more memory"

   CacheIncrement = mCacheIncrement

End Property

Public Property Let CacheIncrement(ByVal vNewValue As Long)

   If vNewValue < 0 Then
      mCacheIncrement = 1
   Else
      mCacheIncrement = vNewValue
   End If

   PropertyChanged "CacheIncrement"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets a value that determines the Grid's Caption"

   Caption = msCaption

End Property

Public Property Let Caption(ByVal vNewValue As String)

   msCaption = vNewValue
   PropertyChanged "Caption"
   Call CreateRenderData
   UserControl.Cls
   Call DrawCaption
   Call DisplayChange

End Property

Public Property Get CaptionAlignment() As lgCaptionAlignmentEnum
Attribute CaptionAlignment.VB_Description = "Returns/sets a value that determines the Grid's Caption Alignment"

   CaptionAlignment = muCaptionAlignment

End Property

Public Property Let CaptionAlignment(ByVal vNewValue As lgCaptionAlignmentEnum)

   muCaptionAlignment = vNewValue
   PropertyChanged "CaptionAlignment"
   Call DisplayChange

End Property

Public Property Get CellAlignment(ByVal vRow As Long, ByVal vCol As Long) As lgAlignmentEnum

   CellAlignment = mItems(mRowPtr(vRow)).Cell(vCol).nAlignment

End Property

Public Property Let CellAlignment(Optional ByVal vRow As Long, _
                                  Optional ByVal vCol As Long, _
                                  ByVal vNewValue As lgAlignmentEnum)

   If FixRef(vRow, vCol) Then '// MLH
      mItems(mRowPtr(vRow)).Cell(vCol).nAlignment = vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellBackColor(ByVal vRow As Long, ByVal vCol As Long) As Long

   CellBackColor = mCF(mItems(mRowPtr(vRow)).Cell(vCol).nFormat).lBackColor

End Property

Public Property Let CellBackColor(Optional ByVal vRow As Long = C_NULL_RESULT, _
                                  Optional ByVal vCol As Long = C_NULL_RESULT, _
                                  ByVal vNewValue As Long)

   If FixRef(vRow, vCol) Then '// MLH
      ApplyCellFormat vRow, vCol, lgCFBackColor, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellChanged(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   CellChanged = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLChanged

End Property

Private Property Let CellChanged(ByVal vRow As Long, ByVal vCol As Long, ByVal vNewValue As Boolean)

   SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLChanged, vNewValue
   mbCellsChanged = True '// MLH
   SetFlag mItems(vRow).nFlags, lgFLChanged, True '// MLH Set RowChanged

End Property

Public Property Get CellChecked(Optional ByVal vRow As Long = C_NULL_RESULT, _
                                Optional ByVal vCol As Long = C_NULL_RESULT) As Boolean

   If FixRef(vRow, vCol) Then '// MLH
      CellChecked = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLChecked
   End If

End Property

Public Property Let CellChecked(Optional ByVal vRow As Long = C_NULL_RESULT, _
                                Optional ByVal vCol As Long = C_NULL_RESULT, _
                                ByVal vNewValue As Boolean)

   If FixRef(vRow, vCol) Then '// MLH
      SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLChecked, vNewValue

      If mCols(vCol).nType = lgBoolean Then '// MLH - If Boolean set CellText to new value
         mItems(mRowPtr(vRow)).Cell(vCol).sValue = CStr(vNewValue)
         SetRowSize vRow
      End If

      CellChanged(vRow, vCol) = True '// MLH
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellFontBold(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   CellFontBold = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLFontBold

End Property

Public Property Let CellFontBold(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Boolean)

   If FixRef(vRow, vCol) Then '// MLH
      SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLFontBold, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellFontItalic(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   CellFontItalic = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLFontItalic

End Property

Public Property Let CellFontItalic(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Boolean)

   If FixRef(vRow, vCol) Then '// MLH
      SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLFontItalic, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellFontName(ByVal vRow As Long, ByVal vCol As Long) As String

   CellFontName = mCF(mItems(mRowPtr(vRow)).Cell(vCol).nFormat).sFontName

End Property

Public Property Let CellFontName(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As String)

   If FixRef(vRow, vCol) Then '// MLH
      ApplyCellFormat vRow, vCol, lgCFFontName, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellFontUnderline(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   CellFontUnderline = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLFontUnderline

End Property

Public Property Let CellFontUnderline(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Boolean)

   If FixRef(vRow, vCol) Then '// MLH
      SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLFontUnderline, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellForeColor(ByVal vRow As Long, ByVal vCol As Long) As Long

   CellForeColor = mCF(mItems(mRowPtr(vRow)).Cell(vCol).nFormat).lForeColor

End Property

Public Property Let CellForeColor(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Long)

   If FixRef(vRow, vCol) Then '// MLH
      ApplyCellFormat vRow, vCol, lgCFForeColor, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellImage(ByVal vRow As Long, ByVal vCol As Long) As Variant

   CellImage = mItems(mRowPtr(vRow)).Cell(vCol).lImage

   '// MLH - Did not work - Row, Col did not return the correct index (.nFormat?)
   '  Dim nImage As Integer
   '   CellImage = mCF(mItems(mRowPtr(vRow)).Cell(mColPtr(vCol)).nFormat).nImage
   '   If nImage >= 0 Then
   '      CellImage = nImage
   '   Else
   '      CellImage = moImageList.ListImages(Abs(nImage)).Key
   '   End If

End Property

Public Property Let CellImage(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Variant)

  Dim nImage As Integer

   On Error Resume Next

   If FixRef(vRow, vCol) Then '// MLH

      If IsNumeric(vNewValue) Then
         nImage = vNewValue
      Else
         nImage = moImageList.ListImages(vNewValue).Index '// MLH - was negitive number
      End If

      mItems(mRowPtr(vRow)).Cell(vCol).lImage = nImage '// MLH

      '// MLH - Did not work - put images in cell that were not suppose to get them (.nFormat?)
      '''ApplyCellFormat vRow, vCol, lgCFImage, nImage
      DrawGrid mbRedraw

   End If

End Property

Public Sub CellPicture(ByVal vNewValue As StdPicture, _
                       Optional ByVal vRow As Long = C_NULL_RESULT, _
                       Optional ByVal vCol As Long = C_NULL_RESULT)

   '// MLH - New
   If FixRef(vRow, vCol) Then '// MLH
      Set mItems(mRowPtr(vRow)).Cell(vCol).pPic = vNewValue
   End If

End Sub

Public Function CellPictureGet(Optional ByVal vRow As Long = C_NULL_RESULT, _
                               Optional ByVal vCol As Long = C_NULL_RESULT) As StdPicture

   '// MLH - New
   If FixRef(vRow, vCol) Then '// MLH
      If Not (mItems(mRowPtr(vRow)).Cell(vCol).pPic Is Nothing) Then
         Set CellPictureGet = mItems(mRowPtr(vRow)).Cell(vCol).pPic
      End If
   End If

End Function

Public Property Get CellProgressBarColor() As OLE_COLOR
Attribute CellProgressBarColor.VB_Description = "Returns/sets a value that determines the Color of progress bar"

   CellProgressBarColor = mProgressBarColor

End Property

Public Property Let CellProgressBarColor(ByVal vNewValue As OLE_COLOR)

   mProgressBarColor = vNewValue
   PropertyChanged "ProgressBarColor"
   DrawGrid mbRedraw

End Property

Public Property Get CellProgressValue(ByVal vRow As Long, ByVal vCol As Long) As Integer
Attribute CellProgressValue.VB_Description = "Returns/sets a value that determines the length of the progress bar. A value between 0 to 100"

   If mCols(vCol).nType = lgProgressBar Then
      CellProgressValue = mItems(mRowPtr(vRow)).Cell(vCol).nFlags
   End If

End Property

Public Property Let CellProgressValue(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Integer)

   If FixRef(vRow, vCol) Then '// MLH

      If mCols(vCol).nType = lgProgressBar Then

         Select Case vNewValue
         Case Is > 100
            vNewValue = 100

         Case Is < 0
            vNewValue = 0
         End Select

         mItems(mRowPtr(vRow)).Cell(vCol).nFlags = vNewValue
         DrawGrid mbRedraw
      End If

   End If

End Property

Public Function CellsChanged() As Boolean

   CellsChanged = mbCellsChanged

End Function

Public Property Get CellTag(Optional ByVal vRow As Long = C_NULL_RESULT, _
                            Optional ByVal vCol As Long = C_NULL_RESULT) As String

   If FixRef(vRow, vCol) Then '// MLH
      CellTag = mItems(mRowPtr(vRow)).Cell(vCol).sTag
   End If

End Property

Public Property Let CellTag(Optional ByVal vRow As Long = C_NULL_RESULT, _
                            Optional ByVal vCol As Long = C_NULL_RESULT, _
                            ByVal vNewValue As String)

   '// MLH - New
   If FixRef(vRow, vCol) Then '// MLH
      mItems(mRowPtr(vRow)).Cell(vCol).sTag = vNewValue
   End If

End Property

Public Property Get CellText(Optional ByVal vRow As Long = C_NULL_RESULT, _
                             Optional ByVal vCol As Long = C_NULL_RESULT) As String

   If FixRef(vRow, vCol) Then '// MLH
      CellText = mItems(mRowPtr(vRow)).Cell(vCol).sValue
   End If

End Property

Public Property Let CellText(Optional ByVal vRow As Long = C_NULL_RESULT, _
                             Optional ByVal vCol As Long = C_NULL_RESULT, _
                             ByVal vNewValue As String)

   '// MLH - removed mColPrt(vCol) from all Cell...

   If FixRef(vRow, vCol) Then '// MLH

      If mCols(vCol).nType = lgBoolean Then '// MLH - If Boolean set CellChecked
         SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLChecked, CBool(vNewValue)
      End If
      
      mudtTotalsVal(mColPtr(vCol)) = mudtTotalsVal(mColPtr(vCol)) - rVal(mItems(mRowPtr(vRow)).Cell(vCol).sValue) + rVal(vNewValue) '// MLH - New

      mItems(mRowPtr(vRow)).Cell(vCol).sValue = vNewValue
      SetRowSize vRow

      CellChanged(vRow, vCol) = True '// MLH
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CellValue(Optional ByVal vRow As Long = C_NULL_RESULT, _
                              Optional ByVal vCol As Long = C_NULL_RESULT) As Variant

   On Error Resume Next
   If FixRef(vRow, vCol) Then '// MLH

      If mCols(vCol).nType = lgBoolean Then
         If LenB(mItems(mRowPtr(vRow)).Cell(vCol).sValue) Then
            CellValue = CBool(mItems(mRowPtr(vRow)).Cell(vCol).sValue)
         Else
            CellValue = False
         End If

      Else
         CellValue = rVal(mItems(mRowPtr(vRow)).Cell(vCol).sValue)
      End If
   End If

End Property

Public Property Let CellValue(Optional ByVal vRow As Long = C_NULL_RESULT, _
                              Optional ByVal vCol As Long = C_NULL_RESULT, _
                              ByVal vNewValue As Variant)

   '// MLH - New

   If FixRef(vRow, vCol) Then '// MLH

      If mCols(vCol).nType = lgBoolean Then
         mItems(mRowPtr(vRow)).Cell(vCol).sValue = CStr(CBool(vNewValue))
         SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLChecked, CBool(vNewValue)
      End If
      
      mudtTotalsVal(mColPtr(vCol)) = mudtTotalsVal(mColPtr(vCol)) - rVal(mItems(mRowPtr(vRow)).Cell(vCol).sValue) + rVal(vNewValue)

      mItems(mRowPtr(vRow)).Cell(vCol).sValue = CStr(rVal(vNewValue))
      SetRowSize vRow

      CellChanged(vRow, vCol) = True
      DrawGrid mbRedraw

   End If

End Property

Public Property Get CellWordWrap(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   CellWordWrap = mItems(mRowPtr(vRow)).Cell(vCol).nFlags And lgFLFontItalic

End Property

Public Property Let CellWordWrap(Optional ByVal vRow As Long, Optional ByVal vCol As Long, ByVal vNewValue As Boolean)

   If FixRef(vRow, vCol) Then '// MLH
      SetFlag mItems(mRowPtr(vRow)).Cell(vCol).nFlags, lgFLWordWrap, vNewValue
      DrawGrid mbRedraw
   End If

End Property

Public Property Get CenterRowImage() As Boolean
Attribute CenterRowImage.VB_Description = "Returns/sets a value that determines if the Row Image is vertically centered"

   '// MLH - new sub
   CenterRowImage = mbCenterRowImage

End Property

Public Property Let CenterRowImage(ByVal vNewValue As Boolean)

   mbCenterRowImage = vNewValue
   PropertyChanged "CenterRowImage"

End Property

Public Function CheckedCount() As Long

   '// Purpose: Return Count of Checked Rows

  Dim lCount As Long

   If (mRowCount = C_NULL_RESULT) Then
      CheckedCount = 0

   Else
      For lCount = LBound(mItems) To UBound(mItems)
         If mItems(lCount).nFlags And lgFLChecked Then
            CheckedCount = CheckedCount + 1
         End If
      Next lCount
   End If

End Function

Private Function CheckForLockedRow(ByVal vbMoveUp As Boolean) As Long

  Dim lNewRow As Long

   '// MLH - New Function

   If vbMoveUp Then
      lNewRow = NavigateUp()
   Else
      lNewRow = NavigateDown()
   End If

   Do

      If RowLocked(lNewRow) Then
         If vbMoveUp Then
            lNewRow = lNewRow - 1

            If lNewRow < 0 Then
               lNewRow = 0
               Exit Do
            End If

         Else
            lNewRow = lNewRow + 1

            If lNewRow > mRowCount Then
               lNewRow = mRowCount
               Exit Do
            End If

         End If

      Else
         Exit Do
      End If

   Loop

   CheckForLockedRow = lNewRow

End Function

Public Sub Clear()

   '// Purpose: Remove all Items from the Grid. Does not affect Column Headers

   '// Clear arrays
   Erase mItems         '// MLH
   Erase mRowPtr        '// MLH
   Erase mCF            '// MLH
   Erase mudtTotalsVal  '// MLH
   
   '// set default array dim
   ReDim mItems(0) As udtItem
   ReDim mRowPtr(0) As Long
   ReDim mCF(0) As udtFormat
   ReDim mudtTotalsVal(UBound(mColPtr)) As Double '// MLH - New

   '// clear system variables
   mMouseRow = C_NULL_RESULT '// MLH
   mMouseDownCol = C_NULL_RESULT
   mMouseDownRow = C_NULL_RESULT

   mCol = C_NULL_RESULT
   mRow = C_NULL_RESULT
   mSelectedRow = C_NULL_RESULT

   mHotColumn = C_NULL_RESULT
   mSwapCol = C_NULL_RESULT
   mDragCol = C_NULL_RESULT
   mResizeCol = C_NULL_RESULT
   mResizeRow = C_NULL_RESULT '// MLH - New

   mSortColumn = C_NULL_RESULT
   mSortSubColumn = C_NULL_RESULT

   mScrollAction = C_SCROLL_NONE
   mRowCount = C_NULL_RESULT

   mbUserRedraw = False       '// MLH
   mbRedraw = False           '// MLH
   mbEditPending = False      '// MLH
   txtEdit.Visible = False    '// MLH
   mbCellsChanged = False     '// MLH
   picTooltip.Visible = False '// MLH

   '// clear displayed grid
   Call DrawGrid(True)

End Sub

Public Sub ClearAll()
   
   '// MLH - New
   '// Purpose: Remove all Items from the Grid, including Column definitions

   Erase mCols
   Erase mColPtr
   Erase mudtTotals
   
   ReDim mCols(0) As udtColumn
   ReDim mColPtr(0) As Long
   ReDim mudtTotals(0) As typTotals
   
   Call Clear

End Sub

Public Property Get Col() As Long
Attribute Col.VB_Description = "Returns/sets a value that of the selected column"

   If Not (mCol = C_NULL_RESULT) Then
      Col = mColPtr(mCol)
   Else
      Col = C_NULL_RESULT
   End If

End Property

Public Property Let Col(ByVal vCol As Long)

   Call RowColSet(, vCol) '// MLH

End Property

Public Property Get ColAlignment(ByVal vCol As Long) As lgAlignmentEnum

   ColAlignment = mCols(vCol).nAlignment

End Property

Public Property Let ColAlignment(ByVal vCol As Long, ByVal vNewValue As lgAlignmentEnum)

   mCols(vCol).nAlignment = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get ColFormat(ByVal vCol As Long) As String

   ColFormat = mCols(vCol).sFormat

End Property

Public Property Let ColFormat(ByVal vCol As Long, ByVal vNewValue As String)

   mCols(vCol).sFormat = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get ColHeading(ByVal vCol As Long) As String

   ColHeading = mCols(vCol).sCaption

End Property

Public Property Let ColHeading(ByVal vCol As Long, ByVal vNewValue As String)

   mCols(vCol).sCaption = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get ColImageAlignment(ByVal vCol As Long) As lgAlignmentEnum

   ColImageAlignment = mCols(vCol).nImageAlignment

End Property

Public Property Let ColImageAlignment(ByVal vCol As Long, ByVal vNewValue As lgAlignmentEnum)

   mCols(vCol).nImageAlignment = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get ColInputFilter(ByVal vCol As Long) As String

   ColInputFilter = mCols(vCol).sInputFilter

End Property

Public Property Let ColInputFilter(ByVal vCol As Long, ByVal vNewValue As String)

   mCols(vCol).sInputFilter = vNewValue

End Property

Public Property Get ColLocked(ByVal vCol As Long) As Boolean

   If vCol = C_NULL_RESULT Or mLCLocked Then
      ColLocked = True
   Else
      ColLocked = mCols(vCol).bLocked
   End If

End Property

Public Property Let ColLocked(ByVal vCol As Long, ByVal vNewValue As Boolean)

   mCols(vCol).bLocked = vNewValue
   DrawGrid mbRedraw

End Property

Private Function ColorBrightness(ByVal lngColor As Long, Optional ByVal Alpha As Integer = -50) As Long

   '// MLH - new sub
   '// Purpose: Change the brightness of the passed color

  Dim lngR As Long
  Dim lngG As Long
  Dim lngB As Long

   lngColor = TranslateColor(lngColor)

   lngR = (lngColor And &HFF) + Alpha
   lngG = ((lngColor And &HFF00&) \ &H100&) + Alpha
   lngB = ((lngColor And &HFF0000) \ &H10000) + Alpha

   If lngR < 0 Then lngR = 0
   If lngG < 0 Then lngG = 0
   If lngB < 0 Then lngB = 0

   If lngR > 255 Then lngR = 255
   If lngG > 255 Then lngG = 255
   If lngB > 255 Then lngB = 255

   ColorBrightness = RGB(lngR, lngG, lngB)

End Function

Public Sub ColForceFit()

  Dim lngI        As Long
  Dim lngVCount   As Long
  Dim dblWidth    As Double

   '// MLH - Suggested by Paulo Cezar
   '// Force fit Column widths to grid
   
   SetRedrawState False
   
   '// Get number if visible columns
   For lngI = 0 To UBound(mCols)
      If ColVisible(lngI) Then
         lngVCount = lngVCount + 1
      End If
   Next lngI
   
   dblWidth = 1 / lngVCount
   dblWidth = VisibleWidth * dblWidth
   
   For lngI = 0 To UBound(mCols)
      ColWidth(lngI) = dblWidth
   Next lngI
   
   mbPendingRedraw = True
   SetRedrawState True
   If mbUserRedraw Then Call Refresh

End Sub

Public Sub ColOrderLoad(ByVal UniqueGridName As String)

   '// MLH - New
   '// Purpose: Load user ordered columns

  Dim lngI As Long

   lngI = rVal(GetSetting(App.Title, UniqueGridName, "Count"))
   '// don't load column widths if the column count changed.
   If lngI = UBound(mCols) Then
      For lngI = 0 To UBound(mColPtr)
         mColPtr(lngI) = rVal(GetSetting(App.Title, UniqueGridName, CStr(lngI), CStr(lngI)))
      Next lngI
   
      Call DisplayChange
   End If
   
End Sub

Public Sub ColOrderRestore(ByVal UniqueGridName As String)

   '// MLH - New
   '// Purpose: Restore default column order

  Dim lngI As Long

   For lngI = 0 To UBound(mColPtr)
      SaveSetting App.Title, UniqueGridName, CStr(lngI), CStr(lngI)
   Next lngI

   Call ColOrderLoad(UniqueGridName)

End Sub

Public Sub ColOrderSave(ByVal UniqueGridName As String)

   '// MLH - New
   '// Purpose: Save user ordered columns

  Dim lngI As Long

   SaveSetting App.Title, UniqueGridName, "Count", CStr(UBound(mCols))
   For lngI = 0 To UBound(mColPtr)
      SaveSetting App.Title, UniqueGridName, CStr(lngI), CStr(mColPtr(lngI))
   Next lngI

End Sub

Public Property Get Cols() As Long

   Cols = UBound(mCols) + 1

End Property

Public Property Get ColTag(ByVal vCol As Long) As String

   ColTag = mCols(vCol).sTag

End Property

Public Property Let ColTag(ByVal vCol As Long, ByVal vNewValue As String)

   mCols(vCol).sTag = vNewValue

End Property

Public Property Get ColType(ByVal vCol As Long) As lgDataTypeEnum

   ColType = mCols(vCol).nType

End Property

Public Property Let ColType(ByVal vCol As Long, ByVal vNewValue As lgDataTypeEnum)

   mCols(vCol).nType = vNewValue

End Property

Public Property Get ColumnHeaderLines() As Integer
Attribute ColumnHeaderLines.VB_Description = "Returns/sets a value that determines the number of lines to display column names"

   '// MLH - new sub
   ColumnHeaderLines = mColumnHeaderLines

End Property

Public Property Let ColumnHeaderLines(ByVal vNewValue As Integer)

   If vNewValue > 0 Then
      mColumnHeaderLines = vNewValue
      PropertyChanged "ColumnHeaderLines"
   Else
      mColumnHeaderLines = 1
   End If

   Call CreateRenderData
   Call DisplayChange

End Property

Public Property Get ColumnHeaders() As Boolean
Attribute ColumnHeaders.VB_Description = "Returns/sets a value that determines whether column headers are visible (Yes/No)"

   '// MLH - new sub
   ColumnHeaders = mbColumnHeaders

End Property

Public Property Let ColumnHeaders(ByVal vNewValue As Boolean)

   mbColumnHeaders = vNewValue
   PropertyChanged "ShowColumnHeaders"
   Call CreateRenderData
   Call DisplayChange

End Property

Public Property Get ColumnHeaderSmall() As Boolean
Attribute ColumnHeaderSmall.VB_Description = "Returns/sets a value that determines whether to use the minimum vertical height to display column header"

   ColumnHeaderSmall = mblnColumnHeaderSmall

End Property

Public Property Let ColumnHeaderSmall(ByVal vNewValue As Boolean)

   mblnColumnHeaderSmall = vNewValue

End Property

Public Property Get ColVisible(ByVal vCol As Long) As Boolean

   ColVisible = mCols(vCol).bVisible

End Property

Public Property Let ColVisible(ByVal vCol As Long, ByVal vNewValue As Boolean)

   mCols(vCol).bVisible = vNewValue

   DrawGrid mbRedraw

End Property

Public Sub ColWidthAutoSize(Optional ByVal vCol As Long = C_NULL_RESULT)
   
  Dim lngI As Long
  
   If vCol = C_NULL_RESULT Then
      For lngI = 0 To UBound(mCols)
         Call ColWAS(lngI)
      Next lngI
   
   Else
      Call ColWAS(vCol)
   End If
   
End Sub

Private Sub ColWAS(ByVal vCol As Long)

  Dim lngR          As Long
  Dim lngL          As Long
  Dim lngW          As Long
  Dim strTemp       As String
  Dim bBold         As Boolean
  Dim bItalic       As Boolean
  Dim bUnderLine    As Boolean
  Dim sFontName     As String
   
   '// MLH -New: Autoresize Column
   '// Size column to fit it's content
   On Error Resume Next
   
   Select Case mCols(vCol).nType
   Case lgBoolean
      ColWidth(vCol) = 480
   
   Case lgProgressBar, lgButton
      '// do nothing
   
   Case Else
      With UserControl
         bBold = .FontBold
         bItalic = .FontItalic
         bUnderLine = .FontUnderline
         sFontName = .FontName
         
         If mbTotalsLineShow Then
            If mCols(vCol).nType = lgNumeric Then
               strTemp = mudtTotals(vCol).sCaption & Format$(mudtTotalsVal(vCol), mCols(vCol).sFormat)
               lngL = .TextWidth(strTemp)
            End If
         End If
            
         For lngR = 0 To mRowCount
            .FontBold = mItems(lngR).Cell(vCol).nFlags And lgFLFontBold
            .FontItalic = mItems(lngR).Cell(vCol).nFlags And lgFLFontItalic
            .FontUnderline = mItems(lngR).Cell(vCol).nFlags And lgFLFontUnderline
            .FontName = mCF(mItems(lngR).Cell(vCol).nFormat).sFontName
            
            strTemp = Format$(mItems(lngR).Cell(vCol).sValue, mCols(vCol).sFormat)
            lngW = .TextWidth(strTemp)
            
            If vCol = 0 Then
               If mItems(lngR).lImage Then
                  lngW = lngW + mR.ImageWidth
               End If
            ElseIf mItems(lngR).Cell(vCol).lImage Then
               lngW = lngW + mR.ImageWidth
            End If
            
            If lngL < lngW Then
               lngL = lngW
            End If
         Next lngR
         
         .FontBold = bBold
         .FontItalic = bItalic
         .FontUnderline = bUnderLine
         .FontName = sFontName
      
      End With
      
      ColWidth(vCol) = (lngL + Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
   End Select

End Sub

Public Property Get ColWidth(ByVal vCol As Long) As Single

   ColWidth = mCols(vCol).dCustomWidth

End Property

Public Property Let ColWidth(ByVal vCol As Long, ByVal vNewValue As Single)

   '// dCustomWidth is in the Units the Control is operating in
   On Error Resume Next
   mCols(vCol).dCustomWidth = vNewValue
   mCols(vCol).lWidth = ScaleX(vNewValue, muScaleUnits, vbPixels)

   DrawGrid mbRedraw

End Property

Public Sub ColWidthsLoad(ByVal UniqueGridName As String)

  Dim lngI As Long

   For lngI = 0 To UBound(mColPtr)
      Me.ColWidth(lngI) = rVal(GetSetting(App.Title, UniqueGridName, "W" & CStr(lngI), Me.ColWidth(lngI)))
   Next lngI

End Sub

Public Sub ColWidthsSave(ByVal UniqueGridName As String)

  Dim lngI As Long

   For lngI = 0 To UBound(mColPtr)
      SaveSetting App.Title, UniqueGridName, "W" & CStr(lngI), CStr(Me.ColWidth(lngI))
   Next lngI

End Sub

Public Property Get ColWordWrap(ByVal vCol As Long) As Boolean

   ColWordWrap = mCols(vCol).nFlags And lgFLWordWrap

End Property

Public Property Let ColWordWrap(ByVal vCol As Long, ByVal vNewValue As Boolean)

   SetFlag mCols(vCol).nFlags, lgFLWordWrap, vNewValue

End Property

Private Sub CreateRenderData()

   '// Purpose: Calculates rendering parameters & sets display options.
   '// Used to prevent unneccesary recalculations when redrawing the Grid

  Dim lCount  As Long

   With mR

      If mMinRowHeight > C_MAX_CHECKBOXSIZE Then
         .CheckBoxSize = C_MAX_CHECKBOXSIZE
      Else
         .CheckBoxSize = mMinRowHeight - 4
      End If

      If mbCheckboxes Then '// Row CheckMarks?
         .LeftText = .CheckBoxSize + 2
      Else
         .LeftImage = 0
         .LeftText = C_TEXT_SPACE
      End If

      .LeftImage = .LeftText

      If moImageList Is Nothing Then
         '.ImageSpace = 0'// MLH
         .ImageSpace = False

      Else
         '// MLH - calculated in sub Drawgrid
         '.ImageSpace = ((mMinRowHeight - moImageList.ImageHeight) / 2)
         .ImageSpace = True '// MLH

         .ImageHeight = moImageList.ImageHeight
         .ImageWidth = moImageList.ImageWidth

         For lCount = LBound(mItems) To UBound(mItems)

            If Not (mItems(lCount).lImage = 0) Then
               .LeftText = .LeftText + moImageList.ImageWidth + 2
               Exit For
            End If

         Next lCount
      End If

      .TextHeight = UserControl.TextHeight(C_CHECKTEXT)

      If LenB(msCaption) Then
         .CaptionHeight = .TextHeight * 1.5
      Else
         .CaptionHeight = 0
      End If

      .HeaderHeight = GetColumnHeadingHeight()

      '// MLH - set minimum row height that the text will fit into
      '         add vertical offset from grid lines
      mMinRowHeight = ScaleY(mMinRowHeightUser, muScaleUnits, vbPixels)

      If mMinRowHeight < .TextHeight Then
         mMinRowHeight = .TextHeight
      End If

      mMinRowHeight = mMinRowHeight + (mMinVerticalOffset * 2) + 2

      If mbDisplayEllipsis Then
         .DTFlag = DT_WORD_ELLIPSIS 'OR DT_SINGLELINE
      Else
         .DTFlag = 0 'DT_SINGLELINE
      End If

   End With

End Sub

Public Sub DeleteSelected()

   '// MLH - New
   '// Purpose:(Delete Selected Rows)

  Dim lngR As Long

   If muMultiSelect Then
      '// turn off redraw in necessary
      SetRedrawState False

      Do

         If RowSelected(lngR) Then
            Call RemoveItem(lngR)
         Else
            lngR = lngR + 1
         End If

      Loop Until lngR = mRowCount + 1

      Call RowColSet(0)

      '// Restore redraw state to user selected
      SetRedrawState True
      DrawGrid mbRedraw

   Else
      Call RemoveItem
   End If

End Sub

Private Sub DisplayChange()

   If mbRedraw Then
      Call Refresh
   Else
      mbPendingRedraw = True
      mbPendingScrollBar = True
   End If

End Sub

Public Property Get DisplayEllipsis() As Boolean
Attribute DisplayEllipsis.VB_Description = "Returns/sets a value that determines whether to show … if a cell's text is truncated because of the column size"

   DisplayEllipsis = mbDisplayEllipsis

End Property

Public Property Let DisplayEllipsis(ByVal vNewValue As Boolean)

   mbDisplayEllipsis = vNewValue
   PropertyChanged "DisplayEllipsis"
   DisplayChange

End Property

Private Sub DrawCaption()

  Dim r As Rect

   If LenB(msCaption) > 0 Then

      With UserControl
         .ForeColor = mForeColorHdr

         Call SetRect(r, -1, 0, ScaleX(VisibleWidth, vbTwips, vbPixels) + 1, mR.CaptionHeight)

         'DrawRect .hDC, r, TranslateColor(BackColorFixed), True '// MLH - Not Needed

         Select Case muThemeStyle
         Case lgTSWindows3D
            Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH)

         Case lgTSWindowsFlat
            Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_FLAT)

         Case lgTSWindowsXP
            '// Try XP Theme API
            If Not DrawTheme("Header", 1, 1, r) Then
               '// Use XP emulation
               DrawXPHeader .hdc, r, 1
            End If

         Case lgTSOfficeXP
            DrawOfficeXPHeader .hdc, r, 1

         End Select

         r.Top = mR.CaptionHeight \ 8
         r.Right = r.Right - 15
         r.Left = r.Left + 15
         Call DrawText(.hdc, msCaption, -1, r, muCaptionAlignment)

      End With
   End If

End Sub

Private Sub DrawGrid(ByVal bRedraw As Boolean, _
                     Optional ByVal bHideFocusRect As Boolean = False)

  '// Purpose: The Primary Rendering routine. Draws Columns & Rows
  
  If mblnDrwGrid Then Exit Sub
  mblnDrwGrid = True
  DoEvents

  Dim IR             As Rect
  Dim r              As Rect
  Dim lX             As Long
  Dim lY             As Long
  Dim lCol           As Long
  Dim lRow           As Long
  Dim lMaxRow        As Long
  Dim lStartCol      As Long
  Dim lColumnsWidth  As Long
  Dim lBottomEdge    As Long
  Dim lGridColor     As Long
  Dim lImageLeft     As Long
  '''Dim lRowWrapSize   As Long     '// MLH
  Dim lStart         As Long
  Dim lValue         As Long
  Dim nImage         As Integer
  Dim bLockColor     As Boolean
  Dim sText          As String
  Dim bBold          As Boolean
  Dim bItalic        As Boolean
  Dim bUnderLine     As Boolean
  Dim sFontName      As String
  Dim lImgTop        As Long     '// MLH
  Dim bLineFeeds     As Boolean  '// MLH
  Dim strTemp        As String   '// MLH
  Dim cWidth         As Long     '// MLH
  Dim bAtFreeze      As Boolean  '// MLH
  Dim sngAspect      As Single   '// MLH
  Dim sngWidth       As Single   '// MLH
  Dim sngHeight      As Single   '// MLH
  Dim lRowsVisible   As Long     '// MLH
  Dim lRP            As Long     '// MLH
  Dim bToggle        As Boolean  '// MLH

   On Error Resume Next

   If bRedraw Then
      lStartCol = SBValue(efsHorizontal)
      lGridColor = TranslateColor(mGridColor)

      If lStartCol < mlngFreezeAtCol Then '// MLH - FreezeAtCol
         lStartCol = mlngFreezeAtCol + 1
         SBValue(efsHorizontal) = lStartCol
      End If

      lY = mR.HeaderHeight
      lRowsVisible = RowsVisible()

      With UserControl
         .Cls

         '// save usercontrol defaults
         bBold = .FontBold
         bItalic = .FontItalic
         bUnderLine = .FontUnderline
         sFontName = .FontName

         Call DrawCaption
         lColumnsWidth = DrawHeaderRow()
         
         If mRowCount < 0 Then GoTo Exit_Here  '// MLH - exit sub if there is nothing to do

         lStart = SBValue(efsVertical)
         
         lMaxRow = lStart + lRowsVisible '// MLH - call SBValue(efsVertical) once
         If lMaxRow > mRowCount Then lMaxRow = mRowCount

         If mbAllowWordWrap Or mbAllowRowResizing Then '// MLH - adjust first visible row
            If mRow > 0 Then
               If miKeyCode = vbKeyDown Then
                  lRP = lStart
                  lValue = (VisibleHeight / Screen.TwipsPerPixelY) - mR.HeaderHeight

                  If mbTotalsLineShow Then
                     lBottomEdge = mMinRowHeight
                  End If

                  For lRow = lStart To lMaxRow
                     If mItems(mRowPtr(lRow)).bVisible Then '// MLH - new
                        lRP = lRP + CInt(mItems(mRowPtr(lRow)).lHeight / mMinRowHeight)
                        lBottomEdge = lBottomEdge + mItems(mRowPtr(lRow)).lHeight
                     End If
                     If lRow = mRow Then Exit For
                  Next lRow

                  If lRP > lMaxRow And lBottomEdge > lValue Then
                     lStart = lStart + (lRP - lMaxRow)
                     If lStart > mRow Then lStart = mRow
                     SBValue(efsVertical) = lStart
                     lMaxRow = lStart + lRowsVisible
                  End If

               End If
            End If
         End If '// mbAllowWordWrap

         If lMaxRow > mRowCount Then
            lMaxRow = mRowCount
         End If
         
         '// Begin drawing visible rows
         For lRow = lStart To mRowCount '// MLH - was lMaxRow

            If mItems(mRowPtr(lRow)).bGroupRow Or mItems(mRowPtr(lRow)).bVisible Then  '// MLH - new
            
               If (muMultiSelect > 0 Or mbFullRowSelect) And (mItems(mRowPtr(lRow)).nFlags And lgFLSelected) Then
                  bLockColor = True
   
                  If lStartCol = 0 Then '// Code for row 0 column only
                     If mCols(0).lWidth < mR.LeftText Then
                        SetRect r, 0, lY + 1, mCols(0).lWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                     Else
                        SetRect r, 0, lY + 1, mR.LeftText, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                     End If
   
                     If mbBackColorEvenRowsE Then '// MLH
                        bToggle = Not bToggle
                        If bToggle Then
                           DrawRect .hdc, r, TranslateColor(mBackColor), True
                        Else
                           DrawRect .hdc, r, TranslateColor(mBackColorEvenRows), True
                        End If
   
                     Else
                        DrawRect .hdc, r, TranslateColor(mBackColor), True
                     End If
   
                  Else '// First row is not visible
                     r.Right = 0
                  End If
   
                  SetRect r, r.Right - 1, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1   '// MLH
   
                     '// MLH - See Below
'''                  If mbAlphaBlendSelection Then
'''                     lValue = BlendColor(TranslateColor(mBackColorSel), TranslateColor(mBackColor), 120)
'''                  Else
'''                     lValue = TranslateColor(mBackColorSel)
'''                  End If
                  
                  If mbAlphaBlendSelection Then
                     '// MLH
                     lValue = mItems(mRowPtr(lRow)).Cell(0).nFormat
                     lValue = mCF(lValue).lBackColor
                     If lValue = mBackColor Then
                        'lValue = TranslateColor(mBackColorSel)
                        lValue = BlendColor(TranslateColor(mBackColorSel), TranslateColor(mBackColor), 150)
                     Else
                        lValue = BlendColor(TranslateColor(lValue), TranslateColor(mBackColor), 150)
                     End If
                  Else
                     lValue = TranslateColor(mBackColorSel)
                  End If
   
                  Select Case muFocusRowHighlightStyle '// MLH - gradient
                  Case [Solid]
                     DrawRect .hdc, r, lValue, True
   
                  Case [Gradient_H]
                     Call FillGradient(.hdc, r, lValue, TranslateColor(mBackColor), False)
   
                  Case [Gradient_V]
                     Call FillGradient(.hdc, r, lValue, TranslateColor(mBackColor), True)
                  End Select
   
                  .ForeColor = mForeColorSel
   
               Else '// row not selected
                  bLockColor = False
                  SetRect r, 0, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
   
                  If mbBackColorEvenRowsE Then '// MLH
                     bToggle = Not bToggle
                     If bToggle Then
                        DrawRect .hdc, r, TranslateColor(mBackColor), True
                     Else
                        DrawRect .hdc, r, TranslateColor(mBackColorEvenRows), True
                     End If
   
                  Else
                     DrawRect .hdc, r, TranslateColor(mBackColor), True
                  End If
   
               End If
   
               lX = 0
   
               '// Loop for each visible column
               'For lCol = lStartCol To UBound(mCols) '// MLH - was before FreezeAtCol
               For lCol = 0 To UBound(mCols)          '// MLH - FreezeAtCol
                  .ForeColor = mForeColorSel
                  If lCol <= mlngFreezeAtCol Or lCol >= lStartCol Then '// MLH - FreezeAtCol
   
                     If mCols(mColPtr(lCol)).bVisible Then
                        SetRectRgn mClipRgn, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight
                        SelectClipRgn .hdc, mClipRgn
   
                        Call SetRect(r, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight)
   
                        If Not bLockColor Then
                           If Not (mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor = mBackColor) Then
                              DrawRect .hdc, r, _
                                 TranslateColor(mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor), True
                           End If
                           .ForeColor = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lForeColor
                        
                        ElseIf mblnKeepForeColor Then
                           .ForeColor = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lForeColor
                        End If
   
                        '----------------------------------------------------------------------------------------------
                        If lCol = 0 Then '// Code for column 0 only (checkbox and row image)
                           If Not mItems(mRowPtr(lRow)).bGroupRow And mbCheckboxes Then '// Row CheckMarks?
                              Call SetRect(r, 3, lY, mR.CheckBoxSize, lY + mItems(mRowPtr(lRow)).lHeight)
   
                              If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                 lValue = 5
                              Else
                                 lValue = 0
                              End If
   
                              If Not DrawTheme("Button", 3, lValue, r) Then
                                 If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                 Else
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                 End If
   
                              End If
                           End If
   
                           '// Draw Row Image in cell 0 if it exists
                           If mR.ImageSpace Then
                              '// If we have an Image Index then Draw it
                              If Not (mItems(mRowPtr(lRow)).lImage = 0) Then
                                 '// Calculate Image offset (using ScaleMode of ImageList)
                                 If lImageLeft = 0 Then
                                    lImageLeft = ScaleX(mR.LeftImage, vbPixels, mImageListScaleMode)
                                 End If
   
                                 '// Center Row Image? '// MLH
                                 If mbCenterRowImage Then
                                    lImgTop = (mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight) \ 2
                                 Else
                                    lImgTop = mMinVerticalOffset
                                 End If
   
                                 If bLockColor And mbApplySelectionToImages Then
                                    moImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY _
                                       + lImgTop, vbPixels, mImageListScaleMode), 2
                                 Else
                                    moImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY _
                                       + lImgTop, vbPixels, mImageListScaleMode), 1
                                 End If
   
                              End If
                           End If
   
                           Call SetRect(r, mR.LeftText + C_TEXT_SPACE, lY, (lX + mCols(mColPtr(lCol)).lWidth) - C_TEXT_SPACE, _
                              lY + mItems(mRowPtr(lRow)).lHeight)
   
                           '----------------------------------------------------------------------------------------------
                        Else '// all columns but 0
                           Call SetRect(r, lX + C_TEXT_SPACE, lY, (lX + mCols(mColPtr(lCol)).lWidth) - C_TEXT_SPACE, lY + _
                              mItems(mRowPtr(lRow)).lHeight)
   
                        End If '// column = 0
   
                        '----------------------------------------------------------------------------------------------
                        '// Determine Column type
                        Select Case mCols(mColPtr(lCol)).nType
                        Case lgBoolean
                           If Not mItems(mRowPtr(lRow)).bGroupRow Then '// MLH- New
                              SetItemRect lRow, lCol, lY, r, lgRTCheckBox
      
                              If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                 lValue = 5
                              Else
                                 lValue = 0
                              End If
      
                              If Not DrawTheme("Button", 3, lValue, r) Then
                                 If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                 Else
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                 End If
      
                              End If
                           End If
      
'                        Case lgProgressBar '// MLH - Moved to allow Text on Progress bar
'                           If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags > 0 Then
'                              lValue = ((mCols(mColPtr(lCol)).lWidth - 2) / 100) * mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags
'                              SetRect R, lX + 2, lY + 2, lX + lValue, (lY + mItems(mRowPtr(lRow)).lHeight) - 2
'                              DrawRect .hDC, R, TranslateColor(mProgressBarColor), True
'                           End If
                           
                        Case Else
                           If Not mItems(mRowPtr(lRow)).bGroupRow Then '// MLH- New
                              Select Case mCols(mColPtr(lCol)).nType
                              Case lgButton '// MLH
                                 Call SetRect(r, lX, lY, (lX + mCols(mColPtr(lCol)).lWidth), lY + mItems(mRowPtr(lRow)).lHeight)
      
                                 If mbMouseDown And mMouseDownRow = lRow And mCol = lCol Then
                                    Call DrawXPButton(r, lgDOWN)
                                    'Call DrawTheme("Button", 1, lgDOWN, R)
                                 Else
                                    Call DrawXPButton(r, lgNormal)
                                    'Call DrawTheme("Button", 1, lgNormal, R)
                                 End If
      
                                 .ForeColor = vbButtonText
      
                              Case lgProgressBar '// MLH - Moved to allow Text on Progress bar
                                 If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags > 0 Then
                                    lValue = ((mCols(mColPtr(lCol)).lWidth - 2) / 100) * mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags
                                    SetRect r, lX + 2, lY + 2, lX + lValue, (lY + mItems(mRowPtr(lRow)).lHeight) - 2
                                    DrawRect .hdc, r, TranslateColor(mProgressBarColor), True
                                 End If
      
                                 SetRect r, lX + C_TEXT_SPACE, lY, (lX + mCols(mColPtr(lCol)).lWidth) - C_TEXT_SPACE, lY + mItems(mRowPtr(lRow)).lHeight
                              End Select
                           End If
                           
                           '// Normal text and Col not 0
                           UserControl.FontName = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).sFontName
   
                           With mItems(mRowPtr(lRow)).Cell(mColPtr(lCol))
   
                              If Not mItems(mRowPtr(lRow)).bGroupRow And mCols(mColPtr(lCol)).nType = lgProgressBar Then '// MLH - Progress bar value is stored in .nFlags
                                 UserControl.FontBold = False
                                 UserControl.FontItalic = False
                                 UserControl.FontUnderline = False
   
                              Else
                                 UserControl.FontBold = .nFlags And lgFLFontBold
                                 UserControl.FontItalic = .nFlags And lgFLFontItalic
                                 UserControl.FontUnderline = .nFlags And lgFLFontUnderline
                              End If
   
                              cWidth = mCols(mColPtr(lCol)).lWidth - 15
   
                              '// format text if necessary
                              If LenB(mCols(mColPtr(lCol)).sFormat) Then
                                 sText = Format$(.sValue, mCols(mColPtr(lCol)).sFormat)
                              Else
                                 sText = .sValue
                              End If
   
                              lValue = .nAlignment Or DT_SINGLELINE  '// MLH Moved
                              nImage = .lImage '// Is there an image index for this cell? '// MLH - new
                              bLineFeeds = InStrB(1, sText, vbCr) '// MLH
   
                              If nImage Or (lCol = 0 And mItems(mRowPtr(lRow)).lImage) Then
                                 cWidth = cWidth - mR.ImageWidth
                              End If
   
                              '// if word wrap and the user has set a Min Row Height > 0 '// MLH
                              If mbAllowWordWrap And (mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLWordWrap) And _
                                 mMinRowHeightUser And Not bLineFeeds Then
   
                                 If nImage Or (lCol = 0 And mItems(mRowPtr(lRow)).lImage) Then
                                    strTemp = SplitToLines(sText, cWidth)
                                 Else
                                    strTemp = SplitToLines(sText, cWidth)
                                 End If
   
                                 bLineFeeds = InStrB(1, strTemp, vbCr)
                              End If
   
                              SetRect IR, 0, 0, cWidth, mItems(mRowPtr(lRow)).lHeight
   
                              If nImage Or (lCol = 0 And mItems(mRowPtr(lRow)).lImage) Then
                                 IR.Right = IR.Right - mR.ImageWidth
                              End If
   
                              Call DrawText(UserControl.hdc, sText, Len(sText), IR, DT_CALCRECT Or DT_SINGLELINE)
   
                              '// Is word wrapping necessary? '// MLH
                              If mbAllowWordWrap And (mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLWordWrap) And _
                                 (IR.Right + 6 > cWidth And mMinRowHeight < mItems(mRowPtr(lRow)).lHeight Or bLineFeeds) Then
   
                                 SetRect IR, 0, 0, cWidth, mItems(mRowPtr(lRow)).lHeight
   
                                 lValue = DT_WORDBREAK
                                 Call DrawText(UserControl.hdc, sText, Len(sText), IR, DT_CALCRECT Or DT_WORDBREAK)
   
                                 If (IR.Bottom - IR.Top) > mR.TextHeight Then
                                    nImage = mExpandRowImage
                                 Else
                                    nImage = 0
                                 End If
   
                                 '// MLH
                                 'If mItems(mRowPtr(lRow)).lHeight < mMinRowHeight Then '// MLH - was lRowWrapSize
                                 'Truncate Rect to prevent wrapped text from showing
                                 '   R.Bottom = R.Top + mR.TextHeight
                                 'End If
   
                                 r.Top = r.Top + mMinVerticalOffset '// MLH
                              End If
   
                              If Not (mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic Is Nothing) Then
                                 '// MLH - Draw Cell Picture if necessary
                                 SetItemRect lRow, lCol, lY, IR, lgRTColumn
   
                                 sngAspect = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic.Height / _
                                    mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic.Width
                                 sngWidth = IR.Right - IR.Left
                                 sngHeight = IR.Bottom - IR.Top
   
                                 If sngHeight / sngWidth > sngAspect Then
                                    sngHeight = sngAspect * sngWidth
                                 Else
                                    sngWidth = sngHeight / sngAspect
                                 End If
   
                                 '// Center picture in cell?
                                 Select Case mCols(mColPtr(lCol)).nImageAlignment
                                 Case lgAlignCenterCenter
                                    cWidth = (mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic.Width \ Screen.TwipsPerPixelX)
   
                                    If cWidth < mCols(mColPtr(lCol)).lWidth Then
                                       IR.Left = IR.Left + (Abs(mCols(mColPtr(lCol)).lWidth - sngWidth) \ 2) - 2
                                    End If
   
                                 Case lgAlignRightTop, lgAlignRightBottom, lgAlignRightCenter
                                    cWidth = (mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic.Width \ Screen.TwipsPerPixelX)
   
                                    If cWidth < mCols(mColPtr(lCol)).lWidth Then
                                       IR.Left = IR.Left + (mCols(mColPtr(lCol)).lWidth - cWidth - 5)
                                    End If
   
                                 End Select
   
                                 UserControl.PaintPicture mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).pPic, IR.Left, IR.Top, _
                                    sngWidth, sngHeight, , , , , vbSrcCopy
   
                              Else
                                 '// Draw Cell Image if necessary
                                 If Not (nImage = 0) Then
                                    SetItemRect lRow, lCol, lY, IR, lgRTImage
   
                                    '// MLH - Aligh cell image
                                    Select Case mCols(mColPtr(lCol)).nImageAlignment
                                    Case lgAlignLeftTop, lgAlignCenterTop ', lgAlignCenterCenter '// MLH
                                       lImgTop = mMinVerticalOffset
                                       r.Left = r.Left + (IR.Right - IR.Left)
   
                                    Case lgAlignLeftBottom, lgAlignCenterBottom
                                       lImgTop = mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight - mMinVerticalOffset
                                       r.Left = r.Left + (IR.Right - IR.Left)
   
                                    Case lgAlignLeftCenter
                                       lImgTop = (mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight) \ 2
                                       r.Left = r.Left + (IR.Right - IR.Left)
   
                                    Case lgAlignRightTop
                                       If mExpandRowImage Then lImgTop = 0
                                       lImgTop = mMinVerticalOffset
                                       r.Right = r.Right - (IR.Right - IR.Left)
   
                                    Case lgAlignRightBottom
                                       lImgTop = mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight - mMinVerticalOffset
                                       r.Right = r.Right - (IR.Right - IR.Left)
   
                                    Case lgAlignRightCenter
                                       lImgTop = (mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight) \ 2
                                       r.Right = r.Right - (IR.Right - IR.Left)
   
                                    Case lgAlignCenterCenter '// MLH
                                       lImgTop = (mItems(mRowPtr(lRow)).lHeight - mR.ImageHeight - mMinVerticalOffset) \ 2
                                       r.Right = r.Right + ((r.Right - (IR.Right - IR.Left)) \ 2)
   
                                    End Select
   
                                    If IR.Left >= 0 Then
                                       If bLockColor And mbApplySelectionToImages Then
                                          moImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, _
                                             mImageListScaleMode), ScaleY(lY + lImgTop, vbPixels, mImageListScaleMode), 2
   
                                       Else
                                          moImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, _
                                             mImageListScaleMode), ScaleY(lY + lImgTop, vbPixels, mImageListScaleMode), 1
                                       End If
   
                                    End If '// End IR.Left >= 0
                                 End If '// End Draw image
   
                              End If
                              Call DrawText(UserControl.hdc, sText, -1, r, lValue)
   
                           End With
   
                        End Select
   
                        lX = lX + mCols(mColPtr(lCol)).lWidth
   
                     End If '// End mCols(mColPtr(lCol)).bVisible
                  End If '// MLH - FreezeAtCol
   
                  '// MLH - Don't draw columns that are beyond the grid's boarder (faster draws).
                  If lX > .ScaleWidth Then Exit For
   
               Next lCol
   
               SelectClipRgn .hdc, 0&
   
               '// Display Horizontal Lines
               If muGridLines = lgGrid_Both Or muGridLines = lgGrid_Horizontal Then
                  DrawLine .hdc, 0, lY, lColumnsWidth, lY, lGridColor, mGridLineWidth
               End If
   
               lY = lY + mItems(mRowPtr(lRow)).lHeight
               '// MLH - Stop drawing rows that are beyond the grid's boarder.
               If lY > .ScaleHeight Then Exit For
            
            End If '// mItems(mRowPtr(lRow)) Visible
         Next lRow

         '---------------------------------------------------------------------------------
         '// MLH - Display Totals Line
         If mbTotalsLineShow Then
            If lRow > mRowCount Then
               Dim dblTemp As Double

               .ForeColor = mForeColorHdr
               SetRect r, 0, lY + 1, lColumnsWidth, lY + mMinRowHeight + 1

               Select Case muThemeStyle
               Case lgTSWindows3D
                  Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH)

               Case lgTSWindowsFlat
                  Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_FLAT)

               Case lgTSWindowsXP
                  If Not DrawTheme("Header", 1, 1, r) Then '// Try XP Theme API
                     DrawXPHeader .hdc, r, 1               '// Use XP emulation
                  End If

               Case lgTSOfficeXP
                  DrawOfficeXPHeader .hdc, r, 1
               End Select

               DrawLine .hdc, 0, lY, lColumnsWidth, lY, mForeColorHdr, mGridLineWidth

               lX = 0

               For lCol = 0 To UBound(mCols)

                  If lCol <= mlngFreezeAtCol Or lCol >= lStartCol Then
                     If mCols(mColPtr(lCol)).bVisible Then
                        Call SetRect(r, lX + C_TEXT_SPACE, lY + C_TEXT_SPACE, (lX + mCols(mColPtr(lCol)).lWidth) - C_TEXT_SPACE, lY + mMinRowHeight)
                        lValue = mCols(mColPtr(lCol)).nAlignment

                        If mCols(mColPtr(lCol)).nType = lgNumeric Then '// MLH - New
                           dblTemp = mudtTotalsVal(mColPtr(lCol))
                           If mudtTotals(mColPtr(lCol)).bAvg Then dblTemp = dblTemp / mRowCount

                           If LenB(mCols(mColPtr(lCol)).sFormat) Then
                              sText = Format$(dblTemp, mCols(mColPtr(lCol)).sFormat)
                           Else
                              sText = dblTemp
                           End If

                           If LenB(mudtTotals(mColPtr(lCol)).sCaption) Then
                              sText = mudtTotals(mColPtr(lCol)).sCaption & " " & sText
                           End If

                           Call DrawText(UserControl.hdc, sText, -1, r, lValue)

                        Else

                           If LenB(mudtTotals(mColPtr(lCol)).sCaption) Then
                              Call DrawText(UserControl.hdc, mudtTotals(mColPtr(lCol)).sCaption, -1, r, lValue)
                           End If

                        End If

                        lX = lX + mCols(mColPtr(lCol)).lWidth
                     End If

                  End If
               Next lCol

            End If
            
         Else
            DrawLine .hdc, 0, lY, lColumnsWidth, lY, lGridColor, mGridLineWidth '// MLH - New
         End If

         '---------------------------------------------------------------------------------
         '// Display Vertical Lines
         If muGridLines = lgGrid_Both Or muGridLines = lgGrid_Vertical Then
            lBottomEdge = r.Bottom
            lX = 0

            'For lCol = lStartCol To UBound(mCols)'// MLH - was before FreezeAtCol
            For lCol = 0 To UBound(mCols)         '// MLH - FreezeAtCol
               If lCol <= mlngFreezeAtCol Or lCol >= lStartCol Then '// MLH - FreezeAtCol

                  If mCols(mColPtr(lCol)).bVisible Then
                     If bAtFreeze Then '// MLH - FreezeAtCol
                        DrawLine .hdc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth * 2
                        bAtFreeze = False
                     Else
                        DrawLine .hdc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth
                     End If

                     lX = lX + mCols(mColPtr(lCol)).lWidth
                  End If

               End If

               If lCol = mlngFreezeAtCol Then bAtFreeze = True '// MLH - FreezeAtCol
            Next lCol
            DrawLine .hdc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth '// MLH - New
         End If

         '---------------------------------------------------------------------------------
         '// Display Focus Rectangle
         If Not (muFocusRectMode = lgFocusRectModeEnum.lgNone) And (mRow >= 0) Then
            If Not bHideFocusRect Then  '// MLH

               lY = RowTop(mRow, lStart) '// MLH

               If Not lY = C_NULL_RESULT Then
                  r.Right = 0

                  If muFocusRectMode = lgCol Then
                     SetColRect mCol, r
                     r.Top = lY + 1
                     r.Bottom = lY + mItems(mRowPtr(mRow)).lHeight

                  Else
                     'SetRect R, mR.LeftText, lY + 1, lColumnsWidth, lY + mItems(mRowPtr(mRow)).lHeight
                     SetRect r, 1, lY + 1, lColumnsWidth, lY + mItems(mRowPtr(mRow)).lHeight '// MLH
                  End If

                  If r.Right > 0 Then

                     Select Case muFocusRectStyle
                     Case lgFRLight
                        Call DrawFocusRect(.hdc, r)

                     Case lgFRHeavy '// MLH - New
                        UserControl.DrawWidth = 3

                        If mbFullRowSelect Then
                           UserControl.ForeColor = TranslateColor(mFocusRectColor)
                        Else
                           UserControl.ForeColor = ColorBrightness(mBackColorSel)
                        End If

                        Call RoundRect(.hdc, r.Left, r.Top, r.Right, r.Bottom, 0&, 0&)
                        UserControl.DrawWidth = 1

                     Case lgFRMedium '// MLH - renamed
                        If mbFullRowSelect Then '// MLH
                           DrawRect .hdc, r, TranslateColor(mFocusRectColor), False
                        Else
                           DrawRect .hdc, r, ColorBrightness(mBackColorSel), False
                        End If

                     End Select

                  End If '// R.Right > 0
               End If '// Not lY = C_NULL_RESULT
            End If '// Not bHideFocusRect
         End If

         .Refresh

         .FontBold = bBold
         .FontItalic = bItalic
         .FontUnderline = bUnderLine
         .FontName = sFontName
      End With

      mbPendingRedraw = False
      
   Else '// Redraw=FALSE
      mbPendingRedraw = True
   End If

Exit_Here:
   mblnDrwGrid = False

End Sub

Private Sub DrawHeader(ByVal lCol As Long, ByVal State As lgHeaderStateEnum, Optional ByVal bDraging As Boolean = False)

   '// Purpose: Renders a Column Header. This involves drawing the Border, displaying
   '// the Caption and optionally Sort Arrows

  Dim r         As Rect
  Dim lngCenter As Long
  Dim sText     As String

   UserControl.FontBold = ucFontBold
   UserControl.FontItalic = ucFontItalic
   UserControl.FontName = ucFontName

   If lCol > C_NULL_RESULT And mbColumnHeaders Then

      If mSwapCol = lCol Or mDragCol = lCol Then State = lgDOWN '// MLH
      If Not (mbAllowColumnSort Or mbAllowColumnSwap Or mbAllowColumnDrag) Then State = lgNormal '// MLH

      With UserControl
         .ForeColor = mForeColorHdr

         '// Draw the Column Headers - '// MLH
         Call SetRect(r, mCols(mColPtr(lCol)).lX - 1, mR.CaptionHeight, mCols(mColPtr(lCol)).lX + _
            mCols(mColPtr(lCol)).lWidth, mR.HeaderHeight)

         'DrawRect .hDC, r, TranslateColor(BackColorFixed), True '// MLH - Not Needed

         Select Case muThemeStyle
         Case lgTSWindows3D
            Select Case State
            Case lgNormal
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH)

            Case lghot
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)

            Case lgDOWN
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
            End Select

         Case lgTSWindowsFlat
            Select Case State
            Case lgNormal
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_FLAT)

            Case lghot
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)

            Case lgDOWN
               Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
            End Select

         Case lgTSWindowsXP
            '// Try XP Theme API
            If Not DrawTheme("Header", 1, State, r) Then
               '// Use XP emulation
               DrawXPHeader .hdc, r, State
            End If

         Case lgTSOfficeXP
            DrawOfficeXPHeader .hdc, r, State
         End Select

         '// Render Sort Arrows

         If mCols(mColPtr(lCol)).lWidth > C_SIZE_SORTARROW Then

            If mColPtr(lCol) = mSortColumn Then
               DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6 + mR.CaptionHeight, 9, 5, _
                  mCols(mColPtr(lCol)).nSortOrder

               Call SetRect(r, mCols(mColPtr(lCol)).lX + C_TEXT_SPACE, mR.CaptionHeight, (mCols(mColPtr(lCol)).lX + _
                  mCols(mColPtr(lCol)).lWidth) - (C_ARROW_SPACE + C_SIZE_SORTARROW), mR.HeaderHeight)

            ElseIf mColPtr(lCol) = mSortSubColumn Then
               DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6 + mR.CaptionHeight, 6, 3, _
                  mCols(mColPtr(lCol)).nSortOrder

               Call SetRect(r, mCols(mColPtr(lCol)).lX + C_TEXT_SPACE, mR.CaptionHeight, (mCols(mColPtr(lCol)).lX + _
                  mCols(mColPtr(lCol)).lWidth) - (C_ARROW_SPACE + C_SIZE_SORTARROW), mR.HeaderHeight)
            Else
               Call SetRect(r, mCols(mColPtr(lCol)).lX + C_TEXT_SPACE, mR.CaptionHeight, (mCols(mColPtr(lCol)).lX + _
                  mCols(mColPtr(lCol)).lWidth) - (C_TEXT_SPACE * 2), mR.HeaderHeight)
            End If

         Else
            Call SetRect(r, mCols(mColPtr(lCol)).lX + C_TEXT_SPACE, mR.CaptionHeight, (mCols(mColPtr(lCol)).lX + _
               mCols(mColPtr(lCol)).lWidth) - (C_TEXT_SPACE * 2), mR.HeaderHeight)
         End If

         '// MLH
         r.Left = r.Left + 2
         If mColumnHeaderLines > 1 And Not bDraging Then '// More than 1 line of text?
            '// needed to vertically center wrapped text
            sText = SplitToLines(mCols(mColPtr(lCol)).sCaption, r.Right - r.Left, mColumnHeaderLines)
            lngCenter = (mR.HeaderHeight - mR.CaptionHeight - UserControl.TextHeight(sText)) / 2

            r.Top = r.Top + lngCenter
            Call DrawText(.hdc, sText, -1, r, mCols(mColPtr(lCol)).nAlignment Or DT_WORDBREAK Or DT_WORD_ELLIPSIS)

         Else '// single line of text
            Call DrawText(.hdc, mCols(mColPtr(lCol)).sCaption, -1, r, mCols(mColPtr(lCol)).nAlignment Or DT_SINGLELINE)
         End If

      End With
   End If

End Sub

Private Function DrawHeaderRow(Optional ByVal bDraging As Boolean = False) As Long

   '// Purpose: Renders all Column Headers

  Dim lCol As Long
  Dim lX   As Long

   mHotColumn = C_NULL_RESULT

   'For lCol = SBValue(efsHorizontal) To UBound(mCols) '// MLH - was before FreezeAtCol
   For lCol = 0 To UBound(mCols)                       '// MLH - FreezeAtCol
      If lCol <= mlngFreezeAtCol Or lCol >= SBValue(efsHorizontal) Then '// MLH - FreezeAtCol

         If mCols(mColPtr(lCol)).bVisible Then
            mCols(mColPtr(lCol)).lX = lX
            Call DrawHeader(lCol, lgNormal, bDraging)
            lX = lX + mCols(mColPtr(lCol)).lWidth
         End If

      End If
   Next lCol

   DrawHeaderRow = lX

End Function

Private Sub DrawLine(ByVal hdc As Long, _
                     ByVal X1 As Long, _
                     ByVal Y1 As Long, _
                     ByVal X2 As Long, _
                     ByVal Y2 As Long, _
                     ByVal lColor As Long, _
                     ByVal lWidth As Long)

  Dim PT      As POINTAPI
  Dim hPen    As Long
  Dim hPenOld As Long

   hPen = CreatePen(0, lWidth, lColor)
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, X1, Y1, PT
   LineTo hdc, X2, Y2
   SelectObject hdc, hPenOld
   DeleteObject hPen

End Sub

Private Sub DrawOfficeXPHeader(ByVal lHDC As Long, ByRef rRect As Rect, ByVal State As lgHeaderStateEnum)

   '---------------------------------------------------------------------------------
   '// Purpose:   Draw a Column Header in Office XP Style
   '// Notes:     Created from original source by Riccardo Cohen
   '---------------------------------------------------------------------------------

   With rRect
      Select Case State
      Case lgNormal
         '// MLH
         Call FillGradient(lHDC, rRect, mlngCustomColorFrom, mlngCustomColorTo, True)

         DrawLine lHDC, .Left, .Top, .Right, .Top, mlngCustomColorFrom, 1
         DrawLine lHDC, .Left, .Bottom - 1, .Right, .Bottom - 1, mlngCustomColorFrom, 1

         DrawLine lHDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, mlngCustomColorTo, 1
         DrawLine lHDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

      Case lghot
         .Right = .Right - 1
         Call FillGradient(lHDC, rRect, &HDCFFFF, &H5BC0F7, True)

         DrawLine lHDC, .Left, .Top, .Right, .Top, &H9C613B, 1
         DrawLine lHDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1

         DrawLine lHDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

      Case lgDOWN
         .Right = .Right - 1
         Call FillGradient(lHDC, rRect, &H87FE8, &H7CDAF7, True)

         DrawLine lHDC, .Left, .Top, .Right, .Top, &H9C613B, 1
         DrawLine lHDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1

         DrawLine lHDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

      End Select
   End With

End Sub

Private Sub DrawRect(ByVal lngHDC As Long, ByRef Rc As Rect, ByVal lColor As Long, ByVal bFilled As Boolean)

  Dim lNewBrush As Long

   lNewBrush = CreateSolidBrush(lColor)

   If bFilled Then
      Call FillRect(lngHDC, Rc, lNewBrush)
   Else
      Call FrameRect(lngHDC, Rc, lNewBrush)
   End If

   Call DeleteObject(lNewBrush)

End Sub

Private Sub DrawSortArrow(ByVal lX As Long, _
                          ByVal lY As Long, _
                          ByVal lWidth As Long, _
                          ByVal lStep As Long, _
                          ByVal nOrientation As lgSortTypeEnum)

   '// Purpose: Renders the Sort/Sub-Sort arrows

  Dim hPenOld         As Long
  Dim hPen            As Long
  Dim lCount          As Long
  Dim lVerticalChange As Long
  Dim X1              As Long
  Dim X2              As Long
  Dim Y1              As Long

   If Not nOrientation = lgSTNormal Then

      hPen = CreatePen(0, 1, TranslateColor(vbButtonShadow))
      hPenOld = SelectObject(hdc, hPen)

      If nOrientation = lgSTDescending Then
         lVerticalChange = -1
         lY = lY + lStep - 1
      Else
         lVerticalChange = 1
      End If

      X1 = lX
      X2 = lWidth
      Y1 = lY

      MoveTo hdc, X1, Y1, ByVal 0&

      For lCount = 1 To lStep
         LineTo hdc, X1 + X2, Y1
         X1 = X1 + 1
         Y1 = Y1 + lVerticalChange
         X2 = X2 - 2
         MoveTo hdc, X1, Y1, ByVal 0&
      Next lCount

      Call SelectObject(hdc, hPenOld)
      Call DeleteObject(hPen)

   End If

End Sub

Private Sub DrawText(ByVal lngHDC As Long, _
                     ByVal lpString As String, _
                     ByVal nCount As Long, _
                     ByRef lpRect As Rect, _
                     ByVal wFormat As Long)

   '---------------------------------------------------------------------------------
   '// Purpose: Renders the Text for Column Headers & Cells.
   '// On Windows NT/2000/XP(or better) the Control supports Unicode

   '// MLH - added DT_NOPREFIX and moved mR.DTFlag here
   '---------------------------------------------------------------------------------

   If mbWinNT Then
      DrawTextW lngHDC, StrPtr(lpString), nCount, lpRect, wFormat Or DT_NOPREFIX Or mR.DTFlag
   Else
      DrawTextA lngHDC, lpString, nCount, lpRect, wFormat Or DT_NOPREFIX Or mR.DTFlag
   End If

End Sub

Private Function DrawTheme(ByVal sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           ByRef rtRect As Rect, _
                           Optional ByVal CloseTheme As Boolean = True) As Boolean

   '// Purpose: On Windows XP allows certain elements of the Grid to be drawn using the current Windows Theme

  Dim lResult As Long

   On Error GoTo DrawThemeError

   If mbWinXP Then
      If (muThemeStyle = lgTSWindowsXP) Or (muThemeStyle = lgTSOfficeXP) Then

         mhTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))

         If (mhTheme) Then
            lResult = DrawThemeBackground(mhTheme, UserControl.hdc, iPart, iState, rtRect, rtRect)

            If CloseTheme Then '// MLH - old version required a second call to closed the theme which never happened.
               Call CloseThemeData(mhTheme)
            End If

            DrawTheme = (lResult = 0)

         Else
            DrawTheme = False
         End If

      End If
   End If
   Exit Function

DrawThemeError:
   DrawTheme = False

End Function

Private Sub DrawXPButton(ByRef btnRect As Rect, ByVal lngState As Long)

  Dim lngTheme     As Long
  Dim strXPclass   As String

   On Error Resume Next

   strXPclass = "Button"
   btnRect.Bottom = btnRect.Bottom + 1
   btnRect.Right = btnRect.Right + 1

   If mbWinXP Then
      lngTheme = OpenThemeData(UserControl.hWnd, StrPtr(strXPclass))

      If lngTheme Then
         Call DrawThemeBackground(lngTheme, UserControl.hdc, 1, lngState, btnRect, btnRect)
         Call CloseThemeData(lngTheme)

      Else '// no themes
         DrawOfficeXPHeader UserControl.hdc, btnRect, lngState
      End If

   Else '// NOT XP or greater
      DrawOfficeXPHeader UserControl.hdc, btnRect, lngState
   End If

End Sub

Private Sub DrawXPHeader(ByVal lHDC As Long, ByRef rRect As Rect, ByVal State As lgHeaderStateEnum)

   '---------------------------------------------------------------------------------
   '// Purpose:   Draw a Column Header in XP Style
   '// Notes:     Created from original source by Riccardo Cohen
   '---------------------------------------------------------------------------------

  Dim TempColor As OLE_COLOR

   With rRect

      Select Case State
      Case lgNormal
         DrawRect lHDC, rRect, TranslateColor(vbButtonFace), True

         DrawLine lHDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &HB2C2C5, 1
         DrawLine lHDC, .Left, .Bottom - 2, .Right, .Bottom - 2, &HBECFD2, 1
         DrawLine lHDC, .Left, .Bottom - 3, .Right, .Bottom - 3, &HC8D8DC, 1

         DrawLine lHDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &H99A8AC, 1
         DrawLine lHDC, .Left, .Top + 2, .Left, .Bottom - 4, &HFFFFFF, 1

      Case lghot
         DrawRect lHDC, rRect, &HF3F8FA, True

         DrawLine lHDC, .Left + 2, .Bottom - 1, .Right - 2, .Bottom - 1, &H19B1F9, 1
         DrawLine lHDC, .Left + 1, .Bottom - 2, .Right - 1, .Bottom - 2, &H47C2FC, 1
         DrawLine lHDC, .Left, .Bottom - 3, .Right, .Bottom - 3, 43512, 1

      Case lgDOWN
         TempColor = ForeColor

         UserControl.ForeColor = InvertThisColor(TempColor)
         .Bottom = .Bottom - 1
         DrawRect lHDC, rRect, &H0&, True

         DrawLine lHDC, .Left, .Bottom - 1, .Right, .Bottom - 1, InvertThisColor(&HB2C2C5), 1
         DrawLine lHDC, .Left, .Bottom - 2, .Right, .Bottom - 2, InvertThisColor(&HBECFD2), 1
         DrawLine lHDC, .Left, .Bottom - 3, .Right, .Bottom - 3, InvertThisColor(&HC8D8DC), 1
         DrawLine lHDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, InvertThisColor(&H99A8AC), 1
         DrawLine lHDC, .Left, .Top + 2, .Left, .Bottom - 4, InvertThisColor(&HFFFFFF), 1
      End Select

   End With

End Sub

Private Sub EditCell(ByVal vRow As Long, ByVal vCol As Long)

   '// Purpose: Used to start an Edit. Note the BeforeEdit event. This event allows
   '// the Edit to be cancelled before anything visible occurs by setting the Cancel flag.

  Dim bCancel As Boolean
  Dim lTemp As Long

   If mbEditPending Then
      If Not UpdateCell() Then
         Exit Sub
      End If
   End If

   If IsEditable() And Not (mCols(mColPtr(vCol)).nType = lgBoolean) Then
      RaiseEvent BeforeEdit(vRow, mColPtr(vCol), bCancel)

      If Not bCancel Then
         mEditCol = vCol
         mEditRow = vRow

         MoveEditControl mCols(mColPtr(mEditCol)).MoveControl

         '// Check if an external Control is used.
         If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
            '// Using internal TextBox
            With txtEdit
               .Alignment = 0

               Select Case mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nAlignment
               Case lgAlignCenterBottom, lgAlignCenterCenter, lgAlignCenterTop
                  lTemp = vbCenter

               Case lgAlignLeftBottom, lgAlignLeftCenter, lgAlignLeftTop
                  lTemp = vbLeftJustify

               Case Else
                  lTemp = vbRightJustify
               End Select

               If mbWinNT Then
                  Select Case mCols(mColPtr(mEditCol)).sInputFilter
                  Case "<"
                     Call SetWindowLongW(.hWnd, GWL_STYLE, mTextBoxStyle Or ES_LOWERCASE)

                  Case ">"
                     Call SetWindowLongW(.hWnd, GWL_STYLE, mTextBoxStyle Or ES_UPPERCASE)

                  Case Else
                     Call SetWindowLongW(.hWnd, GWL_STYLE, mTextBoxStyle)
                  End Select

               Else
                  Select Case mCols(mColPtr(mEditCol)).sInputFilter
                  Case "<"
                     Call SetWindowLongA(.hWnd, GWL_STYLE, mTextBoxStyle Or ES_LOWERCASE)

                  Case ">"
                     Call SetWindowLongA(.hWnd, GWL_STYLE, mTextBoxStyle Or ES_UPPERCASE)

                  Case Else
                     Call SetWindowLongA(.hWnd, GWL_STYLE, mTextBoxStyle)
                  End Select
               End If

               On Local Error Resume Next
               .ForeColor = mForeColorEdit '// MLH
               .BackColor = mBackColorEdit
               If Not mCols(mColPtr(mEditCol)).nType = lgProgressBar Then '// MLH
                  .FontBold = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontBold
                  .FontItalic = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontItalic
                  .FontUnderline = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontUnderline
               End If
               .FontName = mCF(mItems(mRowPtr(vRow)).Cell(mColPtr(vCol)).nFormat).sFontName '// MLH
               .FontSize = UserControl.FontSize  '// MLH
               .Alignment = 0     '// MLH - Don't know why but it doesn't work without it
               .Alignment = lTemp '// MLH

               .Text = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue
               .SelStart = 0
               .SelLength = Len(.Text)
               .Visible = True
               .SetFocus
            End With

         Else '// External Control
            On Local Error Resume Next

            With mCols(mColPtr(mEditCol)).EditCtrl

               If Not (UserControl.ContainerHwnd = .Container.hWnd) Then
                  mEditParent = UserControl.ContainerHwnd
                  SetParent .hWnd, UserControl.ContainerHwnd
               Else
                  mEditParent = 0
               End If

               '// set edit attributes '// MLH
               .ForeColor = mForeColorEdit
               .BackColor = mBackColorEdit
               If Not mCols(mColPtr(mEditCol)).nType = lgProgressBar Then '// MLH
                  .FontBold = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontBold
                  .FontItalic = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontItalic
                  .FontUnderline = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontUnderline
               End If
               .FontName = mCF(mItems(mRowPtr(vRow)).Cell(mColPtr(vCol)).nFormat).sFontName
               .FontSize = UserControl.FontSize

               .Enabled = True
               .Visible = True
               .ZOrder

               sc_Subclass .hWnd
               Call sc_AddMsg(.hWnd, WM_KILLFOCUS, MSG_AFTER)

               If TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox Then
                  SendMessageAsLong mCols(mColPtr(mEditCol)).EditCtrl.hWnd, CB_SHOWDROPDOWN, 1&, 0&
               End If

               .SetFocus
            End With

            On Local Error GoTo 0
         End If

         mbEditPending = True
      End If '// Not bCancel

   End If

End Sub

Public Property Get EditMove() As lgEditMoveEnum
Attribute EditMove.VB_Description = "When editing, pressing the Enter key will do one of the following: Stay on current cell, Move Right, or Move Down"

   '// MLH - New
   EditMove = muEditMove

End Property

Public Property Let EditMove(ByVal vNewValue As lgEditMoveEnum)

   '// MLH - New
   muEditMove = vNewValue
   PropertyChanged "EditMove"

End Property

Public Property Get EditTrigger() As lgEditTriggerEnum
Attribute EditTrigger.VB_Description = "Returns/sets a value that determines how cell edit is started (enter key, double click, etc.)"

   EditTrigger = muEditTrigger

End Property

Public Property Let EditTrigger(ByVal vNewValue As lgEditTriggerEnum)

   muEditTrigger = vNewValue
   PropertyChanged "EditTrigger"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the control can respond to user-generated events"

   '// MLH - new sub
   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)

   '// MLH - new sub
   UserControl.Enabled = vNewValue
   PropertyChanged "Enabled"

End Property

Public Property Get ExpandRowImage() As Variant
Attribute ExpandRowImage.VB_Description = "Returns/sets the value of the ImageList image number used to expand/shrink row size"

   If mExpandRowImage >= 0 Then
      ExpandRowImage = mExpandRowImage
   Else
      ExpandRowImage = moImageList.ListImages(Abs(mExpandRowImage)).Key
   End If

End Property

Public Property Let ExpandRowImage(ByVal vNewValue As Variant)

   On Local Error GoTo ExpandRowImageError

   If IsNumeric(vNewValue) Then
      mExpandRowImage = vNewValue
   Else
      mExpandRowImage = moImageList.ListImages(vNewValue).Index '// MLH - was negative?
   End If

   DrawGrid mbRedraw

ExpandRowImageError:

End Property

Public Sub ExportGrid(ByVal strGridName As String, _
                      Optional ByVal vbVisibleColsOnly As Boolean = True, _
                      Optional ByVal vbVisibleRowsOnly As Boolean = False, _
                      Optional ByVal vbBooleanAsText As Boolean = True)

   '// MLH - new sub
  Dim lngR        As Long
  Dim lngC        As Long
  Dim lngF        As Long
  Dim dblTemp     As Double
  Dim strText     As String
  Dim lMaxCol     As Long
  Dim strFileName As String

   On Local Error GoTo Err_Proc

   strGridName = strGridName & ".csv"
   strFileName = GetDESKTOPDir & "\" & strGridName
   If LenB(Dir$(strFileName)) Then Kill strFileName

   lMaxCol = UBound(mCols)

   lngF = FreeFile
   Open strFileName For Output As #lngF
   
   If LenB(msCaption) Then
      Print #lngF, msCaption
   End If

   '// Header
   For lngC = 0 To lMaxCol

      If vbVisibleColsOnly Then
         If mCols(mColPtr(lngC)).bVisible Then
            Print #lngF, mCols(mColPtr(lngC)).sCaption & ",";
         End If

      Else
         Print #lngF, mCols(mColPtr(lngC)).sCaption & ",";
      End If

   Next lngC

   Print #lngF, ""

   '// Grid Data
   For lngR = 0 To mRowCount
      If Not vbVisibleRowsOnly Or mItems(mRowPtr(lngR)).bVisible Then
         For lngC = 0 To lMaxCol
            If vbVisibleColsOnly Then
               If mCols(mColPtr(lngC)).bVisible Then
                  If Not vbBooleanAsText Then
   
                     If mCols(mColPtr(lngC)).nType = lgBoolean Then
                        Print #lngF, CInt(CBool(mItems(mRowPtr(lngR)).Cell(mColPtr(lngC)).sValue)) & ",";
   
                     Else
                        Print #lngF, mItems(mRowPtr(lngR)).Cell(mColPtr(lngC)).sValue & ",";
                     End If
   
                  Else
                     Print #lngF, mItems(mRowPtr(lngR)).Cell(mColPtr(lngC)).sValue & ",";
                  End If
   
               End If
   
            Else
               Print #lngF, mItems(mRowPtr(lngR)).Cell(mColPtr(lngC)).sValue & ",";
            End If
         Next lngC
      End If

      Print #lngF, vbNullString
   Next lngR

   '// Totals Line
   If mbTotalsLineShow Then
      For lngC = 0 To lMaxCol

         strText = vbNullString

         If mCols(mColPtr(lngC)).nType = lgNumeric Then
            dblTemp = mudtTotalsVal(mColPtr(lngC))
            If mudtTotals(mColPtr(lngC)).bAvg Then dblTemp = dblTemp / mRowCount
            strText = CStr(dblTemp)

            If LenB(mudtTotals(mColPtr(lngC)).sCaption) Then
               strText = mudtTotals(mColPtr(lngC)).sCaption & " " & strText
            End If

         Else
            If LenB(mudtTotals(mColPtr(lngC)).sCaption) Then
               strText = mudtTotals(mColPtr(lngC)).sCaption
            End If
         End If

         If vbVisibleColsOnly Then
            If mCols(mColPtr(lngC)).bVisible Then
               Print #lngF, strText & ",";
            End If

         Else
            Print #lngF, strText & ",";
         End If

      Next lngC
   End If

   Close #lngF

   If MsgBox("File " & strGridName & " was saved to your desktop." & vbNewLine & "Do you want to open it?", vbQuestion Or _
      vbYesNo) = vbYes Then
      Call ExportGridOpen(strFileName)
   End If

   Exit Sub

Err_Proc:
   MsgBox "Error# " & Err.Number & vbNewLine & Err.Description, vbCritical, "LynxGrid.Export"
   Close

End Sub

Private Sub ExportGridOpen(ByVal vstrFileName As String)

   '// MLH - New
   'Const C_SW_NORMAL               As Long = &H1&
   'Const C_SEE_MASK_INVOKEIDLIST   As Long = &HC&
   'Const C_SEE_MASK_NOCLOSEPROCESS As Long = &H40&
   'Const C_SEE_MASK_FLAG_NO_UI     As Long = &H400&

  Dim udtSEI As typSHELLEXECUTEINFO

   With udtSEI
      '// Set the structure's size
      .cbSize = Len(udtSEI)
      '// Set the mask
      .fMask = &H44C 'C_SEE_MASK_NOCLOSEPROCESS Or C_SEE_MASK_INVOKEIDLIST Or C_SEE_MASK_FLAG_NO_UI
      '// Set the owner window
      .hWnd = 0&
      '// Set the action
      .lpVerb = "open"
      '// Set the File Path and Name
      .lpFile = vstrFileName & vbNullChar
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = &H1 'C_SW_NORMAL
      .hInstApp = 0&
      .lpIDList = 0&
   End With

   If ShellExecuteEx(udtSEI) = 0 Then
      MsgBox "Unable to open target file.", vbInformation
   End If

End Sub

Private Sub FillGradient(ByVal lHDC As Long, _
                         ByRef rRect As Rect, _
                         ByVal clrFirst As OLE_COLOR, _
                         ByVal clrSecond As OLE_COLOR, _
                         Optional ByVal bVertical As Boolean)

  Dim pVert(0 To 1)   As TRIVERTEX
  Dim pGradRect       As GRADIENT_RECT

   With pVert(0)
      .X = rRect.Left
      .Y = rRect.Top
      .Red = LongToSignedShort((clrFirst And &HFF&) * 256)
      .Green = LongToSignedShort(((clrFirst And &HFF00&) / &H100&) * 256)
      .Blue = LongToSignedShort(((clrFirst And &HFF0000) / &H10000) * 256)
      .Alpha = 0
   End With

   With pVert(1)
      .X = rRect.Right
      .Y = rRect.Bottom
      .Red = LongToSignedShort((clrSecond And &HFF&) * 256)
      .Green = LongToSignedShort(((clrSecond And &HFF00&) / &H100&) * 256)
      .Blue = LongToSignedShort(((clrSecond And &HFF0000) / &H10000) * 256)
      .Alpha = 0
   End With

   With pGradRect
      .UpperLeft = 0
      .LowerRight = 1
   End With

   '// MLH - was GradientFill
   GradientFillRect lHDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

Public Function FindItem(ByVal SearchText As String, _
                         Optional ByVal SearchColumn As Long = C_NULL_RESULT, _
                         Optional ByVal SearchMode As lgSearchModeEnum = lgSMEqual, _
                         Optional ByVal MatchCase As Boolean, _
                         Optional ByVal StartingRow As Long = 0) As Long

   '---------------------------------------------------------------------------------
   '// Purpose: Search the specified Column for a Cell that matches the search text

   '// SearchText     - The text to look for
   '// SearchColumn   - The Column to search in (defaults to the SearchColumn property if not specified)
   '// SearchMode     - The type of search required. The lgSMNavigate mode is used by the Grid internally
   '//                    when searching for an entry that matches the keys the user is pressing.
   '// MatchCase      - Specify a case sensitive or case insensitive search
   '// StartingRow    - The row from which the search starts
   '---------------------------------------------------------------------------------

  Dim lCount    As Long
  Dim lngI      As Long
  Dim sCellText As String

   FindItem = C_NULL_RESULT

   If mRowCount = C_NULL_RESULT Then '// MLH - prevent error
      If IsInIDE Then
         MsgBox "No Rows added to grid!", vbExclamation, "ERROR - FindItem"
      End If

      Exit Function
   End If

   If StartingRow < 0 Then '// define starting point '// MLH
      StartingRow = 0
   ElseIf StartingRow > mRowCount Then
      StartingRow = 0
   End If

   If SearchColumn = C_NULL_RESULT Then
      SearchColumn = mSearchColumn
   End If

   If (SearchColumn >= 0) And (LenB(SearchText) > 0) Then
      If Not MatchCase Then
         SearchText = UCase$(SearchText)
      End If

      lngI = Len(SearchText)

      For lCount = StartingRow To mRowCount

         If MatchCase Then
            sCellText = mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue
         Else
            sCellText = UCase$(mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue)
         End If

         Select Case SearchMode
         Case lgSMEqual
            If sCellText = SearchText Then
               FindItem = lCount
               Exit For
            End If

         Case lgSMGreaterEqual
            If sCellText >= SearchText Then
               FindItem = lCount
               Exit For
            End If

         Case lgSMLike
            If InStrB(1, sCellText, SearchText) Then
               FindItem = lCount
               Exit For
            End If

         Case lgSMNavigate
            If LenB(sCellText) Then
               If (sCellText >= SearchText) And ((Mid$(sCellText, 1, 1)) = Mid$(SearchText, 1, 1)) Then
                  FindItem = lCount
                  Exit For
               End If
            End If

         Case lgSMBeginsWith
            If Len(sCellText) >= lngI Then
               If Left$(sCellText, lngI) = SearchText Then
                  FindItem = lCount
                  Exit For
               End If
            End If

         End Select

      Next lCount
   End If

End Function

Private Function FixRef(ByRef vRow As Long, Optional ByRef vCol As Long) As Boolean

   '// MLH - new sub
   If Not (mRowCount = C_NULL_RESULT) Then

      Select Case vRow
      Case Is < 0
         If mRow = C_NULL_RESULT Then
            vRow = 0
         Else
            vRow = mRow
         End If

      Case Is > mRowCount
         vRow = mRowCount
      End Select

      Select Case vCol
      Case Is < 0
         If mCol = C_NULL_RESULT Then
            vCol = 0
         Else
            vCol = mCol
         End If

      Case Is > UBound(mCols)
         vCol = UBound(mCols)
      End Select

      FixRef = True

   ElseIf IsInIDE Then
      MsgBox "No Rows Added", vbExclamation, "ERROR - FixRef"
   End If

End Function

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_Description = "Returns/sets a value that determines the Focus Rectangle color"

   FocusRectColor = mFocusRectColor

End Property

Public Property Let FocusRectColor(ByVal vNewValue As OLE_COLOR)

   mFocusRectColor = vNewValue
   PropertyChanged "FocusRectColor"

End Property

Public Property Get FocusRectHide() As Boolean
Attribute FocusRectHide.VB_Description = "Returns/sets a value that determines whether the Focus Rectangle is visible when the grid losses the focus"

   FocusRectHide = mbHideSelection

End Property

Public Property Let FocusRectHide(ByVal vNewValue As Boolean)

   '// MLH - Renamed
   mbHideSelection = vNewValue
   PropertyChanged "HideSelection"
   DisplayChange

End Property

Public Property Get FocusRectMode() As lgFocusRectModeEnum
Attribute FocusRectMode.VB_Description = "Returns/sets a value that determines the Focus Rectangle Type (None, Row, Column)"

   FocusRectMode = muFocusRectMode

End Property

Public Property Let FocusRectMode(ByVal vNewValue As lgFocusRectModeEnum)

   muFocusRectMode = vNewValue
   PropertyChanged "FocusRectMode"
   DisplayChange

End Property

Public Property Get FocusRectStyle() As lgFocusRectStyleEnum
Attribute FocusRectStyle.VB_Description = "Returns/sets a value that determines whether the Focus Rectangle Style (Light, Medium, Heavy)"

   FocusRectStyle = muFocusRectStyle

End Property

Public Property Let FocusRectStyle(ByVal vNewValue As lgFocusRectStyleEnum)

   muFocusRectStyle = vNewValue
   PropertyChanged "FocusRectStyle"
   DisplayChange

End Property

Public Property Get FocusRowHighlight() As Boolean
Attribute FocusRowHighlight.VB_Description = "Returns/sets a value that determines whether the Row Focus bar is visible - On/Off"

   FocusRowHighlight = mbFullRowSelect

End Property

Public Property Let FocusRowHighlight(ByVal vNewValue As Boolean)

   '// MLH - renamed
   mbFullRowSelect = vNewValue
   PropertyChanged "FullRowSelect"
   DisplayChange

End Property

Public Property Get FocusRowHighlightKeepTextForecolor() As Boolean
Attribute FocusRowHighlightKeepTextForecolor.VB_Description = "Returns/sets a value that determines whether the Cells keep their forecolor when the Row is highlighted - On/Off"
   FocusRowHighlightKeepTextForecolor = mblnKeepForeColor

End Property

Public Property Let FocusRowHighlightKeepTextForecolor(ByVal vNewValue As Boolean)

   '// MLH - new sub
   mblnKeepForeColor = vNewValue
   PropertyChanged "FocusRowHighlightKeepTextForecolor"
   Call SetThemeColor
   Call DisplayChange

End Property

Public Property Get FocusRowHighlightStyle() As lgFocusRowHighlightStyle
Attribute FocusRowHighlightStyle.VB_Description = "Returns/sets a value that determines the style of the Focus bar (Solid, Gradient vertical/horizontal)"

   FocusRowHighlightStyle = muFocusRowHighlightStyle

End Property

Public Property Let FocusRowHighlightStyle(ByVal vNewValue As lgFocusRowHighlightStyle)

   '// MLH - new sub
   muFocusRowHighlightStyle = vNewValue
   PropertyChanged "FocusRowHighlightStyle"
   Call SetThemeColor
   Call DisplayChange

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Font Object"

   Set Font = mFont

End Property

Public Property Set Font(ByVal vNewValue As StdFont)

   Set mFont = vNewValue
   Set UserControl.Font = mFont
   PropertyChanged "Font"

   Call CreateRenderData '// MLH - save changes
   txtEdit.FontSize = UserControl.FontSize '// MLH
   DrawGrid mbRedraw

End Property

Public Sub ForceCellEdit(Optional ByVal lNewRow As Long = C_NULL_RESULT, _
                         Optional ByVal lNewCol As Long = C_NULL_RESULT, _
                         Optional ByVal bBlankCell As Boolean = False)

  Dim bAEdit As Boolean

   '// MLH - new sub
   '// Purpose: Force edit through code

   Call RowColSet(lNewRow, lNewCol)

   If Not (mRowCount = C_NULL_RESULT) Then
      bAEdit = mbAllowEdit
      miKeyCode = vbKeyF2
      mbAllowEdit = True
      Call EditCell(mRow, mCol)
      mbAllowEdit = bAEdit

      If bBlankCell Then
         txtEdit.Text = vbNullString
      End If

   End If

End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets a value that determines the Default Forecolor of the grid"

   ForeColor = mForeColor

End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)

   mForeColor = vNewValue
   PropertyChanged "ForeColor"

End Property

Public Property Get ForeColorEdit() As OLE_COLOR
Attribute ForeColorEdit.VB_Description = "Returns/sets a value that determines the Forecolor of Edit box"

   ForeColorEdit = mForeColorEdit

End Property

Public Property Let ForeColorEdit(ByVal vNewValue As OLE_COLOR)

   mForeColorEdit = vNewValue
   PropertyChanged "ForeColorEdit"

End Property

Public Property Get ForeColorHdr() As OLE_COLOR
Attribute ForeColorHdr.VB_Description = "Returns/sets a value that determines the Forecolor of Header Text"

   ForeColorHdr = mForeColorHdr

End Property

Public Property Let ForeColorHdr(ByVal vNewValue As OLE_COLOR)

   mForeColorHdr = vNewValue
   PropertyChanged "ForeColorHdr"

End Property

Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_Description = "Returns/sets a value that determines the Forecolor of Selected Text"

   ForeColorSel = mForeColorSel

End Property

Public Property Let ForeColorSel(ByVal vNewValue As OLE_COLOR)

   mForeColorSel = vNewValue
   PropertyChanged "ForeColorSel"
   DisplayChange

End Property

Public Sub FormatCells(Optional ByVal RowFrom As Long = 0, _
                       Optional ByVal RowTo As Long = C_NULL_RESULT, _
                       Optional ByVal ColFrom As Long = 0, _
                       Optional ByVal ColTo As Long = C_NULL_RESULT, _
                       Optional ByVal Mode As lgCellFormatEnum = lgCFImage, _
                       Optional ByVal vNewValue As String = vbNullString)

  Dim lCol As Long
  Dim lRow As Long

   On Error Resume Next
   
   If Not (mRowCount = C_NULL_RESULT) Then '// MLH - prevent error
      If RowTo = C_NULL_RESULT Then RowTo = mRowCount
      If ColTo = C_NULL_RESULT Then ColTo = UBound(mCols)

      For lRow = RowFrom To RowTo
         For lCol = ColFrom To ColTo
            Select Case Mode
            Case lgCFBackColor
               CellBackColor(lRow, lCol) = rVal(vNewValue)

            Case lgCFForeColor
               CellForeColor(lRow, lCol) = rVal(vNewValue)

            Case lgCFImage
               CellImage(lRow, lCol) = rVal(vNewValue)

            Case lgCFFontName
               If LenB(Trim$(vNewValue)) Then
                  CellFontName(lRow, lCol) = vNewValue
               End If

            Case lgCFFontBold
               CellFontBold(lRow, lCol) = rVal(vNewValue)

            Case lgCFFontItalic
               CellFontItalic(lRow, lCol) = rVal(vNewValue)

            Case lgCFFontUnderline
               CellFontUnderline(lRow, lCol) = rVal(vNewValue)
            End Select

         Next lCol
      Next lRow

   End If

End Sub

Public Sub FormatCellsAlignment(Optional ByVal RowFrom As Long = 0, _
                                Optional ByVal RowTo As Long = C_NULL_RESULT, _
                                Optional ByVal ColFrom As Long = 0, _
                                Optional ByVal ColTo As Long = C_NULL_RESULT, _
                                Optional ByVal vAlignment As lgAlignmentEnum = lgAlignLeftCenter)
      
  Dim lCol As Long
  Dim lRow As Long

   '// MLH - New
   '// Change cell alignment for a range of cells
   
   If Not (mRowCount = C_NULL_RESULT) Then '// MLH - prevent error
      If RowTo = C_NULL_RESULT Then RowTo = mRowCount
      If ColTo = C_NULL_RESULT Then ColTo = UBound(mCols)
   
      For lRow = RowFrom To RowTo
         For lCol = ColFrom To ColTo
            mItems(lRow).Cell(lCol).nAlignment = vAlignment 'mCols(lCol).nAlignment
         Next lCol
      Next lRow
      DrawGrid mbRedraw
   End If

End Sub

Public Property Get FreezeAtCol() As Long
Attribute FreezeAtCol.VB_Description = "Freeze columns 0 through ? so that they are always displayed.  Enter -1 to unfreeze."

   '// MLH - New
   FreezeAtCol = mlngFreezeAtCol

End Property

Public Property Let FreezeAtCol(ByVal vNewValue As Long)

   '// MLH - New
   mlngFreezeAtCol = vNewValue
   PropertyChanged "FreezeAtCol"

   If SBValue(efsHorizontal) < mlngFreezeAtCol + 1 Then
      SBValue(efsHorizontal) = mlngFreezeAtCol + 1
   End If

End Property

Private Function GetColFromX(ByVal X As Single) As Long

  Dim lX   As Long
  Dim lCol As Long

   '// Purpose: Return Column from mouse position

   GetColFromX = C_NULL_RESULT

   'For lCol = SBValue(efsHorizontal) To UBound(mCols) '// MLH - was before FreezeAtCol
   For lCol = 0 To UBound(mCols) '// MLH - FreezeAtCol

      If lCol <= mlngFreezeAtCol Or lCol >= SBValue(efsHorizontal) Then

         With mCols(mColPtr(lCol))
            If .bVisible Then
               If (X > lX) And (X <= lX + .lWidth + 4) Then '// MLH - added + 4
                  GetColFromX = lCol '// MLH
                  Exit For
               End If

               lX = lX + .lWidth
            End If
         End With
      End If

   Next lCol

End Function

Private Function GetColumnHeadingHeight() As Long

  Dim lHeight As Long

   '// Purpose: Return Height of Header Row

   If mbColumnHeaders Then

      lHeight = UserControl.TextHeight(C_CHECKTEXT)

      If mblnColumnHeaderSmall Then
         GetColumnHeadingHeight = lHeight + ((lHeight * mColumnHeaderLines) / 3) + mR.CaptionHeight
      Else
         GetColumnHeadingHeight = lHeight + (lHeight * mColumnHeaderLines) + mR.CaptionHeight
      End If

   Else
      GetColumnHeadingHeight = mR.CaptionHeight
   End If

End Function

Private Function GetDESKTOPDir() As String

   '// MLH - new sub
   '// Purpose: used to find desktop in ExportGrid

  Dim ItmLst  As typITEMIDLIST
  Dim strPath As String
  Dim lngRet  As Long
  Const CSIDL_DESKTOP As Long = &H0

   '// Check witch folder is chosen
   lngRet = SHGetSpecialFolderLocation(0&, CSIDL_DESKTOP, ItmLst)

   If lngRet = 0 Then '// no errors
      '// Create buffer
      strPath = Space$(255)
      '// Set strPath
      If SHGetPathFromIDList(ItmLst.mkid.cb, strPath) <> 0 Then
         '// API string sometimes end on Chr$(0) --> delete it
         strPath = Left$(strPath, InStr(strPath, vbNullChar) - 1)
      Else
         strPath = vbNullString
      End If

      '// Tell the user where the folder is located
      '// MsgBox "The folder you were looking for is " & strPath
   Else
      '// Display the error message
      MsgBox "Unknown error"
   End If

   GetDESKTOPDir = strPath

End Function

Private Function GetFlag(ByVal nFlags As Integer, ByVal nFlag As lgFlagsEnum) As Boolean

   '// Purpose: Gets information by bit flags

   If nFlags And nFlag Then
      GetFlag = True
   End If

End Function

Private Sub GetGradientColor(lhWnd As Long)

   '// MLH - new sub
   '// Purpose: Set colors based on the current windows theme in use

  Dim udtThemeID As lgThemeConst

   On Local Error Resume Next

   GetThemeName lhWnd

   If AppThemed Then '//Check if themed.

      Select Case mstrCurSysThemeName
      Case "Metallic"
         udtThemeID = silver

      Case "HomeStead"
         udtThemeID = Olive

      Case Else
         udtThemeID = Blue
      End Select

   Else
      udtThemeID = CustomTheme
   End If

   Call SetDefaultThemeColor(udtThemeID)

End Sub

Private Function GetRowFromY(ByVal Y As Single) As Long

   '// Purpose: Return Row from mouse position

  Dim lColumnHeadingHeight As Long
  Dim lRow                 As Long
  Dim lStart               As Long
  Dim lY                   As Long

   '// Are we below Header?
   If mR.HeaderHeight > 0 Then
      lColumnHeadingHeight = GetColumnHeadingHeight()

      If Y <= lColumnHeadingHeight Then
         GetRowFromY = C_NULL_RESULT
         Exit Function
      End If

   End If

   lY = lColumnHeadingHeight
   lStart = SBValue(efsVertical)

   For lRow = lStart To mRowCount
      If mItems(mRowPtr(lRow)).bVisible Then '// MLH - new
         lY = lY + mItems(mRowPtr(lRow)).lHeight
      End If
      
      If lY >= Y Then
         Exit For
      End If

   Next lRow

   If lRow <= mRowCount Then
      GetRowFromY = lRow
   Else
      GetRowFromY = C_NULL_RESULT
   End If

End Function

Private Sub GetThemeName(lngHWND As Long)

   '// MLH - new sub
   '// Purpose: Get the windows theme name in use

  Dim lngTheme         As Long
  Dim stringShellStyle As String
  Dim stringThemeFile  As String
  Dim lngPtrThemeFile  As Long
  Dim lngPtrColorName  As Long
  Dim lngPos           As Long

   On Error Resume Next

   lngTheme = OpenThemeData(lngHWND, StrPtr("ExplorerBar"))

   If Not lngTheme = 0 Then

      ReDim bThemeFile(0 To 260 * 2) As Byte
      lngPtrThemeFile = VarPtr(bThemeFile(0))

      ReDim bColorName(0 To 260 * 2) As Byte
      lngPtrColorName = VarPtr(bColorName(0))

      GetCurrentThemeName lngPtrThemeFile, 260, lngPtrColorName, 260, 0, 0
      stringThemeFile = bThemeFile
      lngPos = InStr(stringThemeFile, vbNullChar)

      If lngPos > 1 Then
         stringThemeFile = Left$(stringThemeFile, lngPos - 1)
      End If

      mstrCurSysThemeName = bColorName
      lngPos = InStr(mstrCurSysThemeName, vbNullChar)

      If lngPos > 1 Then
         mstrCurSysThemeName = Left$(mstrCurSysThemeName, lngPos - 1)
      End If

      stringShellStyle = stringThemeFile

      For lngPos = Len(stringThemeFile) To 1 Step -1

         If (Mid$(stringThemeFile, lngPos, 1) = "\") Then
            stringShellStyle = Left$(stringThemeFile, lngPos)
            Exit For
         End If

      Next lngPos

      stringShellStyle = stringShellStyle & "Shell\" & mstrCurSysThemeName & "\ShellStyle.dll"
      CloseThemeData lngTheme

   Else
      mstrCurSysThemeName = "Classic"

   End If

   On Error GoTo 0

End Sub

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets a value that determines the Color of grid background"

   GridColor = mGridColor

End Property

Public Property Let GridColor(ByVal vNewValue As OLE_COLOR)

   mGridColor = vNewValue
   PropertyChanged "GridColor"
   DrawGrid mbRedraw

End Property

Public Property Get GridLines() As lgGridLinesEnum
Attribute GridLines.VB_Description = "Returns/sets a value that determines which grid lines as visible (Horizontal, Vertical, Both, None)"

   GridLines = muGridLines

End Property

Public Property Let GridLines(ByVal vNewValue As lgGridLinesEnum)

   muGridLines = vNewValue
   PropertyChanged "GridLines"
   DisplayChange

End Property

Public Property Get GridLineWidth() As Long
Attribute GridLineWidth.VB_Description = "Returns/sets a value that determines the Width of grid lines"

   GridLineWidth = mGridLineWidth

End Property

Public Property Let GridLineWidth(ByVal vNewValue As Long)

   mGridLineWidth = vNewValue
   PropertyChanged "GridLineWidth"
   DrawGrid mbRedraw

End Property

Public Property Get ImageList() As Object
Attribute ImageList.VB_Description = "Returns/sets a value that determines the Name of ImageList control used to add images to grid"

   Set ImageList = moImageList

End Property

Public Property Let ImageList(ByVal vNewValue As Object)

   On Error Resume Next

   Set moImageList = vNewValue

   If Not moImageList Is Nothing Then
      'mImageListScaleMode = moImageList.Parent.ScaleMode
      mImageListScaleMode = UserControl.Parent.ScaleMode
   End If

   DisplayChange

End Property

Private Function InvertThisColor(ByVal oInsColor As OLE_COLOR)

   '---------------------------------------------------------------------------------
   'Source: Riccardo Cohen
   '---------------------------------------------------------------------------------

  Dim lROut As Long
  Dim lGOut As Long
  Dim lBOut As Long
  Dim lRGB  As Long

   lRGB = TranslateColor(oInsColor)

   lROut = (255 - (lRGB And &HFF&))
   lGOut = (255 - ((lRGB And &HFF00&) / &H100))
   lBOut = (255 - ((lRGB And &HFF0000) / &H10000))
   InvertThisColor = RGB(lROut, lGOut, lBOut)

End Function

Private Function IsColumnTruncated(ByVal vCol As Long) As Boolean

   If (mR.LeftText > C_TEXT_SPACE) And (vCol = 0) Then
      IsColumnTruncated = True
   End If

End Function

Private Function IsEditable() As Boolean

   'If Not mLocked And mbAllowEdit Then '// MLH - mLocked Not Used

   If mbAllowEdit Then
      IsEditable = (mRowCount >= 0)
   End If

End Function

Private Function IsInIDE() As Boolean

   '// Return whether we're running in the IDE.

   '// Assert invocations work only within the development environment and
   '// conditionally suspends execution (if set to False) at the line on which
   '// the method appears.
   '// When the module is compiled into an executable, the method calls on the
   '// Debug object are omitted.

   Debug.Assert IsInIDE_SetTrue(IsInIDE)

End Function

Private Function IsInIDE_SetTrue(ByRef bValue As Boolean) As Boolean

   '// Worker function for IsInIDE
   IsInIDE_SetTrue = True
   bValue = True

End Function

Private Function IsValidRowCol(ByVal vRow As Long, ByVal vCol As Long) As Boolean

   IsValidRowCol = (vRow > C_NULL_RESULT) And (vCol > C_NULL_RESULT)

End Function

Public Property Get ItemCount() As Long

   ItemCount = mRowCount + 1

End Property

Private Function LongToSignedShort(ByVal dwUnsigned As Long) As Integer

   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If

End Function

Public Property Get MaxLineCount() As Long
Attribute MaxLineCount.VB_Description = "Returns/sets a value that determines the maxium the number of lines that will be displayed when a cell is word wrapped (0=no limit)"

   '// MLH - renamed
   MaxLineCount = mMaxLineCount

End Property

Public Property Let MaxLineCount(ByVal vNewValue As Long)

   mMaxLineCount = vNewValue
   PropertyChanged "MaxLineCount"
   DisplayChange

End Property

Public Property Get MinRowHeight() As Long
Attribute MinRowHeight.VB_Description = "Returns/sets a value that determines the Minimum height of rows"

   MinRowHeight = mMinRowHeightUser

End Property

Public Property Let MinRowHeight(ByVal vNewValue As Long)

   '// MLH - renamed
   mMinRowHeightUser = vNewValue
   mMinRowHeight = 0
   PropertyChanged "MinRowHeight"
   Call CreateRenderData '// MLH - Update rendered data
   DisplayChange

End Property

Public Property Get MinVerticalOffset() As Long
Attribute MinVerticalOffset.VB_Description = "Returns/sets a value that determines the space between cell text and grid lines. A value of 2 will add 2 pixels above and below the text, therefore increasing the minimum row height by 4 pixels."

   MinVerticalOffset = mMinVerticalOffset

End Property

Public Property Let MinVerticalOffset(ByVal vNewValue As Long)

   '// MLH - new sub
   '// Purpose: add vertical offset from grid lines

   mMinVerticalOffset = vNewValue
   mMinRowHeight = 0
   PropertyChanged "MinVerticalOffset"
   Call CreateRenderData '// save changes
   DisplayChange

End Property

Public Property Get MouseCol() As Long

   If Not (mMouseCol = C_NULL_RESULT) Then
      MouseCol = mColPtr(mMouseCol)
   Else
      MouseCol = C_NULL_RESULT
   End If

End Property

Public Property Get MouseRow() As Long

   MouseRow = mMouseRow

End Property

Private Sub MoveEditControl(ByVal MoveControl As lgMoveControlEnum)

   '// Purpose: Used to position and optionally resize the Edit control.

  Dim r            As Rect
  Dim lBorderWidth As Long
  Dim nScaleMode   As ScaleModeConstants
  Dim lHeight      As Long

   SetColRect mEditCol, r

   If Not IsColumnTruncated(mEditCol) Then
      r.Left = r.Left + mGridLineWidth
   End If

   On Local Error Resume Next

   '// Check if an external Control is used.
   
   If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
      '// Using internal TextBox
      With txtEdit
         .Left = r.Left
         .Top = RowTop(mEditRow) + mGridLineWidth
         .Height = mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth
         .Width = (r.Right - r.Left)
      End With

   Else '// External Control
      nScaleMode = UserControl.Parent.ScaleMode

      If muBorderStyle = lgSingle Then
         lBorderWidth = 2
      End If

      '// Is VB.ComboBox
      If (TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox) Then
         With mCols(mColPtr(mEditCol)).EditCtrl

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
               .Left = ScaleX(r.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
               .Top = ScaleY(RowTop(mEditRow) + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
               .Width = ScaleX((r.Right - r.Left), vbPixels, nScaleMode)
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
               'lHeight = mMinRowHeight - mGridLineWidth - 4 '// MLH - changed
               lHeight = mItems(mRowPtr(mEditRow)).lHeight - (mGridLineWidth * 2) '// MLH - New
               Call SendMessageAsLong(.hWnd, CB_SETITEMHEIGHT, -1, ByVal lHeight)
               '// MLH - removed Don't adjust the height of the items in the drop down
               'Call SendMessageAsLong(.hWnd, CB_SETITEMHEIGHT, 0, ByVal lHeight)
            End If

         End With

      Else '// Is NOT VB.ComboBox
         With mCols(mColPtr(mEditCol)).EditCtrl

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
               .Left = ScaleX(r.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
               .Top = ScaleY(RowTop(mEditRow) + mGridLineWidth + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
               .Height = ScaleY(mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth, vbPixels, nScaleMode)
            End If

            If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
               .Width = ScaleX((r.Right - r.Left), vbPixels, nScaleMode)
            End If

         End With
      End If
   End If

   On Local Error GoTo 0

End Sub

Public Property Get MultiSelect() As lgMultiSelectEnum
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether multiple row selection is allowed"

   '// MLH - Added New Options
   MultiSelect = muMultiSelect

End Property

Public Property Let MultiSelect(ByVal vNewValue As lgMultiSelectEnum)

   muMultiSelect = vNewValue
   PropertyChanged "MultiSelect"

   If vNewValue = lgSingleSelect Then
      SetSelection False
      RowColSet
      DisplayChange
   End If

End Property

Private Function NavigateDown(Optional ByVal lRow As Long = 1, _
                              Optional ByVal vbVisibleOnly As Boolean = False) As Long

  Dim bSkip As Boolean
  
   If mRow < mRowCount Then
      NavigateDown = mRow + lRow
   Else
      NavigateDown = mRow
   End If

   '// MLH - Prevent locked & invisible rows from getting focus
   bSkip = (Not vbVisibleOnly And (mItems(mRowPtr(NavigateDown)).nFlags And lgFLlocked)) Or Not mItems(mRowPtr(NavigateDown)).bVisible
   If bSkip Then
   
      Do
         NavigateDown = NavigateDown + lRow

         If NavigateDown > mRowCount Then
            NavigateDown = mRow
            Exit Do
         End If
         bSkip = (Not vbVisibleOnly And (mItems(mRowPtr(NavigateDown)).nFlags And lgFLlocked)) Or Not mItems(mRowPtr(NavigateDown)).bVisible
      Loop Until Not bSkip
   End If

End Function

Private Function NavigateLeft(Optional bSkipChange As Boolean = False) As Long

  Dim lngI    As Long
  Dim lMaxCol As Long

   lMaxCol = UBound(mCols) '// MLH

   If bSkipChange Then '// MLH
      NavigateLeft = lMaxCol

   Else
      If mCol > 0 Then
         NavigateLeft = mCol - 1
      Else
         NavigateLeft = lMaxCol
      End If
   End If

   '// MLH - Prevent locked columns from getting focus
   If Not mCols(mColPtr(NavigateLeft)).bVisible Or mCols(mColPtr(NavigateLeft)).bLocked Then
      lngI = NavigateLeft

      Do
         NavigateLeft = NavigateLeft - 1
         If NavigateLeft < 0 Then NavigateLeft = lMaxCol
         If lngI = NavigateLeft Then Exit Do '// MLH - just in case all col are locked
      Loop Until mCols(mColPtr(NavigateLeft)).bVisible And Not mCols((NavigateLeft)).bLocked

   End If

End Function

Private Function NavigateRight(Optional bSkipChange As Boolean = False) As Long

  Dim lngI    As Long
  Dim lMaxCol As Long

   lMaxCol = UBound(mCols) '// MLH

   If bSkipChange Then '// MLH
      NavigateRight = 0

   Else
      If mCol < lMaxCol Then
         NavigateRight = mCol + 1
      Else
         NavigateRight = 0
      End If
   End If

   '// MLH - Prevent locked columns from getting focus
   If Not mCols(mColPtr(NavigateRight)).bVisible Or mCols(mColPtr(NavigateRight)).bLocked Then
      lngI = NavigateRight

      Do
         NavigateRight = NavigateRight + 1
         If NavigateRight > lMaxCol Then NavigateRight = 0
         If lngI = NavigateRight Then Exit Do '// MLH - just in case all col are locked
      Loop Until mCols(mColPtr(NavigateRight)).bVisible And Not mCols(mColPtr(NavigateRight)).bLocked

   End If

End Function

Private Function NavigateUp(Optional ByVal lRow As Long = 1, _
                            Optional ByVal vbVisibleOnly As Boolean = False) As Long

  Dim bSkip As Boolean
  
   If mRow > 0 Then
      NavigateUp = mRow - lRow
      If NavigateUp < 0 Then NavigateUp = 0
   Else
      NavigateUp = mRow
   End If

   '// MLH - Prevent locked & invisible rows from getting focus
   bSkip = (Not vbVisibleOnly And (mItems(mRowPtr(NavigateUp)).nFlags And lgFLlocked)) Or Not mItems(mRowPtr(NavigateUp)).bVisible
   If bSkip Then

      Do
         NavigateUp = NavigateUp - lRow

         If NavigateUp < 0 Then
            NavigateUp = mRow
            Exit Do
         End If
         bSkip = (Not vbVisibleOnly And (mItems(mRowPtr(NavigateUp)).nFlags And lgFLlocked)) Or Not mItems(mRowPtr(NavigateUp)).bVisible
      Loop Until Not bSkip
   End If

End Function

Private Sub pSBClearUp()

   If Not (mSBhWnd = 0) Then
      On Error Resume Next
      '// Stop flat scroll bar if we have it:

      If Not (mbSBNoFlatScrollBars) Then
         UninitializeFlatSB mSBhWnd
      End If

   End If

   mSBhWnd = 0
   'mbSBInitialised = False '// MLH - Not Used

End Sub

Private Sub pSBCreateScrollBar()

   On Error Resume Next

   Call InitialiseFlatSB(mSBhWnd)

   If Not (Err.Number = 0) Or muScrollBarStyle = Style_Regular Then
      '// Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
      '//  Means we have version prior to 4.71
      '//  We get standard scroll bars.
      mbSBNoFlatScrollBars = True

   Else
      SBStyle = muSBStyle
   End If

End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)

  Dim Lo As Long

   Lo = eBar
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)

   If (mbSBNoFlatScrollBars) Then
      GetScrollInfo mSBhWnd, Lo, tSI
   Else
      FlatSB_GetScrollInfo mSBhWnd, Lo, tSI
   End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)

  Dim Lo As Long

   Lo = eBar
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)

   If (mbSBNoFlatScrollBars) Then
      SetScrollInfo mSBhWnd, Lo, tSI, True
   Else
      FlatSB_SetScrollInfo mSBhWnd, Lo, tSI, True
   End If

End Sub

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Allow grid to update.  Turned this off when data is being added to prevent multiple redraws and increase speed. (True/False).  This is False by default and must be turned on."

   Redraw = mbUserRedraw

End Property

Public Property Let Redraw(ByVal vNewValue As Boolean)

   mbRedraw = vNewValue
   mbUserRedraw = mbRedraw
   PropertyChanged "Redraw"

   If mbRedraw Then
      If mbPendingScrollBar Then
         SetScrollBars
      End If

      If mbPendingRedraw Then
         CreateRenderData
         DrawGrid mbRedraw
      End If

   Else
      mbPendingScrollBar = False
      mbPendingRedraw = False
   End If

   If SBValue(efsHorizontal) < mlngFreezeAtCol + 1 Then '// MLH - FreezeAtCol
      SBValue(efsHorizontal) = mlngFreezeAtCol + 1
   End If

End Property

Public Sub Refresh()

   CreateRenderData
   SetScrollBars
   DrawGrid True

End Sub

Public Sub RemoveItem(Optional ByVal vRow As Long = C_NULL_RESULT)

   Call RemoveRow(vRow)

End Sub

Public Sub RemoveRow(Optional ByVal vRow As Long = C_NULL_RESULT)

  Dim lCount    As Long
  Dim lPosition As Long
  Dim bSelected As Boolean

   '---------------------------------------------------------------------------------
   'See AddItem for details of the Arrays used
   '---------------------------------------------------------------------------------

   If (mRowCount = C_NULL_RESULT) Then '// MLH - prevent error

      If IsInIDE Then
         MsgBox "No Rows Added!", vbExclamation, "ERROR - RemoveItem"
      End If

      Exit Sub
   End If

   If vRow = C_NULL_RESULT Then vRow = mRow '// MLH

   '// Note selected state before deletion
   bSelected = mItems(mRowPtr(vRow)).nFlags And lgFLSelected
   '// Note visible start before deletion
   '// Decrement the reference count on each cells format Entry
   If mRowCount >= 0 Then

      For lCount = 0 To UBound(mCols)
         If mItems(vRow).Cell(lCount).nFormat >= 0 Then '// MLH - Typo Bug Fix lCount was Count
            mCF(mItems(vRow).Cell(lCount).nFormat).lRefCount = mCF(mItems(vRow).Cell(lCount).nFormat).lRefCount - 1
         End If

         mudtTotalsVal(mColPtr(lCount)) = mudtTotalsVal(mColPtr(lCount)) - rVal(mItems(mRowPtr(vRow)).Cell(lCount).sValue) '// MLH - New

      Next lCount
   End If

   lPosition = mRowPtr(vRow)

   '// Reset Item Data
   For lCount = mRowPtr(vRow) To mRowCount - 1
      mItems(lCount) = mItems(lCount + 1)
   Next lCount

   '// Adjust vRow
   For lCount = vRow To mRowCount - 1
      mRowPtr(lCount) = mRowPtr(lCount + 1)
   Next lCount

   '// Validate Indexes for Items after deleted Item
   For lCount = 0 To mRowCount - 1

      If mRowPtr(lCount) > lPosition Then
         mRowPtr(lCount) = mRowPtr(lCount) - 1
      End If

   Next lCount

   '// Adjust Row Counts
   mRowCount = mRowCount - 1
   If mRowCount < 0 Then
      Call Clear
   
   Else
      If (mRowCount + mCacheIncrement) < UBound(mItems) Then
         ReDim Preserve mItems(mRowCount) As udtItem
         ReDim Preserve mRowPtr(mRowCount) As Long
      End If

      If bSelected Then
         If muMultiSelect Then
            RaiseEvent SelectionChanged

         ElseIf vRow > mRowCount Then
            SetFlag mItems(mRowPtr(mRowCount)).nFlags, lgFLSelected, True

         ElseIf mRowCount >= 0 Then
            SetFlag mItems(mRowPtr(vRow)).nFlags, lgFLSelected, True
         End If

      End If

      If vRow > mRowCount Then
         SetRowCol mRow - 1, mCol
      End If

   End If

   SetRedrawState True
   DisplayChange

   RaiseEvent RowCountChanged

End Sub

Private Function ReturnAddr(ByVal sDLL As String, ByVal sProc As String) As Long

   '// Return the address of the specified DLL/procedure
   '// Get the specified procedure address

   If mbWinNT Then
      ReturnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)
   Else
      ReturnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
   End If

   '// In the IDE, validate that the procedure address was located
   Debug.Assert ReturnAddr

End Function

Public Property Get Row() As Long
Attribute Row.VB_Description = "Returns/sets a value that determines the selected row"

   Row = mRow

End Property

Public Property Let Row(ByVal vRow As Long)

   Call RowColSet(vRow) '// MLH

End Property

Public Function RowAdded(ByVal vRow As Long) As Boolean

   '// MLH - new sub
   RowAdded = mItems(mRowPtr(vRow)).nFlags And lgFLNewRow

End Function

Public Property Let RowBackColor(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Long)

   '// MLH Renamed
  Dim lCol As Long

   If FixRef(vRow) Then
      For lCol = LBound(mCols) To UBound(mCols)
         CellBackColor(mRowPtr(vRow), lCol) = vNewValue
      Next lCol
   
      DrawGrid mbRedraw
   End If
   
End Property

Public Function RowChanged(ByVal vRow As Long) As Boolean

   '// MLH - new sub
   RowChanged = mItems(mRowPtr(vRow)).nFlags And lgFLChanged

End Function

Public Property Get RowCheckBoxes() As Boolean
Attribute RowCheckBoxes.VB_Description = "Returns/sets a value that determines whether Row checkboxs are visible"

   '// MLH - renamed
   RowCheckBoxes = mbCheckboxes

End Property

Public Property Let RowCheckBoxes(ByVal vNewValue As Boolean)

   mbCheckboxes = vNewValue
   PropertyChanged "CheckBoxes"
   DisplayChange

End Property

Public Property Get RowChecked(Optional ByVal vRow As Long = C_NULL_RESULT) As Boolean

   '// MLH Renamed
   If FixRef(vRow) Then
      RowChecked = mItems(mRowPtr(vRow)).nFlags And lgFLChecked
   End If
   
End Property

Public Property Let RowChecked(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   If FixRef(vRow) Then
      SetFlag mItems(mRowPtr(vRow)).nFlags, lgFLChecked, vNewValue
      DrawGrid mbRedraw
   End If
   
End Property

Public Sub RowColSet(Optional ByVal lNewRow As Long = C_NULL_RESULT, _
                     Optional ByVal lNewCol As Long = C_NULL_RESULT)

   '// MLH - New
  Dim bRedraw  As Boolean
  Dim blnColOk As Boolean

   If lNewCol = C_NULL_RESULT Then
      blnColOk = True
   
   ElseIf mCols(mColPtr(lNewCol)).bVisible Then
      If lNewCol > UBound(mCols) Then lNewCol = UBound(mCols)
      blnColOk = True
      lNewCol = mColPtr(lNewCol) '// FIND MOVED COLUMN NUMBER
   End If

   If FixRef(lNewRow) Then
      If mItems(mRowPtr(lNewRow)).bVisible And blnColOk Then
      
         If lNewRow <= mRowCount And lNewRow >= 0 Then
            If Not (mRow = lNewRow) Then
               bRedraw = SetSelection(False)
            End If
   
            If lNewRow > C_NULL_RESULT Then
               Call SetRowCol(lNewRow, lNewCol)
   
               If Not mItems(mRowPtr(lNewRow)).nFlags And lgFLSelected Then
                  bRedraw = True
                  SetFlag mItems(mRowPtr(lNewRow)).nFlags, lgFLSelected, True
                  RaiseEvent SelectionChanged
               End If
            End If
   
            If bRedraw Or SetRowCol(lNewRow, lNewCol, True) Then
               mRow = lNewRow
               If lNewRow < SBValue(efsVertical) Then
                  SBValue(efsVertical) = mRow
               ElseIf lNewRow > (SBValue(efsVertical) + RowsVisible()) Then
                  SBValue(efsVertical) = mRow
               End If
   
               DrawGrid mbRedraw
            End If
      
         End If
      End If
   End If

End Sub

Public Property Get RowData(Optional ByVal vRow As Long = C_NULL_RESULT) As Long

   '// MLH Renamed
   If FixRef(vRow) Then
      RowData = mItems(mRowPtr(vRow)).lItemData
   End If
   
End Property

Public Property Let RowData(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Long)

   If FixRef(vRow) Then
      mItems(mRowPtr(vRow)).lItemData = vNewValue
   End If
   
End Property

Public Property Let RowFontBold(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   '// MLH Renamed
  Dim lCol As Long

   If FixRef(vRow) Then
      For lCol = LBound(mCols) To UBound(mCols)
         CellFontBold(mRowPtr(vRow), lCol) = vNewValue
      Next lCol
   
      DrawGrid mbRedraw
   End If
   
End Property

Public Property Let RowForeColor(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Long)

   '// MLH Renamed
  Dim lCol As Long
   
   If FixRef(vRow) Then
      For lCol = LBound(mCols) To UBound(mCols)
         CellForeColor(mRowPtr(vRow), lCol) = vNewValue
      Next lCol
   
      DrawGrid mbRedraw
   End If
   
End Property

Public Property Get RowGroupHeader(Optional ByVal vRow As Long = C_NULL_RESULT) As Boolean

   '// MLH- New
   If FixRef(vRow) Then
      RowGroupHeader = mItems(mRowPtr(vRow)).bGroupRow
   End If
   
End Property

Public Property Let RowGroupHeader(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   If FixRef(vRow) Then
      mItems(mRowPtr(vRow)).bGroupRow = vNewValue
   End If
   
End Property

Public Property Get RowHeight(Optional ByVal vRow As Long = C_NULL_RESULT) As Single

   If FixRef(vRow) Then
      RowHeight = mItems(mRowPtr(vRow)).lHeight
   End If
   
End Property

Public Property Let RowHeight(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Single)

   If FixRef(vRow) Then
      If vNewValue = C_NULL_RESULT Then
         SetRowSize mRow
      Else
         mItems(mRowPtr(vRow)).lHeight = vNewValue
      End If
   
      SetScrollBars
      DrawGrid mbRedraw
   End If
   
End Property

Public Property Get RowImage(Optional ByVal vRow As Long = C_NULL_RESULT) As Variant

   '// MLH Renamed
   If FixRef(vRow) Then
      If mItems(mRowPtr(vRow)).lImage >= 0 Then
         RowImage = mItems(mRowPtr(vRow)).lImage
      Else
         RowImage = moImageList.ListImages(Abs(mItems(mRowPtr(vRow)).lImage)).Key
      End If
   End If
   
End Property

Public Property Let RowImage(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Variant)

   On Local Error GoTo ItemImageError

   If FixRef(vRow) Then
      If IsNumeric(vNewValue) Then
         mItems(mRowPtr(vRow)).lImage = vNewValue
      Else
         mItems(mRowPtr(vRow)).lImage = -moImageList.ListImages(vNewValue).Index
      End If
   
      DrawGrid mbRedraw
   End If
   
   Exit Property

ItemImageError:
   mItems(mRowPtr(vRow)).lImage = 0

End Property

Public Property Get RowLocked(Optional ByVal vRow As Long = C_NULL_RESULT) As Boolean

   '// MLH - New
   If FixRef(vRow) Then
      If mLRLocked Then
         RowLocked = True
      Else
         RowLocked = mItems(mRowPtr(vRow)).nFlags And lgFLlocked
      End If
   Else
      RowLocked = True
   End If
   
End Property

Public Property Let RowLocked(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   '// MLH - New
   If FixRef(vRow) Then
      SetFlag mItems(mRowPtr(vRow)).nFlags, lgFLlocked, vNewValue
      DrawGrid mbRedraw
   End If
   
End Property

Public Property Get Rows() As Long

   '// MLH - new sub
   Rows = mRowCount + 1

End Property

Public Property Get RowSelected(Optional ByVal vRow As Long = C_NULL_RESULT) As Boolean

   '// MLH Renamed
   If FixRef(vRow) Then
      RowSelected = mItems(mRowPtr(vRow)).nFlags And lgFLSelected
   End If

End Property

Public Property Let RowSelected(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   '// MLH Renamed
   If FixRef(vRow) Then
      SetFlag mItems(mRowPtr(vRow)).nFlags, lgFLSelected, vNewValue
      DrawGrid mbRedraw
   End If
   
End Property

Private Function RowsVisible() As Long

   '// MLH - renamed
  Dim lBorderWidth As Long

   If muBorderStyle = lgSingle Then
      lBorderWidth = 2
   End If

   With UserControl
      RowsVisible = (.ScaleHeight - GetColumnHeadingHeight() - (lBorderWidth * 2)) / mMinRowHeight
   End With

End Function

Public Property Get RowTag(Optional ByVal vRow As Long = C_NULL_RESULT) As String

   If FixRef(vRow) Then
      RowTag = mItems(mRowPtr(vRow)).sTag
   End If
   
End Property

Public Property Let RowTag(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As String)

   '// MLH Renamed
   If FixRef(vRow) Then
      mItems(mRowPtr(vRow)).sTag = vNewValue
   End If
   
End Property

Private Function RowTop(ByVal Index As Long, Optional ByVal lStart As Long = C_NULL_RESULT) As Long

  '// MLH - was public
  Dim lRow   As Long
  'Dim lStart As Long '// MLH
  Dim lY     As Long

   If lStart = C_NULL_RESULT Then
      lStart = SBValue(efsVertical)
   End If

   If Index >= lStart Then
      lY = GetColumnHeadingHeight()

      For lRow = lStart To Index - 1
         If mItems(mRowPtr(lRow)).bVisible Then '// MLH - new
            lY = lY + mItems(mRowPtr(lRow)).lHeight
         End If
      Next lRow

   Else
      lY = C_NULL_RESULT
   End If

   RowTop = lY

End Function

Public Property Get RowVisible(Optional ByVal vRow As Long = C_NULL_RESULT) As Boolean

   '// MLH - new
   If FixRef(vRow) Then
      RowVisible = mItems(mRowPtr(vRow)).bVisible
   End If
   
End Property

Public Property Let RowVisible(Optional ByVal vRow As Long = C_NULL_RESULT, ByVal vNewValue As Boolean)

   '// MLH - new
   If FixRef(vRow) Then
      If Not mItems(mRowPtr(vRow)).bVisible = vNewValue Then
         mItems(mRowPtr(vRow)).bVisible = vNewValue
         Call SetScrollBars
         Call DrawGrid(mbRedraw)
      End If
   End If
   
End Property

Public Function rVal(ByVal vString As String) As Double

   '// MLH - New
   '// Returns the numbers contained in a string as a numeric value
   '// The Val function recognizes only the period (.) as a valid decimal separator.
   '// The CDbl errors on empty strings or values containing non-numeric values
   '// Returns the numbers contained in a string as a numeric value

  Dim lngI     As Long
  Dim lngS     As Long
  Dim bytAscV  As Byte
  Dim strTemp  As String
  
  On Error Resume Next

   vString = Trim$(UCase$(vString))
   If LenB(vString) Then
   
      If Left$(vString, 4) = "TRUE" Then        '// Boolean ?
         rVal = True
         
      ElseIf Left$(vString, 5) = "FALSE" Then   '// Boolean ?
         rVal = False
      
      Else
         Select Case Left$(vString, 2)          '// Hex or Octal?
         Case Is = "&H", Is = "&O"
            lngS = 3
            strTemp = Left$(vString, 2)
         Case Else
            lngS = 1
         End Select
         
         For lngI = lngS To Len(vString)
            bytAscV = Asc(Mid$(vString, lngI, 1))
            Select Case bytAscV
            Case 48 To 57 '// 1234567890
               strTemp = strTemp & Mid$(vString, lngI, 1)
            
            Case 44, 45, 46 '// , - .
               strTemp = strTemp & Mid$(vString, lngI, 1)
            
            Case 36, 163, 32 '// $
               '// Ignore
               
            Case Is > 57, Is < 44
               If Left$(strTemp, 2) = "&H" Then '// Hex Values ?
                  Select Case bytAscV
                  Case 65 To 70 '// ABCDEF
                     strTemp = strTemp & Mid$(vString, lngI, 1)
                  Case Else
                     Exit For
                  End Select
               Else
                  Exit For
               End If
            End Select
         Next lngI
         
         If LenB(strTemp) Then
            rVal = CDbl(strTemp)
            If rVal = 0 Then
               strTemp = Replace$(strTemp, ".", ",")
               rVal = CDbl(strTemp)
            End If
         End If
      End If
   
   Else
      rVal = 0
   End If
   
   On Error GoTo 0

End Function

Private Property Get SBCanBeFlat() As Boolean

   SBCanBeFlat = Not (mbSBNoFlatScrollBars)

End Property

Private Sub SBCreate(ByVal vhWnd As Long)

   pSBClearUp
   mSBhWnd = vhWnd
   pSBCreateScrollBar

End Sub

Private Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long

  Dim tSI As SCROLLINFO

   pSBGetSI eBar, tSI, SIF_PAGE
   SBLargeChange = tSI.nPage

End Property

Private Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)

  Dim tSI As SCROLLINFO

   pSBGetSI eBar, tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE

End Property

Private Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long

  Dim tSI As SCROLLINFO

   pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
   SBMax = tSI.nMax  '- tSI.nPage

End Property

Private Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)

  Dim tSI As SCROLLINFO

   tSI.nMax = iMax + SBLargeChange(eBar)
   tSI.nMin = SBMin(eBar)
   pSBLetSI eBar, tSI, SIF_RANGE

End Property

Private Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long

  Dim tSI As SCROLLINFO

   pSBGetSI eBar, tSI, SIF_RANGE
   SBMin = tSI.nMin

End Property

Private Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)

  Dim tSI As SCROLLINFO

   tSI.nMin = iMin
   tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
   pSBLetSI eBar, tSI, SIF_RANGE

End Property

Private Property Get SBStyle() As ScrollBarStyleEnum

   SBStyle = muSBStyle

End Property

Private Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)

   If Not mbSBNoFlatScrollBars Then
      If (muSBOrienation = Scroll_Horizontal) Or (muSBOrienation = Scroll_Both) Then
         Call FlatSB_SetScrollProp(mSBhWnd, WSB_PROP_HSTYLE, eStyle, True)
      End If

      If (muSBOrienation = Scroll_Vertical) Or (muSBOrienation = Scroll_Both) Then
         Call FlatSB_SetScrollProp(mSBhWnd, WSB_PROP_VSTYLE, eStyle, True)
      End If

      muSBStyle = eStyle

      'Else
      '// Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71
   End If

End Property

Private Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long

  Dim tSI As SCROLLINFO

   pSBGetSI eBar, tSI, SIF_POS
   SBValue = tSI.nPos

End Property

Private Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)

  Dim tSI As SCROLLINFO

   If SBVisible(eBar) Then

      If eBar = efsHorizontal Then '// MLH - FreezeAtCol
         If iValue <= mlngFreezeAtCol Then
            Exit Property
         End If
      End If

      If Not (iValue = SBValue(eBar)) Then
         tSI.nPos = iValue
         pSBLetSI eBar, tSI, SIF_POS
      End If

   End If

End Property

Private Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean

   If (eBar = efsHorizontal) Then
      SBVisible = mbSBVisibleHorz
   Else
      SBVisible = mbSBVisibleVert
   End If

End Property

Private Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)

   If (eBar = efsHorizontal) Then
      mbSBVisibleHorz = bState
   Else
      mbSBVisibleVert = bState
   End If

   If (mbSBNoFlatScrollBars) Then
      ShowScrollBar mSBhWnd, eBar, Abs(bState)
   Else
      FlatSB_ShowScrollBar mSBhWnd, eBar, Abs(bState)
   End If

End Property

Public Property Get ScaleUnits() As ScaleModeConstants
Attribute ScaleUnits.VB_Description = "User defined scale units"

   ScaleUnits = muScaleUnits

End Property

Public Property Let ScaleUnits(ByVal vNewValue As ScaleModeConstants)

   muScaleUnits = vNewValue
   PropertyChanged "ScaleUnits"

End Property

Public Property Get ScrollBars() As ScrollBarOrienationEnum
Attribute ScrollBars.VB_Description = "Returns/sets a value that determines which scroll bars are visible"

   ScrollBars = muSBOrienation

End Property

Public Property Let ScrollBars(ByVal vNewValue As ScrollBarOrienationEnum)

   muSBOrienation = vNewValue
   PropertyChanged "ScrollBars"

End Property

Public Property Get ScrollBarStyle() As ScrollBarStyleEnum
Attribute ScrollBarStyle.VB_Description = "Returns/sets a value that determines type (Flat or Regular)"

   ScrollBarStyle = muScrollBarStyle

End Property

Public Property Let ScrollBarStyle(ByVal vNewValue As ScrollBarStyleEnum)

   '// MLH - New
   muScrollBarStyle = vNewValue
   PropertyChanged "ScrollBarStyle"

   SBVisible(efsHorizontal) = False
   SBVisible(efsVertical) = False

   UninitializeFlatSB mSBhWnd
   Call pSBCreateScrollBar

   Select Case muSBOrienation
   Case Scroll_Vertical
      SBVisible(efsVertical) = True

   Case Scroll_Horizontal
      SBVisible(efsHorizontal) = True

   Case Else
      SBVisible(efsHorizontal) = True
      SBVisible(efsVertical) = True
   End Select

End Property

Private Sub ScrollList(nDirection As Integer)

   '// Purpose: Used to automatically scroll the list up or down when the mouse
   '// is dragged out of the Control

  Dim lCount        As Long
  Dim lRowsVisible  As Long

   mScrollAction = nDirection

   Do While mScrollAction = nDirection
      mScrollTick = GetTickCount()

      If nDirection = C_SCROLL_UP Then
         If SBValue(efsVertical) > SBMin(efsVertical) Then
            mRow = SBValue(efsVertical)
            SBValue(efsVertical) = NavigateUp() '// MLH

            If muMultiSelect Then
               SetFlag mItems(mRowPtr(SBValue(efsVertical))).nFlags, lgFLSelected, True

            Else
               mRow = SBValue(efsVertical)
               SetSelection False
               SetSelection True, mRow, mRow
            End If

            RaiseEvent RowColChanged
         Else
            Exit Do
         End If

      Else
         If SBValue(efsVertical) < SBMax(efsVertical) Then
            lRowsVisible = RowsVisible()
            mRow = SBValue(efsVertical)
            SBValue(efsVertical) = NavigateDown()

            If muMultiSelect Then

               For lCount = SBValue(efsVertical) To SBValue(efsVertical) + lRowsVisible
                  If lCount > mRowCount Then
                     Exit For
                  Else
                     SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, True
                  End If
               Next lCount

            Else
               mRow = SBValue(efsVertical) + (lRowsVisible - 1)

               If mRow > mRowCount Then
                  mRow = mRowCount
               End If

               SetSelection False
               SetSelection True, mRow, mRow
            End If

            RaiseEvent RowColChanged
         Else
            Exit Do
         End If

      End If

      RaiseEvent SelectionChanged
      RaiseEvent Scroll
      DrawGrid mbRedraw

      Sleep C_AUTOSCROLL_TIMEOUT
      DoEvents
   Loop

End Sub

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets a value that determines how the Grid moves with the scroll bars"

   ScrollTrack = mbScrollTrack

End Property

Public Property Let ScrollTrack(ByVal vNewValue As Boolean)

   mbScrollTrack = vNewValue
   PropertyChanged "ScrollTrack"

End Property

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      If When And MSG_BEFORE Then                  'If the message is to be added to the before original WndProc table...
         zAddMsg uMsg, IDX_BTABLE                  'Add the message to the before table
      End If

      If When And MSG_AFTER Then                   'If message is to be added to the after original WndProc table...
         zAddMsg uMsg, IDX_ATABLE                  'Add the message to the after table
      End If

   End If

End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long

   'Ensure that the thunk hasn't already released its memory
   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
      'Call the original WndProc of the passed window handle parameter
      sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
   End If

End Function

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      If When And MSG_BEFORE Then                  'If the message is to be deleted from the before original WndProc table...
         zDelMsg uMsg, IDX_BTABLE                  'Delete the message from the before table
      End If

      If When And MSG_AFTER Then                   'If the message is to be deleted from the after original WndProc table...
         zDelMsg uMsg, IDX_ATABLE                  'Delete the message from the after table
      End If

   End If

End Sub

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long

   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      sc_lParamUser = zData(IDX_PARM_USER)         'Get the lParamUser callback parameter
   End If

End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)

   If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then   'Ensure that the thunk hasn't already released its memory
      zData(IDX_PARM_USER) = NewValue              'Set the lParamUser callback parameter
   End If

End Property

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = False) As Boolean 'Subclass the specified window handle

   '*************************************************************************************************
   '* lng_hWnd   - Handle of the window to subclass
   '* lParamUser - Optional, user-defined callback parameter
   '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
   '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
   '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
   '*************************************************************************************************
  Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
  Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
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

   If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
      Set oCallback = Me                                                      'Then it is me
   End If

   nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
   If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
   End If

   If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
      Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
      z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: _
         z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: _
         z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: _
         z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: _
         z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
      z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: _
         z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: _
         z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: _
         z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: _
         z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
      
      If GetProcAddress(LoadLibraryA("user32"), "IsWindowUnicode") Then
         If IsWindowUnicode(GetDesktopWindow()) Then
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW")  'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW")   'Store the SetWindowLong function address in the thunk data
         Else
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")  'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")   'Store the SetWindowLong function address in the thunk data
         End If
      End If
      
      'z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
      'z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
   End If

   z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

   If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
      On Error GoTo CatchDoubleSub                                           'Catch double subclassing
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

   Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
   zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
   VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory

End Function

'Terminate all subclassing
Private Sub sc_Terminate()

  Dim i As Long

   If Not (z_Funk Is Nothing) Then                 'Ensure that subclassing has been started

      With z_Funk
         For i = .Count To 1 Step -1               'Loop through the collection of window handles in reverse order
            z_ScMem = .Item(i)                     'Get the thunk address
            If IsBadCodePtr(z_ScMem) = 0 Then      'Ensure that the thunk hasn't already released its memory
               sc_UnSubclass zData(IDX_HWND)       'UnSubclass
            End If
         Next i                                    'Next member of the collection
      End With

      Set z_Funk = Nothing                         'Destroy the hWnd/thunk-address collection
   End If

End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)

   If z_Funk Is Nothing Then                                      'Ensure that subclassing has been started
      zError "sc_UnSubclass", "Window handle isn't subclassed"
   
   Else
      If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then               'Ensure that the thunk hasn't already released its memory
         zData(IDX_SHUTDOWN) = -1                                 'Set the shutdown indicator
         zDelMsg ALL_MESSAGES, IDX_BTABLE                         'Delete all before messages
         zDelMsg ALL_MESSAGES, IDX_ATABLE                         'Delete all after messages
      End If

      z_Funk.Remove "h" & lng_hWnd                                'Remove the specified window handle from the collection
   End If

End Sub

Public Property Get SearchColumn() As Long
Attribute SearchColumn.VB_Description = "Returns/sets a value that determines the default search column used in FindItem"

   SearchColumn = mSearchColumn

End Property

Public Property Let SearchColumn(ByVal vNewValue As Long)

   mSearchColumn = vNewValue
   PropertyChanged "SearchColumn"

End Property

Public Sub SelectedClearAll()

   '// MLH - new
  Dim lCount As Long

   If Not (mRowCount = C_NULL_RESULT) Then
      For lCount = LBound(mItems) To UBound(mItems)
         SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, False
      Next lCount
   End If

End Sub

Public Function SelectedCount() As Long

   '// Purpose: Return Count of Selected Items

  Dim lCount As Long

   If (mRowCount = C_NULL_RESULT) Then
      SelectedCount = 0

   Else
      For lCount = LBound(mItems) To UBound(mItems)
         If mItems(lCount).nFlags And lgFLSelected Then
            SelectedCount = SelectedCount + 1
         End If
      Next lCount
   End If

End Function

Private Sub SetColRect(ByVal Index As Long, ByRef r As Rect)

   '// Purpose: Set the drawing boundary for a Column
   '// MLH - was function changed to Sub

  Dim lCol         As Long
  Dim lCount       As Long
  Dim lScrollValue As Long
  Dim lScrollV     As Long
  Dim lX           As Long

   lScrollValue = SBValue(efsHorizontal)

   '// MLH - FreezeAtCol
   lScrollV = lScrollValue

   If Not mlngFreezeAtCol < 0 Then
      If lScrollV > mlngFreezeAtCol Then
         lScrollV = lScrollV - mlngFreezeAtCol - 1
      End If
   End If

   'If Index < lScrollValue Then  '// MLH - was before FreezeAtCol
   '// MLH - FreezeAtCol
   If mlngFreezeAtCol >= 0 Then
      If Index < lScrollValue And Index > mlngFreezeAtCol And lScrollValue > 0 Then
         r.Left = -1

      Else
         For lCol = 0 To Index - 1
            If lCol <= mlngFreezeAtCol Or lCol >= lScrollValue Then
               If mCols(mColPtr(lCol)).bVisible Then
                  lX = lX + mCols(mColPtr(lCol)).lWidth
                  lCount = lCount + 1
               End If

            End If
         Next lCol

         If IsColumnTruncated(Index) Then
            r.Left = mR.LeftText
            r.Right = r.Left + (mCols(mColPtr(Index)).lWidth - mR.LeftText)
         Else
            r.Left = lX
            r.Right = r.Left + mCols(mColPtr(Index)).lWidth
         End If

      End If

   Else  '// MLH - NOT FreezeAtCol
      If Index < lScrollValue And lScrollValue > 0 Then
         r.Left = -1

      Else
         For lCol = lScrollValue To Index - 1
            If mCols(mColPtr(lCol)).bVisible Then
               lX = lX + mCols(mColPtr(lCol)).lWidth
               lCount = lCount + 1
            End If
         Next lCol

         If IsColumnTruncated(Index) Then
            r.Left = mR.LeftText
            r.Right = r.Left + (mCols(mColPtr(Index)).lWidth - mR.LeftText)
         Else
            r.Left = lX
            r.Right = r.Left + mCols(mColPtr(Index)).lWidth
         End If

      End If
   End If

End Sub

Private Sub SetDefaultThemeColor(lngThemeType As Long)

   '// MLH - new sub
  Const C_ChangeBy As Integer = 25

   On Error GoTo Err_Proc

   Select Case lngThemeType

   Case 0 '// NormalColor
      mlngCustomColorFrom = RGB(203, 225, 252)
      mlngCustomColorTo = RGB(125, 165, 224)
      mBackColorBkg = vbApplicationWorkspace
      mBackColor = vbWindowBackground
      mForeColor = vbWindowText
      mBackColorSel = &HC56A31
      mForeColorSel = &HFFFFFF
      mFocusRectColor = &H96FFFE
      mGridColor = ColorBrightness(mlngCustomColorTo, C_ChangeBy)

      If Not (muFocusRowHighlightStyle = Solid) Then
         mForeColorSel = vbWindowText
      End If

   Case 1 '// Metallic
      mlngCustomColorFrom = RGB(244, 244, 251)
      mlngCustomColorTo = RGB(130, 130, 146)
      mBackColorBkg = vbApplicationWorkspace
      mBackColor = vbWindowBackground
      mForeColor = vbWindowText
      mBackColorSel = &HBFB4B2
      mForeColorSel = &H0
      mFocusRectColor = &H433D39
      mGridColor = ColorBrightness(mlngCustomColorTo, C_ChangeBy)

      If Not (muFocusRowHighlightStyle = Solid) Then
         mForeColorSel = vbWindowText
      End If

   Case 2 '// HomeStead
      mlngCustomColorFrom = RGB(247, 249, 225)
      mlngCustomColorTo = RGB(139, 161, 105)
      mBackColorBkg = vbApplicationWorkspace
      mBackColor = vbWindowBackground
      mForeColor = vbWindowText
      mBackColorSel = &H70A093
      mForeColorSel = &HFFFFFF
      mFocusRectColor = &H96FFFE
      mGridColor = ColorBrightness(mlngCustomColorTo, C_ChangeBy)

      If Not (muFocusRowHighlightStyle = Solid) Then
         mForeColorSel = vbWindowText
      End If

   Case 3 '// Storm
      mlngCustomColorFrom = RGB(248, 248, 242)
      mlngCustomColorTo = RGB(150, 159, 124)
      mBackColorBkg = &H6A7D6A
      mBackColor = vbWindowBackground
      mForeColor = vbWindowText
      mBackColorSel = &H778A77
      mForeColorSel = &HE1F9F7
      mFocusRectColor = &H96FFFE
      mGridColor = ColorBrightness(mlngCustomColorTo, C_ChangeBy)

      If Not (muFocusRowHighlightStyle = Solid) Then
         mForeColorSel = vbWindowText
      End If

   Case 4 '// Earth
      mlngCustomColorFrom = RGB(255, 239, 165)
      mlngCustomColorTo = RGB(160, 134, 73)
      mBackColorBkg = &HF4C66
      mBackColor = vbWindowBackground
      mForeColor = vbWindowText
      mBackColorSel = &H37748E
      mForeColorSel = &HE1EEF9
      mFocusRectColor = &H96FFFE
      mGridColor = ColorBrightness(mlngCustomColorTo, C_ChangeBy)

      If Not (muFocusRowHighlightStyle = Solid) Then
         mForeColorSel = vbWindowText
      End If

   End Select

Exit_Proc:
   On Error GoTo 0
   Exit Sub

Err_Proc:

   If IsInIDE Then
      MsgBox Err.Number & " - " & Err.Description, vbCritical, "ERROR - SetDefaultThemeColor"
   End If

   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub SetFlag(ByRef nFlags As Integer, ByVal nFlag As lgFlagsEnum, ByVal bValue As Boolean)

   If bValue Then
      nFlags = (nFlags Or nFlag)
   Else
      nFlags = (nFlags And Not (nFlag))
   End If

End Sub

Private Sub SetItemRect(ByVal vRow As Long, _
                        ByVal vCol As Long, _
                        ByVal lY As Long, _
                        ByRef r As Rect, _
                        ByVal ItemType As lgRectTypeEnum)

  Dim lHeight    As Long
  Dim lWidth     As Long
  Dim lLeft      As Long
  Dim lTop       As Long
  Dim nAlignment As lgAlignmentEnum

   vCol = mColPtr(vCol) '// MLH
   vRow = mRowPtr(vRow) '// MLH

   Select Case ItemType
   Case lgRTColumn
      nAlignment = mCols(vCol).nAlignment
      lHeight = mItems(vRow).lHeight
      lWidth = mCols(vCol).lWidth

   Case lgRTCheckBox
      nAlignment = mCols(vCol).nAlignment
      lHeight = mR.CheckBoxSize
      lWidth = mR.CheckBoxSize

   Case lgRTImage
      nAlignment = mCols(vCol).nImageAlignment
      lHeight = mR.ImageHeight
      lWidth = mR.ImageWidth
   End Select

   Select Case nAlignment
   Case lgAlignLeftTop
      lLeft = mCols(vCol).lX + 1
      lTop = lY + 2

   Case lgAlignLeftCenter
      lLeft = mCols(vCol).lX + 1
      lTop = (lY + (mItems(vRow).lHeight) / 2) - (lHeight / 2)

   Case lgAlignLeftBottom
      lLeft = mCols(vCol).lX + 1
      lTop = (lY + (mItems(vRow).lHeight)) - (lHeight + 2)

   Case lgAlignCenterTop
      lLeft = (mCols(vCol).lX + (mCols(vCol).lWidth) / 2) - (lWidth / 2)
      lTop = lY + 2

   Case lgAlignCenterCenter
      lLeft = (mCols(vCol).lX + (mCols(vCol).lWidth) / 2) - (lWidth / 2)
      lTop = (lY + (mItems(vRow).lHeight) / 2) - (lHeight / 2)

   Case lgAlignCenterBottom
      lLeft = (mCols(vCol).lX + (mCols(vCol).lWidth) / 2) - (lWidth / 2)
      lTop = (lY + (mItems(vRow).lHeight)) - (lHeight + 2)

   Case lgAlignRightTop
      lLeft = (mCols(vCol).lX + mCols(vCol).lWidth) - (lWidth + 1)
      lTop = lY + 2

   Case lgAlignRightCenter
      lLeft = (mCols(vCol).lX + mCols(vCol).lWidth) - (lWidth + 1)
      lTop = (lY + (mItems(vRow).lHeight) / 2) - (lHeight / 2)

   Case lgAlignRightBottom
      lLeft = (mCols(vCol).lX + mCols(vCol).lWidth) - (lWidth + 1)
      lTop = (lY + (mItems(vRow).lHeight)) - (lHeight + 2)

   End Select

   Call SetRect(r, lLeft, lTop, lLeft + lWidth, lTop + lHeight)

End Sub

Private Sub SetRedrawState(ByVal bState As Boolean)

   '// Purpose: Used to prevent Internal Redraws while preserving User Controlled Redraw state
   '// MLH - simplified

   If bState Then
      mbRedraw = mbUserRedraw
   Else
      mbRedraw = False
   End If

End Sub

Private Function SetRowCol(ByVal lRow As Long, _
                           Optional ByVal lCol As Long = C_NULL_RESULT, _
                           Optional ByVal bSetScroll As Boolean = False, _
                           Optional ByVal bMoveFocus As Boolean = True) As Boolean

   '// Purpose: To update current Row/Col and fire Events if necessary

  Dim r      As Rect
  Dim lCount As Long

   On Error Resume Next

   If RowLocked(lRow) Then Exit Function

   If Not (mCol = lCol And mRow = lRow) Then

      If Not (lCol = C_NULL_RESULT) Then
         If Not mCols(mColPtr(lCol)).bVisible Or mCols(mColPtr(lCol)).bLocked Then  '// MLH
            lCol = mCol
            mLCLocked = True
         End If
      End If

      If Not (lRow = C_NULL_RESULT) Then
         If Not mItems(mRowPtr(lRow)).bVisible Or (mItems(mRowPtr(lRow)).nFlags = lgFLlocked) Then  '// MLH - new
            lRow = mRow
            mLRLocked = True
         End If
      End If

      
      If bMoveFocus Then
         mCol = lCol
         mRow = lRow
         RaiseEvent RowColChanged
      End If

      '// Do we need to change Bars?
      If bSetScroll Then
         SetColRect lCol, r

         '// Scroll to make Column visible
         If r.Left < 0 Then

            For lCount = SBValue(efsHorizontal) To SBMin(efsHorizontal) Step -1

               If r.Left > 0 Then
                  Exit For
               End If

               SBValue(efsHorizontal) = SBValue(efsHorizontal) - 1
               SetColRect lCol, r
            Next lCount

         Else
            For lCount = SBValue(efsHorizontal) To SBMax(efsHorizontal)

               If r.Left + mCols(lCol).lWidth < UserControl.ScaleWidth Then
                  Exit For
               End If

               SBValue(efsHorizontal) = SBValue(efsHorizontal) + 1
               SetColRect lCol, r
            Next lCount

         End If

         If SBValue(efsHorizontal) = SBMin(efsHorizontal) Then
            SetScrollBars
         End If

         If lRow < SBValue(efsVertical) Then
            SBValue(efsVertical) = SBValue(efsVertical) - 1
         ElseIf lRow > SBValue(efsVertical) + (RowsVisible() - 1) Then
            SBValue(efsVertical) = SBValue(efsVertical) + 1
         End If

         RaiseEvent Scroll
      End If

      SetRowCol = True
   End If

End Function

Private Sub SetRowSize(ByVal vRow As Long)

  Dim r       As Rect
  Dim lCol    As Long
  Dim lHeight As Long
  Dim sText   As String
  Dim lMinRowHeight As Long

   '// MLH - modified

   If mbAllowWordWrap Then

      For lCol = LBound(mCols) To UBound(mCols)
         sText = mItems(mRowPtr(vRow)).Cell(lCol).sValue

         If (mItems(mRowPtr(vRow)).Cell(lCol).nFlags And lgFLWordWrap) Then
            SetRect r, 0, 2, mCols(lCol).lWidth - 5, 0
            DrawText UserControl.hdc, sText, Len(sText), r, DT_CALCRECT Or DT_WORDBREAK

         Else
            SetRect r, 0, 0, mCols(lCol).lWidth, 0
            DrawText UserControl.hdc, sText, Len(sText), r, DT_CALCRECT
         End If

         If r.Bottom > lHeight Then
            lHeight = r.Bottom
         End If

      Next lCol

      lMinRowHeight = mMinRowHeight - (mMinVerticalOffset * 2) - 2

      '// change Height to user selected scale
      'lHeight = ScaleY(lHeight, vbPixels, muScaleUnits)

      If lHeight < lMinRowHeight Then '// expand row height if necessary
         lHeight = lMinRowHeight

      ElseIf lHeight > (lMinRowHeight * mMaxLineCount) Then '// Limit Number of lines
         If mMaxLineCount > 0 Then
            lHeight = lMinRowHeight * mMaxLineCount
         End If

      End If

      '// change Height to Pixels and add vertical offset from grid lines
      mItems(mRowPtr(vRow)).lHeight = lHeight + (mMinVerticalOffset * 2)

   End If

End Sub

Private Sub SetScrollBars()

   '// Purpose: Sets the visibilty of scroll bars and sets max scroll values

  Dim lCol      As Long
  Dim lRow      As Long
  Dim lHeight   As Long
  Dim lWidth    As Long
  Dim lVSB      As Long
  Dim bHVisible As Boolean
  Dim bVVisible As Boolean

   On Error Resume Next
   
   If Not (mSBhWnd = 0) Then
      '// Calculate total width of columns
      For lCol = LBound(mCols) To UBound(mCols)
         If mCols(mColPtr(lCol)).bVisible Then
            lWidth = lWidth + mCols(mColPtr(lCol)).lWidth
         End If
      Next lCol

      If (lWidth > UserControl.ScaleWidth) Then
         SBMax(efsHorizontal) = UBound(mCols) - 1
         bHVisible = True
      Else
         SBMax(efsHorizontal) = UBound(mCols)
         bHVisible = (SBValue(efsHorizontal) > SBMin(efsHorizontal))
      End If

      '// Calculate total height of rows
      lHeight = GetColumnHeadingHeight() + UserControl.TextHeight(C_CHECKTEXT) '// MLH
      For lRow = LBound(mItems) To mRowCount
         If mItems(mRowPtr(lRow)).bVisible Then '// MLH - new
            lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
         End If
      Next lRow
      
      If lHeight > UserControl.ScaleHeight Then
         '// Adjust scrollbar to best-fit Rows to Grid
         lHeight = GetColumnHeadingHeight()

         For lRow = mRowCount To LBound(mItems) Step -1
            If mItems(mRowPtr(lRow)).bVisible Then '// MLH - new
               lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
            End If

            If lHeight > UserControl.ScaleHeight Then
               Exit For
            End If

            lVSB = lVSB + 1
         Next lRow

         If mbTotalsLineShow Then '// MLH - New
            SBMax(efsVertical) = mRowCount - lVSB + 2
         Else
            SBMax(efsVertical) = mRowCount - lVSB
         End If

         bVVisible = True
      Else
         SBMax(efsVertical) = mRowCount
         SBValue(efsVertical) = 0 '// MLH - new
      End If

      SBVisible(efsHorizontal) = bHVisible And (muSBOrienation = Scroll_Horizontal Or muSBOrienation = Scroll_Both) '// MLH
      SBVisible(efsVertical) = bVVisible And (muSBOrienation = Scroll_Vertical Or muSBOrienation = Scroll_Both)     '// MLH

   End If

End Sub

Private Function SetSelection(ByVal bState As Boolean, _
                              Optional ByRef lFromRow As Long = C_NULL_RESULT, _
                              Optional ByRef lToRow As Long = C_NULL_RESULT) As Boolean

  Dim lCount As Long
  Dim lStep  As Long
  Dim bSelectionChanged As Boolean

   If lFromRow = C_NULL_RESULT Then
      lFromRow = LBound(mItems)
   End If

   If lToRow = C_NULL_RESULT Then
      lToRow = UBound(mItems)
   End If

   If lFromRow >= lToRow Then
      lStep = -1
   Else
      lStep = 1
   End If

   For lCount = lFromRow To lToRow Step lStep

      If (mItems(mRowPtr(lCount)).nFlags And lgFLSelected) <> bState Then
         SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, bState
         bSelectionChanged = True
      End If

   Next lCount

   SetSelection = bSelectionChanged

End Function

Private Sub SetThemeColor()

   '// MLH - new sub

   If muThemeColor = Autodetect Then
      GetGradientColor UserControl.hWnd
   Else
      SetDefaultThemeColor muThemeColor
   End If

End Sub

Private Sub ShowCompleteCell(ByVal lRow As Long, ByVal lCol As Long)

   '// MLH - new sub
   '// Purpose: wait 1.5 seconds before displaying complete cell's text

   If mbAutoToolTips Then                                                        '// If Allow Auto Tool Tips?
      If Not mbWorking Then                                                      '// If not Already Timing?
         If IsValidRowCol(lRow, lCol) Then                                       '// If Is Row/Col Valid?
            If Not mbEditPending Then                                            '// If Not Editing
               If Not (mCols(mColPtr(lCol)).nType = lgBoolean) Then              '// If Cell is not boolean
                  If LenB(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).sValue) Then '// If Cell contains text?

                     mbWorking = True
                     mlTickCount = GetTickCount()

                     Do
                        If mlTickCount + 1500 <= GetTickCount() Then Exit Do
                        DoEvents
                     Loop Until mbCancelShow

                     mbWorking = False

                     If Not mbCancelShow Then
                        '// send mouse position in case it has changed during timing
                        Call ShowCompleteCellx(mMouseRow, mMouseCol)
                     End If

                     mbCancelShow = False
                  End If

               End If
            End If
         End If
      End If
   End If

End Sub

Private Sub ShowCompleteCellx(ByVal lRow As Long, ByVal lCol As Long)

  Dim r            As Rect
  Dim CR           As Rect
  Dim RectM        As Rect
  Dim cHeight      As Long
  Dim cWidth       As Long
  Dim tWidth       As Long
  Dim sText        As String
  Dim lngW         As Long
  Dim sCharW       As Long
  Dim lMinRowH     As Long

   '// MLH - new sub
   On Local Error Resume Next

   If lRow = C_NULL_RESULT Then Exit Sub

   picTooltip.Visible = False

   picTooltip.FontBold = ucFontBold
   picTooltip.FontItalic = ucFontItalic
   picTooltip.FontName = ucFontName

   If Not mbEditPending Then
      If LenB(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).sValue) Then
         If IsValidRowCol(lRow, lCol) Then
            If Not (mCols(mColPtr(lCol)).nType = lgBoolean) Then

               '// needed to return correct text height/width
               UserControl.FontBold = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLFontBold
               UserControl.FontItalic = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLFontItalic
               UserControl.FontName = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).sFontName

               picTooltip.FontBold = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLFontBold
               picTooltip.FontItalic = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLFontItalic
               picTooltip.FontName = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).sFontName
               picTooltip.FontSize = UserControl.FontSize

               sText = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).sValue

               If LenB(mCols(mColPtr(lCol)).sFormat) Then
                  sText = Format$(sText, mCols(mColPtr(lCol)).sFormat)
               End If

               sCharW = CInt(TextWidth(C_CHECKTEXT) / 14) '// 3/4 avg. char. width

               lMinRowH = (mMinVerticalOffset * 2) + 2
               cHeight = mItems(mRowPtr(lRow)).lHeight
               tWidth = TextWidth(sText)

               If IsColumnTruncated(mColPtr(lCol)) Then
                  cWidth = mCols(mColPtr(lCol)).lWidth - mR.LeftText
               Else
                  cWidth = mCols(mColPtr(lCol)).lWidth
               End If

               If tWidth < cWidth Then tWidth = cWidth

               If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).lImage > 0 Then
                  cWidth = cWidth - sCharW - mR.ImageWidth
               Else
                  cWidth = cWidth - sCharW
               End If

               '// test to see if we need to show
               If mbAllowWordWrap And CBool(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLWordWrap) Then

                  If mExpandRowImage > 0 And cHeight > mR.TextHeight + lMinRowH Then
                     cWidth = cWidth - sCharW - mR.ImageWidth
                  End If

                  cWidth = cWidth - 2

                  SetRect r, 0, 0, mCols(mColPtr(lCol)).lWidth, 0
                  Call DrawText(UserControl.hdc, sText, Len(sText), r, DT_CALCRECT Or DT_WORDBREAK)

                  '// test to see if we need to show
                  r.Right = r.Right + sCharW
                  If r.Bottom + lMinRowH <= mItems(mRowPtr(lRow)).lHeight And r.Right <= cWidth Then GoTo ExitShowToolTip

               Else
                  SetRect r, 0, 0, mCols(mColPtr(lCol)).lWidth, 0
                  Call DrawText(UserControl.hdc, sText, Len(sText), r, DT_CALCRECT Or DT_SINGLELINE)

                  '// test to see if we need to show
                  r.Right = r.Right + sCharW
                  If r.Right <= cWidth Then GoTo ExitShowToolTip
               End If

               '// begin show of 'Full View'
               sCharW = sCharW * 2
               SetRect r, 0, 0, tWidth, 0
               Call DrawText(UserControl.hdc, sText, Len(sText), r, DT_CALCRECT Or DT_WORDBREAK)

               SetColRect lCol, CR
               CR.Top = RowTop(lRow)

               GetWindowRect hWnd, RectM
               lngW = Screen.Width / Screen.TwipsPerPixelX - (RectM.Left + CR.Left)
               If lngW < tWidth Then '// does it go beyond the edge of the screen?
                  tWidth = lngW - sCharW
                  SetRect r, 0, 0, tWidth, 0
                  Call DrawText(UserControl.hdc, sText, Len(sText), r, DT_CALCRECT Or DT_WORDBREAK)
                  lngW = C_NULL_RESULT
               End If

               If lCol = 0 Then
                  r.Left = r.Left + 2
               Else
                  r.Left = r.Left - 1
               End If

               r.Top = r.Top - 1

               r.Right = r.Right + sCharW + C_TEXT_SPACE

               If r.Bottom < cHeight Then
                  r.Bottom = cHeight
               Else
                  r.Bottom = r.Bottom + lMinRowH
               End If

               RectM = r

               '// Draw rect
               picTooltip.Cls
               GetWindowRect hWnd, r
               picTooltip.Move (r.Left + CR.Left) * Screen.TwipsPerPixelX, (r.Top + CR.Top) * Screen.TwipsPerPixelY, _
                  RectM.Right * Screen.TwipsPerPixelX, RectM.Bottom * Screen.TwipsPerPixelY

               '// Draw Text
               RectM.Left = RectM.Left + C_TEXT_SPACE
               RectM.Right = RectM.Right - (C_TEXT_SPACE * 2)

               If InStrB(1, sText, vbCr) Or lngW = C_NULL_RESULT Then
                  lngW = DT_WORDBREAK
               Else
                  lngW = mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nAlignment Or DT_WORDBREAK Or DT_SINGLELINE
               End If

               Call DrawText(picTooltip.hdc, sText, Len(sText), RectM, lngW)

               picTooltip.Visible = True
               picTooltip.ZOrder

            End If
         End If
      End If
   End If

   UserControl.FontBold = ucFontBold
   UserControl.FontItalic = ucFontItalic
   UserControl.FontName = ucFontName

   Exit Sub

ExitShowToolTip:
   UserControl.FontBold = ucFontBold
   UserControl.FontItalic = ucFontItalic
   UserControl.FontName = ucFontName

   picTooltip.Visible = False

End Sub

Public Sub Sort(Optional ByVal vCol1 As Long = C_NULL_RESULT, _
                Optional ByVal vCol1SortType As lgSortTypeEnum = C_NULL_RESULT, _
                Optional ByVal vCol2 As Long = C_NULL_RESULT, _
                Optional ByVal vCol2SortType As lgSortTypeEnum = C_NULL_RESULT)

   '// Purpose: Sort Grid based on current Sort Columns.
  Dim lCount    As Long
  Dim lRowIndex As Long

   If UpdateCell() Then

      '// Set new Columns if specified
      If Not (vCol1 = C_NULL_RESULT) Then
         mSortColumn = vCol1
      End If

      If Not (vCol2 = C_NULL_RESULT) Then
         mSortSubColumn = vCol2
      End If

      '// Validate Sort Columns
      If (mSortColumn = C_NULL_RESULT) And Not (mSortSubColumn = C_NULL_RESULT) Then
         mSortColumn = mSortSubColumn
         mSortSubColumn = C_NULL_RESULT

      ElseIf mSortColumn = mSortSubColumn Then
         mSortSubColumn = C_NULL_RESULT
      End If

      '// MLH - Fix column number in case column order was changed
      If (vCol1 = C_NULL_RESULT) Then
         mSortColumn = mColPtr(mSortColumn)
      End If
      If (vCol2 = C_NULL_RESULT) Then
         If Not mSortSubColumn = C_NULL_RESULT Then
            mSortSubColumn = mColPtr(mSortSubColumn)
         End If
      End If
      
      '// Set Sort Order if specified - otherwise inverse last Sort Order
      With mCols(mSortColumn)
         If vCol1SortType = C_NULL_RESULT Then

            Select Case .nSortOrder '// MLH
            Case lgSTNormal
               .nSortOrder = lgSTDescending

            Case lgSTAscending
               .nSortOrder = lgSTNormal

            Case lgSTDescending
               .nSortOrder = lgSTAscending
            End Select

         Else
            .nSortOrder = vCol1SortType
         End If
      End With

      If Not (mSortSubColumn = C_NULL_RESULT) Then
         With mCols(mSortSubColumn)
            If vCol2SortType = C_NULL_RESULT Then
               .nSortOrder = mCols(mSortColumn).nSortOrder '// MLH
'               Select Case .nSortOrder '// MLH
'               Case lgSTNormal
'                  .nSortOrder = lgSTDescending
'               Case lgSTAscending
'                  .nSortOrder = lgSTNormal
'               Case lgSTDescending
'                  .nSortOrder = lgSTAscending
'               End Select
            Else
               .nSortOrder = vCol2SortType
            End If
         End With
      End If

      '// Note previously selected Row
      If Not (mRow = C_NULL_RESULT) Then
         lRowIndex = mRowPtr(mRow)
      End If

      If mCols(mSortColumn).nSortOrder = lgSTNormal Then '// MLH
         For lCount = 0 To mRowCount
            mRowPtr(lCount) = lCount
         Next lCount

         mSortColumn = C_NULL_RESULT
         mSortSubColumn = C_NULL_RESULT

      Else
         Call SortArray(LBound(mItems), mRowCount, mSortColumn, mCols(mSortColumn).nSortOrder)
         Call SortSubList
      End If

      'If Not mSortColumn = C_NULL_RESULT Then
         For lCount = LBound(mRowPtr) To mRowCount
            If mRowPtr(lCount) = lRowIndex Then
               mRow = lCount
               RowColSet lCount ', mColPtr(mSortColumn) '// MLH - keep selected row visible
               Exit For
            End If
         Next lCount
      'End If
      
      'DrawGrid mbRedraw '// MLH

      RaiseEvent SortComplete
   End If

End Sub

Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

   Select Case mCols(lSortColumn).nType
   Case lgBoolean
      SortArrayBool lFirst, lLast, lSortColumn, nSortType

   Case lgDate
      SortArrayDate lFirst, lLast, lSortColumn, nSortType

   Case lgNumeric
      SortArrayNumeric lFirst, lLast, lSortColumn, nSortType

   Case lgCustom
      SortArrayCustom lFirst, lLast, lSortColumn, nSortType

   Case Else
      SortArrayString lFirst, lLast, lSortColumn, nSortType
   End Select

End Sub

Private Sub SortArrayBool(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

  Dim lBoundary As Long
  Dim lIndex    As Long
  Dim bSwap     As Boolean

   If Not (lLast <= lFirst) Then

      SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)

      lBoundary = lFirst

      For lIndex = lFirst + 1 To lLast
         bSwap = False

         If nSortType = lgSTAscending Then
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) > _
               GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
         Else
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) < _
               GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
         End If

         If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
         End If

      Next lIndex

      SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
      SortArrayBool lFirst, lBoundary - 1, lSortColumn, nSortType
      SortArrayBool lBoundary + 1, lLast, lSortColumn, nSortType

   End If

End Sub

Private Sub SortArrayCustom(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

  Dim lBoundary As Long
  Dim lIndex    As Long
  Dim bSwap     As Boolean

   If Not (lLast <= lFirst) Then

      SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)

      lBoundary = lFirst

      For lIndex = lFirst + 1 To lLast
         bSwap = False

         If nSortType = lgSTAscending Then
            RaiseEvent CustomSort(True, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, _
               mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
         Else
            RaiseEvent CustomSort(False, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, _
               mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
         End If

         If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
         End If

      Next lIndex

      SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
      SortArrayCustom lFirst, lBoundary - 1, lSortColumn, nSortType
      SortArrayCustom lBoundary + 1, lLast, lSortColumn, nSortType

   End If

End Sub

Private Sub SortArrayDate(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

  Dim lBoundary  As Long
  Dim lIndex     As Long
  Dim bIsDate(1) As Boolean
  Dim bSwap      As Boolean

   If Not (lLast <= lFirst) Then

      SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)

      lBoundary = lFirst

      For lIndex = lFirst + 1 To lLast
         bIsDate(0) = IsDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue)
         bIsDate(1) = IsDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)

         If nSortType = lgSTAscending Then
            If Not bIsDate(0) Then
               bSwap = False

            ElseIf Not bIsDate(1) Then
               bSwap = True

            Else
               bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > _
                  CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If

         Else
            If Not bIsDate(0) Then
               bSwap = True

            ElseIf Not bIsDate(1) Then
               bSwap = False
            
            Else
               bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < _
                  CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If

         End If

         If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
         End If

      Next lIndex

      SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
      SortArrayDate lFirst, lBoundary - 1, lSortColumn, nSortType
      SortArrayDate lBoundary + 1, lLast, lSortColumn, nSortType

   End If

End Sub

Private Sub SortArrayNumeric(ByVal lFirst As Long, _
                             ByVal lLast As Long, _
                             ByVal lSortColumn As Long, _
                             ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

  Dim lBoundary As Long
  Dim lIndex    As Long
  Dim bSwap     As Boolean

   If Not (lLast <= lFirst) Then

      SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)

      lBoundary = lFirst

      For lIndex = lFirst + 1 To lLast
         bSwap = False

         If nSortType = lgSTAscending Then
            bSwap = rVal(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > _
               rVal(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
         Else
            bSwap = rVal(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < _
               rVal(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
         End If

         If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
         End If

      Next lIndex

      SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
      SortArrayNumeric lFirst, lBoundary - 1, lSortColumn, nSortType
      SortArrayNumeric lBoundary + 1, lLast, lSortColumn, nSortType

   End If

End Sub

Private Sub SortArrayString(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows

  Dim lBoundary As Long
  Dim lIndex    As Long
  Dim bSwap     As Boolean
  Dim vResult   As Variant

   On Error Resume Next
   
   If Not (lLast <= lFirst) Then

      SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)

      lBoundary = lFirst

      For lIndex = lFirst + 1 To lLast

         '// MLH - modified to ignore string case
         'bSwap = False
         vResult = StrComp(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, vbTextCompare)
         
         If nSortType = lgSTAscending Then
            bSwap = (vResult = 1) 'mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue > mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
         Else
            bSwap = (vResult = -1) 'mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue < mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
         End If

         If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
         End If

      Next lIndex

      SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
      SortArrayString lFirst, lBoundary - 1, lSortColumn, nSortType
      SortArrayString lBoundary + 1, lLast, lSortColumn, nSortType
   End If
   
   On Error GoTo 0

End Sub

Private Sub SortSubList()

   '// Purpose: Used to sort by a secondary Column after a Sort

  Dim lCount     As Long
  Dim lStartSort As Long
  Dim bDifferent As Boolean
  Dim sMajorSort As String

   If mSortSubColumn > C_NULL_RESULT Then
      '// Re-Sort the Items by a secondary column, preserving the sort sequence of the primary sort
      lStartSort = LBound(mItems)

      For lCount = LBound(mItems) To mRowCount
         bDifferent = Not (mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue = sMajorSort)

         If bDifferent Or lCount = mRowCount Then
            If lCount > 1 Then
               If lCount - lStartSort > 1 Then
                  If lCount = mRowCount And Not bDifferent Then
                     SortArray lStartSort, lCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                  Else
                     SortArray lStartSort, lCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                  End If

               End If
               lStartSort = lCount
            End If

            sMajorSort = mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue
         End If

      Next lCount
   End If

End Sub

Private Function SplitToLines(ByVal sText As String, _
                              ByVal lLength As Long, _
                              Optional ByVal iLines As Integer = C_NULL_RESULT) As String

  Dim strChar    As String
  Dim strTemp    As String
  Dim lngPos     As Long
  Dim lngI       As Long
  Dim lLineCount As Long
  Dim lChrCount  As Long

   '// MLH - new sub
   '// Purpose: Split a single line of text into multi-line seperated by vbNewLine

   sText = Replace(Trim$(sText), vbNewLine, " ")

   For lngI = 1 To Len(sText)

      strChar = Mid$(sText, lngI, 1) '// get single character
      strTemp = strTemp & strChar    '// add character to temp string
      lChrCount = Len(strTemp)       '// get temp string's length

      '// Check if the char is a Delimiter
      Select Case Asc(strChar)
      Case 65 To 90, 95, 97 To 122
         '// alphnumeric and "_" are not delimiters
      Case Else
         lngPos = lChrCount
      End Select

      '// split if column width is exceded and we have a break point
      If TextWidth(strTemp) >= lLength And lngPos Then

         If LenB(SplitToLines) Then
            '// not first join
            SplitToLines = SplitToLines & vbNewLine & Trim$(Mid$(strTemp, 1, lngPos))
         Else
            '// first join
            SplitToLines = Trim$(Mid$(strTemp, 1, lngPos))
         End If

         If lChrCount > lngPos Then
            '// save leftover text
            strTemp = Trim$(Mid$(strTemp, lngPos + 1))
         Else
            strTemp = vbNullString
         End If

         lChrCount = Len(strTemp)
         lngPos = 0

         '// limit split to ? lines
         If Not (iLines = C_NULL_RESULT) Then
            lLineCount = lLineCount + 1
         End If

      End If

      If iLines = lLineCount Then
         Exit For
      End If

   Next lngI

   '// catch any remaining text
   If iLines = C_NULL_RESULT Then '// no line limit

      If LenB(SplitToLines) Then
         If LenB(strTemp) Then
            SplitToLines = SplitToLines & vbNewLine & strTemp
         End If

      Else
         SplitToLines = strTemp
      End If

   ElseIf lLineCount < iLines Then '// less then or equal to line limit
      If LenB(SplitToLines) = 0 Then
         SplitToLines = Trim$(strTemp)
      Else
         SplitToLines = SplitToLines & vbNewLine & Trim$(strTemp)
      End If

   Else '// greater than line limit
      If LenB(strTemp) Then strTemp = " " & Left$(strTemp, Len(strTemp) - 1)
      SplitToLines = SplitToLines & strTemp
      SplitToLines = Trim$(Left$(SplitToLines, Len(SplitToLines) - 3)) & "..."
   End If

End Function

Private Sub SwapLng(ByRef Value1 As Long, ByRef Value2 As Long)

  Dim lTemp As Long

   lTemp = Value1
   Value1 = Value2
   Value2 = lTemp

End Sub

Public Property Get ThemeColor() As lgThemeConst
Attribute ThemeColor.VB_Description = "Returns/sets a value that determines the theme color (Blue, Silver, Olive, etc.)"

   ThemeColor = muThemeColor '// MLH

End Property

Public Property Let ThemeColor(ByVal vData As lgThemeConst)

   '// MLH - New

   If Not (muThemeColor = vData) Then
      muThemeColor = vData
      PropertyChanged "ThemeColor"
      Call SetThemeColor
      DrawGrid True 'mbRedraw
   End If

End Property

Public Property Get ThemeCustomColorFrom() As OLE_COLOR
Attribute ThemeCustomColorFrom.VB_Description = "Used when Custom Theme Color and XP Office to draw Column Headers"

   '// MLH
   ThemeCustomColorFrom = mlngCustomColorFrom

End Property

Public Property Let ThemeCustomColorFrom(ByVal new_ColorFrom As OLE_COLOR)

   '// MLH - New
   new_ColorFrom = TranslateColor(new_ColorFrom)
   PropertyChanged "ThemeColorFrom"

   If muThemeColor = CustomTheme Then
      mlngCustomColorFrom = new_ColorFrom
   End If

   Call Refresh

End Property

Public Property Get ThemeCustomColorTo() As OLE_COLOR
Attribute ThemeCustomColorTo.VB_Description = "Used when Custom Theme Color and XP Office to draw Column Headers"

   '// MLH
   ThemeCustomColorTo = mlngCustomColorTo

End Property

Public Property Let ThemeCustomColorTo(ByVal new_ColorTo As OLE_COLOR)

   '// MLH - New
   new_ColorTo = TranslateColor(new_ColorTo)
   PropertyChanged "ThemeColorTo"

   If muThemeColor = CustomTheme Then
      mlngCustomColorTo = new_ColorTo
   End If

   Call Refresh

End Property

Public Property Get ThemeStyle() As lgThemeStyleEnum
Attribute ThemeStyle.VB_Description = "Returns/sets a value that determines Header type (3D, Flat, XP, XP Office)"

   ThemeStyle = muThemeStyle

End Property

Public Property Let ThemeStyle(ByVal vNewValue As lgThemeStyleEnum)

   muThemeStyle = vNewValue
   PropertyChanged "ThemeStyle"
   DrawGrid True 'mbRedraw

End Property

Private Function ToggleEdit(Optional ByVal bAllowMove As Boolean = False) As Boolean

   '// Purpose: Used to start a new Edit or commit a pending one

   If IsEditable() Then
      ToggleEdit = True

      If mbEditPending Then
         Call UpdateCell(bAllowMove)

      ElseIf Not (mRow = C_NULL_RESULT) And Not (mCol = C_NULL_RESULT) Then
         EditCell mRow, mCol
      End If

   End If

End Function

Public Property Get TotalsCol(ByVal vCol As Long) As Double

   '// MLH - New
   On Error Resume Next
   TotalsCol = mudtTotalsVal(vCol)

End Property

Public Property Get TotalsLineCaption(ByVal Index As Long) As String

   '// MLH - New
   TotalsLineShow = mbTotalsLineShow

End Property

Public Property Let TotalsLineCaption(ByVal Index As Long, ByVal vNewValue As String)

   On Error Resume Next
   mudtTotals(Index).sCaption = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get TotalsLineColAvg(ByVal Index As Long) As Boolean

   '// MLH - New
   TotalsLineColAvg = mudtTotals(Index).bAvg

End Property

Public Property Let TotalsLineColAvg(ByVal Index As Long, ByVal vNewValue As Boolean)

   On Error Resume Next
   mudtTotals(Index).bAvg = vNewValue
   DrawGrid mbRedraw

End Property

Public Property Get TotalsLineShow() As Boolean
Attribute TotalsLineShow.VB_Description = "Returns/sets a value that determines whether the Total line shows for numeric column types"

   '// MLH - New
   TotalsLineShow = mbTotalsLineShow

End Property

Public Property Let TotalsLineShow(ByVal vNewValue As Boolean)

   On Error Resume Next
   mbTotalsLineShow = vNewValue
   PropertyChanged "TotalsLineShow"
   DrawGrid mbRedraw

End Property

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

  Dim tme As TRACKMOUSEEVENT_STRUCT

   'If bTrack Then '// MLH - bTrack was never set = True
   On Error Resume Next

   With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
   End With

   'If bTrackUser32 Then '// MLH
   If Not (mbWinNT Or mbWinXP) Then
      Call TrackMouseEvent(tme)
   Else
      Call TrackMouseEventComCtl(tme)
   End If

   On Error GoTo 0
   'End If

End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long

   If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If

End Function

Private Sub txtEdit_DblClick()
   Call UserControl_DblClick '// MLH
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

   Select Case mCols(mColPtr(mEditCol)).sInputFilter
   Case vbNullString
      '// No Filter
      
   Case "<", ">"
      '// lowercase/UPPERCASE
      
   Case Else '// Custom Filter
      Select Case KeyAscii
      Case vbKeyBack, vbKeyDelete
         '// Do not restrict!
      Case Else
         If InStr(mCols(mColPtr(mEditCol)).sInputFilter, Chr$(KeyAscii)) = 0 Then
            KeyAscii = 0
         End If
      End Select
   End Select

   '// Allow outside filtering
   RaiseEvent EditKeyPress(mColPtr(mEditCol), KeyAscii)

End Sub

Private Function UpdateCell(Optional ByVal bAllowMove As Boolean = False) As Boolean

   '// Purpose: Used to commit an Edit. Note the RequestUpate event. This event allows
   '// the Upate to be cancelled by setting the Cancel flag.

  Dim bCancel        As Boolean
  Dim bRequestUpdate As Boolean
  Dim sNewValue      As String

   If mbEditPending Then
      If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
         If LenB(Trim$(txtEdit.Text)) = 0 Then txtEdit.Text = vbNullString '// MLH - Don't save blank spaces
         sNewValue = txtEdit.Text
         bRequestUpdate = Not (mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue = sNewValue)

      Else
         bRequestUpdate = True
      End If

      If bRequestUpdate Then
         RaiseEvent AfterEdit(mEditRow, mColPtr(mEditCol), sNewValue, bCancel)
      End If

      If Not bCancel Then
         '// Turn off redraw if necessary
         SetRedrawState False

         If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
            txtEdit.Visible = False
         Else
            On Local Error Resume Next

            With mCols(mColPtr(mEditCol)).EditCtrl

               If Not (mEditParent = 0) Then
                  SetParent .hWnd, mEditParent
               End If

               sc_UnSubclass .hWnd

               .Visible = False
            End With

            On Local Error GoTo 0
         End If

         mbEditPending = False

         If bRequestUpdate Then
            If LenB(Trim$(sNewValue)) = 0 Then sNewValue = vbNullString '// MLH - Don't save blank spaces
            '// MLH - new
             mudtTotalsVal(mColPtr(mEditCol)) = mudtTotalsVal(mColPtr(mEditCol)) - rVal(mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue)
             mudtTotalsVal(mColPtr(mEditCol)) = mudtTotalsVal(mColPtr(mEditCol)) + rVal(sNewValue)

            mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue = sNewValue
            Call SetRowSize(mEditRow)

            SetFlag mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags, lgFLChanged, True

            Call DisplayChange
         End If

         '// MLH - New

         If bAllowMove Then

            Select Case muEditMove
            Case lgMoveRight
               mCol = NavigateRight(False)

            Case lgMoveDown
               mRow = NavigateDown
            End Select

            Call SetSelection(False)
            RowColSet mRow, mCol
         End If

         '// Restore redraw state to user selected
         Call SetRedrawState(True)
         Call DrawGrid(mbRedraw)

      End If

   End If

   UpdateCell = Not bCancel

End Function

Private Sub UserControl_Click()

   If Not mMouseRow = C_NULL_RESULT Then
      If Not mItems(mRowPtr(mMouseRow)).nFlags And lgFLlocked Then '// MLH

         If (muEditTrigger And lgMouseClick) Then
            Call ToggleEdit
         End If

      End If
   End If

   If mMouseDownRow > C_NULL_RESULT Then
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_DblClick()

   If Not mMouseRow = C_NULL_RESULT Then
      If Not mItems(mRowPtr(mMouseRow)).nFlags And lgFLlocked Then '// MLH
         If (muEditTrigger And lgMouseDblClick) Then
            miKeyCode = vbKeyF2
            Call ToggleEdit
         End If
      End If
   End If
   
   If Not mMouseDownRow = C_NULL_RESULT Then
      RaiseEvent DblClick
   
   ElseIf UserControl.MousePointer = vbSizeWE Then '// MLH Autoresize Column
      If mbAllowColumnResizing Then
         Call ColWidthAutoSize(mColPtr(mMouseCol))
      End If
   End If

End Sub

Public Function GetInfo(ByVal lInfo As Long) As String
  '// MLH - New
  Dim Buffer   As String
  Dim Ret      As String
   
   Buffer = String$(256, 0)
   Ret = GetLocaleInfo(&H400, lInfo, Buffer, Len(Buffer))
   If Ret > 0 Then
      GetInfo = Left$(Buffer, Ret - 1)
   Else
      GetInfo = ""
   End If

End Function

Private Sub UserControl_Initialize()

  Dim OS As OSVersionInfo

   mClipRgn = CreateRectRgn(0, 0, 0, 0)

   'mbLockFocusDraw = IsWindowUnicode(UserControl.hWnd) '// MLH
   mbLockFocusDraw = (GetInfo(&H1001) = "English")      '// MLH

   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)

   mbWinNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

   If (OS.dwMajorVersion > 5) Then '// Vista
      mbWinXP = True
   ElseIf (OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1) Then '// XP
      mbWinXP = True
   End If

   Set txtEdit = UserControl.Controls.Add("VB.TextBox", "txtEdit")

   With txtEdit
      .BorderStyle = 0
      .Visible = False

      If mbWinNT Then
         mTextBoxStyle = GetWindowLongW(.hWnd, GWL_STYLE)
      Else
         mTextBoxStyle = GetWindowLongA(.hWnd, GWL_STYLE)
      End If
   End With

   Set picTooltip = UserControl.Controls.Add("VB.PictureBox", "picToolTip")

   With picTooltip
      .AutoRedraw = True
      .BorderStyle = 1
      .Appearance = 0
      .BackColor = vbInfoBackground
      .ForeColor = vbBlack
      .Enabled = False
   End With

   SetParent picTooltip.hWnd, GetDesktopWindow
   SetWindowLongA picTooltip.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW

   ReDim mCols(0) As udtColumn
   ReDim mColPtr(0) As Long

   '''Clear '// MLH - called twice during load if mFormatString="" in UserControl_ReadProperties

End Sub

Private Sub UserControl_InitProperties()

   Set mFont = Ambient.Font

   '---------------------------------------------------------------------------------
   '// Appearance Properties
   mbApplySelectionToImages = False
   mBackColor = vbWindowBackground
   mBackColorBkg = vbApplicationWorkspace
   mBackColorEdit = vbInfoBackground
   'mBackColorFixed = vbButtonFace'// MLH - Not Needed
   mBackColorSel = vbHighlight
   mForeColor = vbWindowText
   mForeColorEdit = vbInfoText
   'mForeColorFixed = vbButtonText''''// MLH - Not Used
   mForeColorHdr = vbWindowText
   mForeColorSel = vbHighlightText
   '''   mForeColorTotals = DEF_FORECOLORTOTALS '// MLH - not used

   mFocusRectColor = vbYellow
   mGridColor = DEF_GRIDCOLOR
   mProgressBarColor = DEF_PROGRESSBARCOLOR

   mbAlphaBlendSelection = False
   mbDisplayEllipsis = True
   muFocusRectMode = DEF_FOCUSRECTMODE
   muFocusRectStyle = DEF_FOCUSRECTSTYLE
   muGridLines = lgGridLinesEnum.lgGrid_Both
   mblnColumnHeaderSmall = False
   mGridLineWidth = DEF_GRIDLINEWIDTH
   muThemeStyle = DEF_THEMESTYLE
   mbColumnHeaders = True '// MLH
   mbCenterRowImage = True '// MLH
   muSBOrienation = ScrollBarOrienationEnum.Scroll_Both '// MLH
   mMinVerticalOffset = DEF_MinVerticalOffset '// MLH
   mColumnHeaderLines = 1 '// MLH
   msCaption = vbNullString '// MLH
   muCaptionAlignment = lgAlignCenter '// MLH

   '---------------------------------------------------------------------------------
   '// Behaviour Properties
   mbAllowRowResizing = False '// MLH - New
   mbAllowColumnResizing = True '// MLH
   mbAllowWordWrap = False '// MLH
   mbAllowColumnSwap = False
   mbAllowColumnDrag = False '// MLH
   mbAllowColumnSort = False
   mbAllowEdit = False
   mbAllowDelete = False '// MLH
   mbAllowInsert = False '// MLH
   muBorderStyle = DEF_BORDERSTYLE
   mbCheckboxes = False
   muEditTrigger = lgEditTriggerEnum.lgAnyF2DblCk
   mbFullRowSelect = True
   muFocusRowHighlightStyle = lgFocusRowHighlightStyle.Solid '// MLH - new
   mbHideSelection = False
   mbAllowColumnHover = True
   muMultiSelect = lgMultiSelectEnum.lgSingleSelect '// MLH - Added New Options
   mbRedraw = False '// MLH
   mbUserRedraw = mbRedraw '// MLH
   mbScrollTrack = True
   mbAutoToolTips = True
   mlngFreezeAtCol = -1

   '---------------------------------------------------------------------------------
   '// Miscellaneous Properties
   muScaleUnits = vbTwips
   mSearchColumn = DEF_SEARCHCOLUMN

   mCacheIncrement = DEF_CACHEINCREMENT
   mbEnabled = True
   '''mFormatString = DEF_FORMATSTRING'// MLH - function removed
   '''mLocked = DEF_LOCKED '// MLH - Not Used
   mMaxLineCount = DEF_MaxLineCount
   mMinRowHeightUser = DEF_MinRowHeight
   mMinRowHeight = ScaleY(mMinRowHeightUser, vbTwips, vbPixels) 'mMinRowHeightUser

   mBackColorEvenRows = &HEDEBE0                   '// MLH
   mbBackColorEvenRowsE = True                     '// MLH
   mlngCustomColorFrom = TranslateColor(&H404080)  '// MLH
   mlngCustomColorTo = TranslateColor(&HC0E0FF)    '// MLH
   mlngCustomColorFrom = mlngCustomColorFrom       '// MLH
   mlngCustomColorTo = mlngCustomColorTo           '// MLH
   muThemeColor = Autodetect                       '// MLH

   '---------------------------------------------------------------------------------
   '// Apply Settings
   Call SetThemeColor '// MLH

   With UserControl
      .BackColor = mBackColorBkg
      .BorderStyle = muBorderStyle
   End With

   Call CreateRenderData

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim lNewCol         As Long
  Dim lNewRow         As Long
  Dim iKeyCode        As Integer
  Dim bClearSelection As Boolean
  Dim bRedraw         As Boolean
  Dim bCancel         As Boolean
  Dim strTemp         As String
  Dim bState          As Boolean

   lNewCol = mCol
   lNewRow = mRow
   iKeyCode = KeyCode '// MLH
   picTooltip.Visible = False
   mLRLocked = False
   mLCLocked = False

   '// Turn off redraw in necessary
   SetRedrawState False

   '// Used to determine if selected Items need to be cleared
   bClearSelection = True

   Select Case KeyCode
   Case vbKeyReturn, vbKeyEscape '// Allow escape to abort editing
      miKeyCode = KeyCode '// MLH
      bClearSelection = False

      If (muEditTrigger And lgAnyKey) Then
         KeyCode = 0

      Else

         If KeyCode = vbKeyEscape Then
            txtEdit.Visible = False
            mbEditPending = False
            KeyCode = 0
         Else
            If ToggleEdit() Then KeyCode = 0
         End If

      End If

   Case vbKeyInsert '// MLH
      miKeyCode = KeyCode

      If mbAllowInsert And Not mbEditPending Then
         RaiseEvent BeforeInsert(mRow, bCancel)

         If Not bCancel Then
            If mRow = C_NULL_RESULT Then mRow = 0
            Call AddItem("", mRow, , True)
            mRow = mRow + 1
            Call RowColSet(mRow - 1)
            RaiseEvent AfterInsert(mRow)
            bRedraw = True
         End If

      End If
      KeyCode = 0

   Case vbKeyDelete '// MLH
      If mbAllowDelete And Not mbEditPending Then
         If Not mRowCount = C_NULL_RESULT Then
            If muFocusRectMode = lgCol Then '// FocusRect = Col

               If Not (mCol = C_NULL_RESULT) Then
                  If Not (mCols(mColPtr(mCol)).nType = lgBoolean Or mCols(mColPtr(mCol)).bLocked) Then
                     RaiseEvent BeforeEdit(mRow, mColPtr(mCol), bCancel)
                     RaiseEvent BeforeDelete(mRow, bCancel)
   
                     If Not bCancel Then
                        mudtTotalsVal(mColPtr(mCol)) = mudtTotalsVal(mColPtr(mCol)) - rVal(mItems(mRowPtr(mRow)).Cell(mCol).sValue) '// MLH - New
                        mItems(mRowPtr(mRow)).Cell(mColPtr(mCol)).sValue = vbNullString
                        RaiseEvent AfterEdit(mRow, mColPtr(mCol), strTemp, bCancel)
                        RaiseEvent AfterDelete
                        bRedraw = True
                     End If
                  End If
               End If
               
            Else '// FocusRect = Row or None
               RaiseEvent BeforeDelete(mRow, bCancel)

               If Not bCancel Then
                  Call RemoveItem
                  RaiseEvent AfterDelete
                  bRedraw = True
               End If

            End If
         End If
         
         If lNewRow > mRowCount Then lNewRow = C_NULL_RESULT
         KeyCode = 0
      End If

   Case vbKeyF2
      miKeyCode = KeyCode '// MLH
      bClearSelection = False

      If (muEditTrigger And lgF2Key) Then
         If ToggleEdit() Then
            KeyCode = 0
         End If
      End If

   Case vbKeySpace
      bClearSelection = False
      If mbCheckboxes Then '// Row CheckMark
         mbIgnoreKeyPress = True

         bRedraw = True
         SetFlag mItems(mRowPtr(mRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mRow)).nFlags, lgFLChecked)
         RaiseEvent RowChecked(mRow)
         KeyCode = 0

      Else '// MLH - Cell CheckMark
         If Not mCol = C_NULL_RESULT Then
            If IsEditable() And mCols(mColPtr(mCol)).nType = lgBoolean Then
               bRedraw = True
               RaiseEvent BeforeEdit(mRow, mColPtr(mCol), bCancel)

               If Not bCancel Then
                  bState = (mItems(mRowPtr(mRow)).Cell(mColPtr(mCol)).nFlags And lgFLChecked)
                  bState = Not bState
                  SetFlag mItems(mRowPtr(mRow)).Cell(mColPtr(mCol)).nFlags, lgFLChecked, bState
                  mItems(mRowPtr(mRow)).Cell(mColPtr(mCol)).sValue = CStr(bState)
                  CellChanged(mRow, mCol) = True
                  strTemp = CStr(bState)
                  RaiseEvent AfterEdit(mRow, mColPtr(mCol), strTemp, bCancel)
                  DrawGrid bRedraw
                  KeyCode = 0
               End If

            End If
         End If
      End If

   Case vbKeyA
      bClearSelection = False

      If (Shift And vbCtrlMask) And muMultiSelect Then
         mbIgnoreKeyPress = True

         SetSelection True
         RaiseEvent SelectionChanged
         KeyCode = 0
      End If

   Case vbKeyC '// Allow Ctrl+C for copy '// MLH
      If (Shift And vbCtrlMask) Then
         mbIgnoreKeyPress = True
      End If

   Case vbKeyUp
      miKeyCode = KeyCode '// MLH

      If (Shift And vbShiftMask) And muMultiSelect Then
         bClearSelection = False
      End If

      If UpdateCell() Then
         'lNewRow = NavigateUp()
         lNewRow = CheckForLockedRow(True) '// MLH - New

         If Not lNewRow = C_NULL_RESULT Then
            If lNewRow < SBValue(efsVertical) Or lNewRow > SBValue(efsVertical) + RowsVisible() Then
               If SBVisible(efsVertical) Then
                  SBValue(efsVertical) = lNewRow
               Else
                  lNewRow = mRow
               End If

            End If
         End If
         KeyCode = 0
      End If

   Case vbKeyDown
      miKeyCode = KeyCode '// MLH

      If (Shift And vbShiftMask) And muMultiSelect Then
         bClearSelection = False
      End If

      If UpdateCell() Then
         'lNewRow = NavigateDown()
         lNewRow = CheckForLockedRow(False) '// MLH - New

         If Not lNewRow = C_NULL_RESULT Then
            If lNewRow < SBValue(efsVertical) Or lNewRow >= SBValue(efsVertical) + RowsVisible() Then
               If SBVisible(efsVertical) Then
                  SBValue(efsVertical) = lNewRow
               Else
                  lNewRow = mRow
               End If

            End If
         End If
         KeyCode = 0
      End If

   Case vbKeyLeft
      If mbEditPending Then '// MLH
         '// terminate edit if did not enter using F2 or DblClick

         If Not (miKeyCode = vbKeyF2) Then
            If txtEdit.SelStart = 0 Then '// at the begining of the string?
               miKeyCode = KeyCode
               If ToggleEdit() Then KeyCode = 0
            End If

         End If
      End If

      If Not mbEditPending Then
         lNewCol = NavigateLeft()
'         '// MLH - Make sure row is visible
'         If mRow < SBValue(efsVertical) Or mRow > SBValue(efsVertical) + RowsVisible() Then
'            SBValue(efsVertical) = mRow
'         End If

         miKeyCode = 0
         KeyCode = 0
      End If

   Case vbKeyRight
      If mbEditPending Then '// MLH
         '// terminate edit if did not enter using F2 or DblClick
         If Not (miKeyCode = vbKeyF2) Then
            If txtEdit.SelStart = Len(txtEdit.Text) Then '// at the end of the string?
               miKeyCode = KeyCode
               If ToggleEdit() Then KeyCode = 0
            End If

         End If
      End If

      If Not mbEditPending Then
         lNewCol = NavigateRight()
'         '// MLH - Make sure row is visible
'         If lNewCol < SBValue(efsHorizontal) Then ' Or mCol > SBValue(efsHorizontal) + RowsVisible() Then
'            SBValue(efsHorizontal) = lNewCol
'         End If

         miKeyCode = 0
         KeyCode = 0
      End If

   Case vbKeyPageUp
      miKeyCode = KeyCode '// MLH
      KeyCode = 0

      If Not mbEditPending Then
         If UpdateCell() Then
            If mRow > 0 Then
               'lNewRow = (mRow - RowsVisible()) + 1 '// MLH - prevent selection change
               lNewRow = (SBValue(efsVertical) - RowsVisible()) + 1
               If lNewRow < 0 Then lNewRow = 0
               bRedraw = True
               SBValue(efsVertical) = lNewRow
            End If

         End If
      End If

   Case vbKeyPageDown
      miKeyCode = KeyCode '// MLH
      KeyCode = 0

      If Not mbEditPending Then
         If UpdateCell() Then
            If mRow < mRowCount Then
               'lNewRow = (mRow + RowsVisible()) - 1 '// MLH - prevent selection change
               lNewRow = (SBValue(efsVertical) + RowsVisible()) - 1
               If lNewRow > mRowCount Then lNewRow = mRowCount
               bRedraw = True
               SBValue(efsVertical) = lNewRow
            End If

         End If
      End If

   Case vbKeyHome
      miKeyCode = KeyCode '// MLH

      If Shift And vbShiftMask Then
         If UpdateCell() Then
            If muMultiSelect Then
               bClearSelection = False

               SetSelection False
               SetSelection True, 1, mRow
               RaiseEvent SelectionChanged
            End If

            lNewRow = 0
            SBValue(efsVertical) = SBMin(efsVertical)
            KeyCode = 0
         End If

      ElseIf Shift And vbCtrlMask Then

         If UpdateCell() Then
            lNewRow = 0
            SBValue(efsVertical) = SBMin(efsVertical)
            KeyCode = 0
         End If

      ElseIf Not mbEditPending Then
         lNewCol = NavigateRight(True)
         SBValue(efsHorizontal) = SBMin(efsHorizontal)
         KeyCode = 0
      End If

   Case vbKeyEnd
      miKeyCode = KeyCode '// MLH

      If Shift And vbShiftMask Then
         If UpdateCell() Then
            If muMultiSelect Then
               bClearSelection = False

               SetSelection False
               SetSelection True, mRow, mRowCount
               RaiseEvent SelectionChanged
            End If

            lNewRow = mRowCount
            SBValue(efsVertical) = SBMax(efsVertical)
            KeyCode = 0
         End If

      ElseIf Shift And vbCtrlMask Then

         If UpdateCell() Then
            lNewRow = mRowCount
            SBValue(efsVertical) = SBMax(efsVertical)
            KeyCode = 0
         End If

      ElseIf Not mbEditPending Then
         lNewCol = NavigateLeft(True)
         SBValue(efsHorizontal) = SBMax(efsHorizontal)
         KeyCode = 0
      End If

   'Case Else
   '// MLH - Moved to KeyPress
   '      If Not mbEditPending Then
   '         If (muEditTrigger And lgAnyKey) Then
   '            bClearSelection = False
   '
   '            If ToggleEdit() Then
   '               KeyCode = 0
   '            End If
   '
   '         End If
   '      End if

   End Select

   '// Restore redraw state to user selected
   SetRedrawState True

   If KeyCode = 0 Then

      If Not (miKeyCode = vbKeyPageDown Or miKeyCode = vbKeyPageUp) Then '// MLH - prevent selection change Page Up/Down

         '// Do we want to clear selection?
         'If muFocusRectMode = lgRow Then lNewCol = 0
         If Not (lNewRow = C_NULL_RESULT) Then  '// MLH

            If bClearSelection And Not (mRow = lNewRow) Then
               If SetSelection(False) Then bRedraw = True
            End If

            'End If
            'If lNewRow > C_NULL_RESULT Then '// MLH - same If statment as above

            If Not mItems(mRowPtr(lNewRow)).nFlags And lgFLSelected Then
               bRedraw = True
               SetFlag mItems(mRowPtr(lNewRow)).nFlags, lgFLSelected, True
               RaiseEvent SelectionChanged
            End If

         End If

         If bRedraw Or SetRowCol(lNewRow, lNewCol, True) Then
            DrawGrid mbRedraw
         End If

      Else '// prevent selection change Page Up/Down
         If bRedraw Then
            DrawGrid mbRedraw
         End If

      End If
   End If

   RaiseEvent KeyDown(iKeyCode, Shift) '// MLH

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   '---------------------------------------------------------------------------------
   '// Purpose: This will find the Item that contains a Cell with text that is >= to the text typed. Each
   '// character entered is appended to the previous one if the time interval is less than 1 second.
   '//
   '// Key searching is disabled if the Grid is Disabled, and Edit is in progress or the KeyPress event is
   '// in an Ignore State (setting the SearchColumn to -1 will also prevent searches).
   '---------------------------------------------------------------------------------

  Dim lResult As Long
  Dim bEatKey As Boolean

   If mbEnabled Then

      '// MLH - Edit with any key
      If (muEditTrigger And lgAnyKey) Then
         If Not mbEditPending Then
            If ToggleEdit() Then
               Call txtEdit_KeyPress(KeyAscii) '// Check Formatting

               If KeyAscii Then
                  If Not (KeyAscii = Asc(vbCr)) Then
                     txtEdit.Text = Chr$(KeyAscii)
                     txtEdit.SelStart = 1
                  End If

               End If

            End If

         Else
            Select Case KeyAscii
            Case vbKeyEscape
               txtEdit.Visible = False
               mbEditPending = False
               bEatKey = True
               KeyAscii = 0

            Case vbKeyReturn
               If ToggleEdit(True) Then '// MLH
                  bEatKey = True
                  KeyAscii = 0
               End If
            End Select
         
         End If
      End If

      '// Used to prevent a beep
      If (muEditTrigger And lgEnterKey) And (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
         KeyAscii = 0
         bEatKey = True

      ElseIf Not mbIgnoreKeyPress And Not mbEditPending Then

         If IsCharAlphaNumeric(KeyAscii) Then
            If (GetTickCount() - mlTime) < 1000 Then
               msCode = msCode & Chr$(KeyAscii)
            Else
               msCode = Chr$(KeyAscii)
            End If

            mlTime = GetTickCount()

            lResult = FindItem(msCode, mSearchColumn, lgSMNavigate)

            If lResult > C_NULL_RESULT Then
               '// MLH
               If lResult > SBMax(efsVertical) Then
                  SBValue(efsVertical) = SBMax(efsVertical)
               Else
                  SBValue(efsVertical) = lResult
               End If

               SetRowCol lResult, mCol, True
               DrawGrid mbRedraw
               '''Call TopRow(lResult)'// MLH - not needed
            End If

         End If
      End If

      If Not bEatKey Then RaiseEvent KeyPress(KeyAscii)
   End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   mbIgnoreKeyPress = False

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

'''Private Sub UserControl_LostFocus()
'''
'''   bLostFocus = True
'''   If Not bLostFocus Then
'''      '// MLH
'''      If mbEditPending Then
'''         Call UpdateCell
'''      End If
'''
'''      If mbHideSelection Then
'''         DrawGrid mbRedraw, True
'''      End If
'''   End If
'''   bLostFocus = False
'''
'''End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim r                 As Rect
  Dim lngC              As Long
  Dim lngR              As Long
  Dim bCancel           As Boolean
  Dim bProcessed        As Boolean
  Dim bRedraw           As Boolean
  Dim bSelectionChanged As Boolean
  Dim bState            As Boolean
  '''Dim strTemp           As String

   picTooltip.Visible = False
   mbCancelShow = True
   mLRLocked = False
   mLCLocked = False

   '''If Not mLocked And Not (Button = 0) And (mRowCount >= 0) Then '// MLH - mLocked Not Used

   If Not (Button = 0) And (mRowCount >= 0) Then
      miKeyCode = 0
      mScrollAction = C_SCROLL_NONE

      lngC = GetColFromX(X)

      If mResizeRow = C_NULL_RESULT Then
         lngR = GetRowFromY(Y)
      Else
         lngR = mResizeRow
      End If

      mMouseDownX = X

      If Button = vbLeftButton Then
         '// Prevent Locked row from getting focus '// MLH
         If Not lngR = C_NULL_RESULT Then
            If mItems(mRowPtr(lngR)).nFlags And lgFLlocked Then
               mLRLocked = True
               mLCLocked = True
               Exit Sub
            End If

         End If

         mMouseDownRow = lngR

         Call SetCapture(UserControl.hWnd)
         mbMouseDown = True

         If Y < mR.HeaderHeight Then
            If mbAllowColumnSort Or mbAllowColumnSwap Or mbAllowColumnDrag Then '// MLH

               If Not (UserControl.MousePointer = vbSizeWE) Then
                  mMouseDownCol = lngC

                  If Not (mMouseDownCol = C_NULL_RESULT) Then

                     With UserControl
                        DrawHeader mMouseCol, lgDOWN
                        .Refresh
                     End With

                  End If
               End If
            End If

         ElseIf mMouseDownRow > C_NULL_RESULT Then

            If UpdateCell() Then
               If mbCheckboxes And (X <= C_RIGHT_CHECKBOX) Then '// Row CheckMark
                  bRedraw = True
                  mbMouseDown = False
                  SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked)
                  RaiseEvent RowChecked(mMouseDownRow)

               Else
                  If lngC > C_NULL_RESULT Then
                     If IsEditable() And mCols(mColPtr(lngC)).nType = lgBoolean Then '// Cell CheckMark
                        SetItemRect mMouseDownRow, lngC, RowTop(mMouseDownRow), r, lgRTCheckBox

                        If (X >= r.Left) And (Y >= r.Top) And (X <= r.Left + mR.CheckBoxSize) And (Y <= r.Top + mR.CheckBoxSize) Then
                           bRedraw = True
                           RaiseEvent BeforeEdit(mMouseDownRow, mColPtr(lngC), bCancel)

                           If Not bCancel Then
                              bState = (mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lngC)).nFlags And lgFLChecked)
                              bState = Not bState
                              
                              RaiseEvent AfterEdit(mMouseDownRow, mColPtr(lngC), CStr(bState), bCancel)  '// MLH
                              If Not bCancel Then
                                 SetFlag mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lngC)).nFlags, lgFLChecked, bState
                                 mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lngC)).sValue = CStr(bState)  '// MLH
                                 CellChanged(mMouseDownRow, lngC) = True '// MLH
                              End If
                           End If '// Not Canceled AfterEdit
                        End If '// Not Canceled BeforeEdit
                     End If
                  End If

                  If Not bProcessed Then
                     bState = (mItems(mRowPtr(mMouseDownRow)).nFlags And lgFLSelected)

                     If muMultiSelect Then '// MLH - Added New Options
                        If (Shift And vbCtrlMask) Or muMultiSelect = lgMultiLatch Then '// Latch Mode
                           SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                           bSelectionChanged = True

                        ElseIf (Shift And vbShiftMask) Then
                           bSelectionChanged = SetSelection(False) Or SetSelection(True, mRow, mMouseDownRow)

                        Else
                           SetSelection False
                           SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                           bSelectionChanged = True
                        End If

                     Else
                        If Shift And vbCtrlMask Then
                           SetSelection False
                           SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                           bSelectionChanged = True
                        ElseIf Not bState Then
                           SetSelection False
                           SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                           bSelectionChanged = True
                        End If
                     End If
                  End If

                  bRedraw = bRedraw Or SetRowCol(mMouseDownRow, lngC)
               End If

               If Not lngC = C_NULL_RESULT Then
                  If bRedraw Or mCols(mColPtr(lngC)).nType = lgButton Then
                     DrawGrid mbRedraw
                  End If
               End If
            End If   '// UpdateCell()
         End If      '// Row Not Locked

      Else '// Right Button
         mMouseDownRow = lngR

         If Not lngR = C_NULL_RESULT Then
            If muMultiSelect = lgSingleSelect Or (mItems(mRowPtr(lngR)).nFlags And lgFLSelected) = False Then
               If mMouseDownRow > C_NULL_RESULT Then
                  If UpdateCell() Then
                     SetRowCol mMouseDownRow, lngC
                     bSelectionChanged = SetSelection(False) Or SetSelection(True, mMouseDownRow, mMouseDownRow)
                     DrawGrid mbRedraw
                  End If
               End If
            End If
         End If
      End If

      If bSelectionChanged Then
         RaiseEvent SelectionChanged
      End If

   End If

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim lCount            As Long
  Dim lngTmpVal         As Long
  Dim lCol              As Long
  Dim lRow              As Long
  Dim nMove             As lgMoveControlEnum
  Dim nPointer          As Integer
  Dim bSelectionChanged As Boolean

   mLCLocked = False

   If mbIgnoreMove Then Exit Sub '// MLH

   '''If Not mLocked And (mRowCount >= 0) Then '// MLH - mLocked Not Used
   If Not (mRowCount = C_NULL_RESULT) Then
      lCol = GetColFromX(X)
      lRow = GetRowFromY(Y)

      If Not (mMouseCol = lCol) Or Not (mMouseRow = lRow) Then
         picTooltip.Visible = False
         mMouseCol = lCol
         mMouseRow = lRow

         '// Tool Tip
         If Not (lRow = C_NULL_RESULT) Then
            Call ShowCompleteCell(mMouseRow, mMouseCol)
         Else
            mbCancelShow = True
         End If
      End If

      '---------------------------------------------------------------------------------
      If mbAllowColumnSort Or mbAllowColumnSwap Or mbAllowColumnDrag Then '// MLH
         '// Header button tracking
         If Not (mMouseDownCol = C_NULL_RESULT) Then
            If (mMouseDownCol = mMouseCol) And (MouseRow = C_NULL_RESULT) Then
               Call DrawHeader(mMouseCol, lgDOWN)
            Else
               Call DrawHeader(mMouseDownCol, lgNormal)
            End If

            UserControl.Refresh
         End If

         '// Hot tracking
         If mbAllowColumnHover And (Button = 0) Then
            If Y < mR.HeaderHeight Then
               '// Do we need to draw a new "hot" header?
               If Not (mMouseCol = mHotColumn) Then
                  Call DrawHeaderRow
                  Call DrawHeader(mMouseCol, lghot)
                  mHotColumn = mMouseCol
               End If

            ElseIf Not (mHotColumn = C_NULL_RESULT) Then
               '// We have a previous "hot" header to clear
               Call DrawHeaderRow
            End If
         End If
      End If

      '---------------------------------------------------------------------------------

      If (Button = vbLeftButton) Then
         If Not (mResizeRow = C_NULL_RESULT) Then '// MLH - New
            '// We are resizing a Row
            lngTmpVal = Y - mlResizeY

            If lngTmpVal > mMinRowHeight Then
               mItems(mRowPtr(mResizeRow)).lHeight = lngTmpVal
               DrawGrid mbRedraw
            End If

         ElseIf Not (mResizeCol = C_NULL_RESULT) Then
            '// We are resizing a Column
            lngTmpVal = X - mlResizeX

            If lngTmpVal > C_SIZE_VARIANCE Then
               mCols(mColPtr(mResizeCol)).lWidth = lngTmpVal
               mCols(mColPtr(mResizeCol)).dCustomWidth = ScaleX(mCols(mColPtr(mResizeCol)).lWidth, vbPixels, muScaleUnits)

               DrawGrid mbRedraw

               nMove = mCols(mColPtr(mResizeCol)).MoveControl
               RaiseEvent ColumnSizeChanged(mColPtr(mResizeCol), nMove)

               If mbEditPending Then
                  MoveEditControl nMove
               End If

            End If

         ElseIf (mMouseDownRow = C_NULL_RESULT) Then '// Mouse Row = Header
            '// draging or swapping column?
            If mbAllowColumnSwap Then
               Call DrawHeaderRow(True)

               If (mMouseDownCol > C_NULL_RESULT) And (mSwapCol = C_NULL_RESULT) Then
                  mSwapCol = mMouseDownCol
               ElseIf Not (mSwapCol = C_NULL_RESULT) Then
                  mCols(mColPtr(mSwapCol)).lX = mCols(mColPtr(mSwapCol)).lX - (mMouseDownX - X)
               End If

            ElseIf mbAllowColumnDrag Then
               Call DrawHeaderRow(True)

               If (mMouseDownCol > C_NULL_RESULT) And (mDragCol = C_NULL_RESULT) Then
                  mDragCol = mMouseDownCol
               ElseIf Not (mDragCol = C_NULL_RESULT) Then
                  mCols(mColPtr(mDragCol)).lX = mCols(mColPtr(mDragCol)).lX - (mMouseDownX - X)
               End If
            End If

         Else
            If mbMouseDown And Y < 0 Then
               '// Mouse has been dragged off the control
               ScrollList C_SCROLL_UP

            ElseIf mbMouseDown And Y > UserControl.ScaleHeight Then
               '// Mouse has been dragged off the control
               ScrollList C_SCROLL_DOWN

            ElseIf mbMouseDown And (Shift = 0) And (mMouseRow > C_NULL_RESULT) Then

               If mScrollAction = C_SCROLL_NONE Then
                  bSelectionChanged = SetSelection(False)

                  If muMultiSelect Then
                     SetSelection True, mMouseDownRow, mMouseRow
                  Else
                     SetSelection True, mMouseRow, mMouseRow
                  End If

                  If SetRowCol(mMouseRow, mMouseCol) Then
                     RaiseEvent SelectionChanged
                     DrawGrid mbRedraw
                  End If

               Else
                  mScrollAction = C_SCROLL_NONE
               End If

            End If
         End If

      ElseIf Button = 0 Then '// No button pressed
         'nPointer = vbDefault '// MLH - not needed
         '// Only check for resize cursor if no buttons depressed
         If (mMouseRow = C_NULL_RESULT) Then '// mouse on header row
            '// allow column resizing
            If mbAllowColumnResizing Then
               mbIgnoreMove = True
               mlResizeX = 0
               mResizeCol = C_NULL_RESULT

               'For lCount = SBValue(efsHorizontal) To UBound(mCols) '// MLH - was before FreezeAtCol
               For lCount = 0 To UBound(mCols)                       '// MLH - FreezeAtCol
                  If lCount >= SBValue(efsHorizontal) Or lCount <= mlngFreezeAtCol Then '// MLH - FreezeAtCol

                     If mCols(mColPtr(lCount)).bVisible Then
                        mlResizeX = mlResizeX + mCols(mColPtr(lCount)).lWidth

                        If (X < mlResizeX + C_SIZE_VARIANCE) And (X > mlResizeX - C_SIZE_VARIANCE) Then
                           nPointer = vbSizeWE
                           mResizeCol = lCount
                           mlResizeX = X - mCols(mColPtr(lCount)).lWidth '// MLH
                           Exit For
                        End If
                        'mlResizeX = mlResizeX + mCols(mColPtr(lCount)).lWidth'// MLH
                     End If
                  End If
               Next lCount

               mbIgnoreMove = False
            End If '// mbAllowColumnResizing

         Else '// NOT (mMouseRow = C_NULL_RESULT) '// MLH - allow row resizing
            If mbAllowRowResizing Then
               mbIgnoreMove = True
               mlResizeY = mR.HeaderHeight
               mResizeRow = C_NULL_RESULT

               For lCount = SBValue(efsVertical) To mRowCount
                  mlResizeY = mlResizeY + mItems(mRowPtr(lCount)).lHeight

                  If (Y < mlResizeY + C_SIZE_VARIANCE) And (Y > mlResizeY - C_SIZE_VARIANCE) Then
                     nPointer = vbSizeNS
                     mResizeRow = GetRowFromY(Y - (C_SIZE_VARIANCE * 2))
                     mlResizeY = Y - mItems(mRowPtr(mMouseRow)).lHeight
                     Exit For
                  End If

               Next lCount
               mbIgnoreMove = False
            End If '// mbAllowRowResizing
         End If    '// (mMouseRow = C_NULL_RESULT)

         With UserControl
            If Not (.MousePointer = nPointer) Then
               .MousePointer = nPointer
            End If
         End With

      End If '// Button = vbLeftButton
   End If    '// Not mRowCount = C_NULL_RESULT

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim r                As Rect
  Dim lCurrentMouseCol As Long
  Dim lCurrentMouseRow As Long
  Dim lTemp            As Long
  Dim lngC             As Long

   If (Button = vbLeftButton) Then
      Call ReleaseCapture

      lCurrentMouseCol = GetColFromX(X)
      lCurrentMouseRow = GetRowFromY(Y)

      If Button = vbLeftButton Then
         '// Prevent Locked row from getting focus '// MLH
         If Not lCurrentMouseRow = C_NULL_RESULT Then
            If mItems(mRowPtr(lCurrentMouseRow)).nFlags And lgFLlocked Then
               mLRLocked = True
               mLCLocked = True
               Call SelectedClearAll
               Call RowColSet
               RaiseEvent ColumnClick(mColPtr(lCurrentMouseCol))  '// MLH
               Exit Sub
            End If
         End If
      End If

      If Not (mSwapCol = C_NULL_RESULT) Then
         '// We swapped a Column
         If lCurrentMouseCol > C_NULL_RESULT Then
            lTemp = mColPtr(mSwapCol)
            mColPtr(mSwapCol) = mColPtr(lCurrentMouseCol)
            mColPtr(lCurrentMouseCol) = lTemp
            DrawGrid mbRedraw
         End If

      ElseIf Not (mDragCol = C_NULL_RESULT) Then
         If Not (mDragCol = lCurrentMouseCol) Then
            '// We moved a Column
            If Not (lCurrentMouseCol = C_NULL_RESULT) Then
               On Error Resume Next

               If mDragCol > lCurrentMouseCol Then
                  lTemp = mColPtr(mDragCol)

                  For lngC = mDragCol To lCurrentMouseCol Step -1
                     mColPtr(lngC) = mColPtr(lngC - 1)
                  Next lngC

                  mColPtr(lCurrentMouseCol) = lTemp

               Else
                  lTemp = mColPtr(mDragCol)

                  For lngC = mDragCol To lCurrentMouseCol
                     mColPtr(lngC) = mColPtr(lngC + 1)
                  Next lngC

                  mColPtr(lCurrentMouseCol) = lTemp
               End If

               DrawGrid mbRedraw
            End If
         End If
         
      ElseIf Not (mResizeCol = C_NULL_RESULT And mResizeRow = C_NULL_RESULT) Then '// MLH
         '// We resized a Column so reset Scrollbars
         SetScrollBars
         DrawGrid mbRedraw
         UserControl.MousePointer = vbDefault

      ElseIf (lCurrentMouseRow = C_NULL_RESULT) Then
         '// Sort requested from Column Header click
         If (lCurrentMouseCol = mMouseDownCol) And Not (mMouseDownCol = C_NULL_RESULT) Then
            If mbAllowColumnSort Then
               If (Shift And vbCtrlMask) And Not (mSortColumn = C_NULL_RESULT) Then
                  If Not (mSortSubColumn = mColPtr(mMouseDownCol)) Then    '// MLH
                     mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTNormal '// MLH
                  End If

                  mSortSubColumn = mMouseDownCol
                  Sort , mCols(mColPtr(mSortColumn)).nSortOrder

               Else
                  If Not (mSortColumn = mColPtr(mMouseDownCol)) Then       '// MLH
                     mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTNormal '// MLH
                     mSortSubColumn = C_NULL_RESULT
                  End If

                  mSortColumn = mMouseDownCol
                  If Not (mSortSubColumn = C_NULL_RESULT) Then
                     Sort , , , mCols(mColPtr(mSortSubColumn)).nSortOrder
                  Else
                     Sort
                  End If
               End If

               Call RowColSet(, mMouseDownCol)  '// MLH

            Else
               Call DrawHeaderRow
               'RaiseEvent ColumnClick(mColPtr(mMouseDownCol)) '// MLH
            End If

         End If

         If Not lCurrentMouseCol = C_NULL_RESULT Then          '// MLH
            RaiseEvent ColumnClick(mColPtr(lCurrentMouseCol))  '// MLH
         End If

      ElseIf Not (mMouseDownRow = C_NULL_RESULT) Then
         If IsValidRowCol(mMouseRow, mMouseCol) Then

            'If SetRowCol(mMouseRow, mMouseCol) Then '// MLH - not needed
            '   DrawGrid mbRedraw
            'End If

            'If Not (mCF(mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).nFormat).nImage = 0) Then'// MLH - (.nFormat?)
            If Not (mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).lImage = 0) Then
               '// Cell has an image
               SetItemRect mMouseRow, mMouseCol, RowTop(mMouseRow), r, lgRTImage
               '// has the cell's image been clicked?
               If (X >= r.Left) And (Y >= r.Top) And (X <= r.Left + mR.ImageWidth) And (Y <= r.Top + mR.ImageHeight) Then
                  RaiseEvent CellImageClick(mMouseRow, mColPtr(mMouseCol))
               End If

            ElseIf mbAllowWordWrap And (mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).nFlags And lgFLWordWrap) Then '// MLH
               '// Using Expand/Shrink Image in word wrapped rows
               If mExpandRowImage > 0 Then
                  SetItemRect mMouseRow, mMouseCol, RowTop(mMouseRow), r, lgRTImage

                  If (X >= r.Left) And (Y >= r.Top) And (X <= r.Left + mR.ImageWidth) And (Y <= r.Top + mR.ImageHeight) Then
                     If RowHeight(mMouseRow) = mR.TextHeight + (mMinVerticalOffset * 2) + 2 Then '// MLH
                        '// Restore to normal
                        RowHeight(mRow) = C_NULL_RESULT
                     Else
                        '// Shrink to minimum
                        RowHeight(mRow) = mR.TextHeight + (mMinVerticalOffset * 2) + 2 '// MLH
                     End If

                  End If
               End If '// mExpandRowImage=0
            End If

            RaiseEvent ColumnClick(mColPtr(mMouseCol)) '// MLH
         End If '// IsValidRowCol

         '''RaiseEvent Click '// MLH

      Else
         Call DrawHeaderRow
      End If

   End If

   mbMouseDown = False

   '// Restore Button to normal state
   If Not (mMouseRow = C_NULL_RESULT Or mMouseCol = C_NULL_RESULT) Then '// MLH - New
      If mCols(mColPtr(mMouseCol)).nType = lgButton Then
         DrawGrid mbRedraw
         RaiseEvent CellButtonClick(mMouseRow, mColPtr(mMouseCol))
      End If

   ElseIf mMouseDownRow = C_NULL_RESULT Then '// clicked on header
      DrawGrid mbRedraw
   End If

   mMouseDownCol = C_NULL_RESULT
   mSwapCol = C_NULL_RESULT
   mDragCol = C_NULL_RESULT
   mResizeCol = C_NULL_RESULT
   mScrollAction = C_SCROLL_NONE
   
   If mMouseDownRow = C_NULL_RESULT Then Call DrawHeaderRow '// MLH - New

   RaiseEvent MouseUp(Button, Shift, X, Y)
   '// MLH
   'If Not lCurrentMouseRow = C_NULL_RESULT Then
   '   RaiseEvent Click
   'End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   '// Read the properties from the property bag - a good place to start the subclassing

   With PropBag
      '---------------------------------------------------------------------------------
      '// Appearance Properties
      mbApplySelectionToImages = .ReadProperty("ApplySelectionToImages", False)
      mBackColor = .ReadProperty("BackColor", vbWindowBackground)
      mBackColorBkg = .ReadProperty("BackColorBkg", vbApplicationWorkspace)
      mBackColorEdit = .ReadProperty("BackColorEdit", vbInfoBackground)
      'mBackColorFixed = .ReadProperty("BackColorFixed", vbButtonFace)'// MLH - Not Needed
      mBackColorSel = .ReadProperty("BackColorSel", vbHighlight)
      mForeColor = .ReadProperty("ForeColor", vbWindowText)
      mForeColorEdit = .ReadProperty("ForeColorEdit", vbInfoText)
      'mForeColorFixed = .ReadProperty("ForeColorFixed", vbButtonText)''''// MLH - Not Used
      mForeColorHdr = .ReadProperty("ForeColorHdr", vbWindowText)
      mForeColorSel = .ReadProperty("ForeColorSel", vbHighlightText)
      '''mForeColorTotals = .ReadProperty("ForeColorTotals", DEF_FORECOLORTOTALS) '// MLH
      mbColumnHeaders = .ReadProperty("ShowColumnHeaders", True) '// MLH
      mbCenterRowImage = .ReadProperty("CenterRowImage", True) '// MLH
      muSBOrienation = .ReadProperty("ScrollBars", ScrollBarOrienationEnum.Scroll_Both) '// MLH
      mMinVerticalOffset = .ReadProperty("MinVerticalOffset", DEF_MinVerticalOffset) '// MLH

      mBackColorEvenRows = .ReadProperty("BackColorEvenRows", &HEDEBE0) '// MLH
      mbBackColorEvenRowsE = .ReadProperty("BackColorEvenRowsEnabled", True) '// MLH

      mGridColor = .ReadProperty("GridColor", DEF_GRIDCOLOR)
      mProgressBarColor = .ReadProperty("ProgressBarColor", DEF_PROGRESSBARCOLOR)

      mbAlphaBlendSelection = .ReadProperty("AlphaBlendSelection", False)
      muBorderStyle = .ReadProperty("BorderStyle", DEF_BORDERSTYLE)
      mbDisplayEllipsis = .ReadProperty("DisplayEllipsis", True)
      mFocusRectColor = .ReadProperty("FocusRectColor", vbYellow)
      muFocusRectMode = .ReadProperty("FocusRectMode", DEF_FOCUSRECTMODE)
      muFocusRectStyle = .ReadProperty("FocusRectStyle", DEF_FOCUSRECTSTYLE)
      muGridLines = Abs(.ReadProperty("GridLines", lgGridLinesEnum.lgGrid_Both)) '// MLH
      mGridLineWidth = .ReadProperty("GridLineWidth", DEF_GRIDLINEWIDTH)
      muThemeColor = .ReadProperty("ThemeColor", lgThemeConst.Autodetect) '// MLH
      muThemeStyle = .ReadProperty("ThemeStyle", DEF_THEMESTYLE)
      mblnColumnHeaderSmall = .ReadProperty("ColumnHeaderSmall", False)

      mlngCustomColorFrom = .ReadProperty("CustomColorFrom", TranslateColor(&H404080)) '// MLH
      mlngCustomColorTo = .ReadProperty("CustomColorTo", TranslateColor(&HC0E0FF)) '// MLH
      mColumnHeaderLines = .ReadProperty("ColumnHeaderLines", 1) '// MLH
      msCaption = .ReadProperty("Caption", "") '// MLH
      muCaptionAlignment = .ReadProperty("CaptionAlignment", lgCaptionAlignmentEnum.lgAlignCenter) '// MLH
      UserControl.Appearance = .ReadProperty("Appearance", lgAppearanceEnum.Appear_3D) '// MLH
      muScrollBarStyle = .ReadProperty("ScrollBarStyle", ScrollBarStyleEnum.Style_Flat) '// MLH

      mbTotalsLineShow = .ReadProperty("TotalsLineShow", False) '// MLH - new
      mblnKeepForeColor = .ReadProperty("FocusRowHighlightKeepTextForecolor", False) '// MLH - new

      '---------------------------------------------------------------------------------
      '// Behaviour Properties
      mbAllowColumnResizing = .ReadProperty("AllowColumnResizing", False) '// MLH
      mbAllowRowResizing = .ReadProperty("AllowRowResizing", False) '// MLH - New
      mbAllowWordWrap = .ReadProperty("AllowWordWrap", False) '// MLH
      mbCheckboxes = .ReadProperty("Checkboxes", False)
      mbAllowColumnDrag = .ReadProperty("ColumnDrag", False) '// MLH
      mbAllowColumnSwap = .ReadProperty("ColumnSwap", False) '// MLH - renamed
      mbAllowColumnSort = .ReadProperty("ColumnSort", False)
      mbAllowEdit = .ReadProperty("Editable", False)
      mbAllowDelete = .ReadProperty("AllowDelete", False) '// MLH
      mbAllowInsert = .ReadProperty("AllowInsert", False) '// MLH
      muEditTrigger = .ReadProperty("EditTrigger", lgEditTriggerEnum.lgAnyF2DblCk)
      muEditMove = .ReadProperty("EditMove", lgEditMoveEnum.lgDontNone) '// MLH - New
      mbFullRowSelect = .ReadProperty("FullRowSelect", True)
      muFocusRowHighlightStyle = .ReadProperty("FocusRowHighlightStyle", lgFocusRowHighlightStyle.Solid) '// MLH - new
      mbHideSelection = .ReadProperty("HideSelection", False)
      mbAllowColumnHover = .ReadProperty("HotHeaderTracking", True)
      muMultiSelect = Abs(.ReadProperty("MultiSelect", 0)) '// MLH - Added New Options
      '''mbRedraw = .ReadProperty("Redraw", DEF_REDRAW)'// MLH - should always start as false
      mbScrollTrack = .ReadProperty("ScrollTrack", True)
      mbAutoToolTips = .ReadProperty("AutoToolTips", True)
      UserControl.Enabled = .ReadProperty("Enabled", True) '// MLH
      mlngFreezeAtCol = .ReadProperty("FreezeAtCol", -1) '// MLH - New

      '---------------------------------------------------------------------------------
      '// Miscellaneous Properties
      mCacheIncrement = .ReadProperty("CacheIncrement", DEF_CACHEINCREMENT)
      mbEnabled = .ReadProperty("Enabled", True)
      ''' mFormatString = .ReadProperty("FormatString", DEF_FORMATSTRING)'// MLH - function removed
      ''' mLocked = .ReadProperty("Locked", DEF_LOCKED)'// MLH - Not Used
      mMaxLineCount = .ReadProperty("MaxLineCount", DEF_MaxLineCount)
      muScaleUnits = .ReadProperty("ScaleUnits", vbTwips)
      mSearchColumn = .ReadProperty("SearchColumn", DEF_SEARCHCOLUMN)

      mMinRowHeightUser = .ReadProperty("MinRowHeight", DEF_MinRowHeight)
      'mMinRowHeight = ScaleY(mMinRowHeightUser, muScaleUnits, vbPixels) 'mMinRowHeightUser

      Set Font = .ReadProperty("Font", Ambient.Font)
   End With

   '---------------------------------------------------------------------------------
   '// Apply Settings
   Call SetThemeColor

   With UserControl
      .BackColor = mBackColorBkg
      .BorderStyle = muBorderStyle
      ucFontBold = .FontBold
      ucFontItalic = .FontItalic
      ucFontName = .FontName
   End With

   ''' If LenB(mFormatString) Then Clear '// MLH - function removed
   ''' FormatString = mFormatString      '// MLH - function removed

   Call Clear
   Call CreateRenderData

   '// sc_Subclass
   If Ambient.UserMode Then '// If running, not designing

      With UserControl
         Call sc_Subclass(.hWnd)
         Call sc_AddMsg(.hWnd, WM_KILLFOCUS)
         Call sc_AddMsg(.hWnd, WM_SETFOCUS)
         Call sc_AddMsg(.hWnd, WM_MOUSEWHEEL)
         Call sc_AddMsg(.hWnd, WM_MOUSEMOVE)
         Call sc_AddMsg(.hWnd, WM_MOUSELEAVE)
         Call sc_AddMsg(.hWnd, WM_MOUSEHOVER)
         Call sc_AddMsg(.hWnd, WM_HSCROLL)
         Call sc_AddMsg(.hWnd, WM_VSCROLL)

         If mbWinXP Then
            Call sc_AddMsg(.hWnd, WM_THEMECHANGED)
         End If

      End With

      SBCreate UserControl.hWnd
      SBStyle = Style_Regular

      SBLargeChange(efsHorizontal) = 5
      'SBSmallChange(efsHorizontal) = 1 '// MLH - Changed mSBSmallChangeHorz to Const

      SBLargeChange(efsVertical) = 5
      'SBSmallChange(efsVertical) = 1 '// MLH - Changed mSBSmallChangeVert to Const
   End If

End Sub

Private Sub UserControl_Resize()

   If Not (mSBhWnd = 0) Then
      '// make sure 1st row is visible
      If mRowCount Then
         SBValue(efsVertical) = 0
         SetScrollBars
      End If

      '// make sure selected row is visible
      If Not mRow = C_NULL_RESULT Then
         SBValue(efsVertical) = mRow
         '// SetScrollBars - called in Refresh
      End If

      Call Refresh
   End If

End Sub

Private Sub UserControl_Terminate()

   On Error Resume Next
   
   mblnDrwGrid = True '// MLH - prevent redraws during close
   mbRedraw = False   '// MLH - prevent redraws during close
   mbUserRedraw = False
   DoEvents

   If Not mClipRgn = 0 Then
      DeleteObject mClipRgn
   End If

   Call pSBClearUp
   DoEvents

UserControl_TerminateError:
   '// Clean up array data '// MLH
   Erase mCols
   Erase mItems
   Erase mColPtr
   Erase mRowPtr
   Erase mCF

   Call sc_Terminate '// MLH

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      '---------------------------------------------------------------------------------
      '// Appearance Properties
      Call .WriteProperty("Font", mFont, Ambient.Font)
      Call .WriteProperty("ApplySelectionToImages", mbApplySelectionToImages, False)
      Call .WriteProperty("BackColor", mBackColor, vbWindowBackground)
      Call .WriteProperty("BackColorBkg", mBackColorBkg, vbApplicationWorkspace)
      Call .WriteProperty("BackColorEdit", mBackColorEdit, vbInfoBackground)
      'Call .WriteProperty("BackColorFixed", mBackColorFixed, vbButtonFace)'// MLH - Not Needed
      Call .WriteProperty("BackColorSel", mBackColorSel, vbHighlight)
      Call .WriteProperty("ForeColor", mForeColor, vbWindowText)
      Call .WriteProperty("ForeColorEdit", mForeColorEdit, vbInfoText)
      'Call .WriteProperty("ForeColorFixed", mForeColorFixed, vbButtonText)''''// MLH - Not Used
      Call .WriteProperty("ForeColorHdr", mForeColorHdr, vbWindowText)
      Call .WriteProperty("ForeColorSel", mForeColorSel, vbHighlightText)
      ''' Call .WriteProperty("ForeColorTotals", mForeColorTotals, DEF_FORECOLORTOTALS) '// MLH - Not Used

      Call .WriteProperty("BackColorEvenRows", mBackColorEvenRows, &HEDEBE0) '// MLH
      Call .WriteProperty("BackColorEvenRowsEnabled", mbBackColorEvenRowsE, True) '// MLH
      Call .WriteProperty("CustomColorFrom", mlngCustomColorFrom, &H404080) '// MLH
      Call .WriteProperty("CustomColorTo", mlngCustomColorTo, &HC0E0FF) '// MLH

      Call .WriteProperty("GridColor", mGridColor, DEF_GRIDCOLOR)
      Call .WriteProperty("ProgressBarColor", mProgressBarColor, DEF_PROGRESSBARCOLOR)

      Call .WriteProperty("AlphaBlendSelection", mbAlphaBlendSelection, False)
      Call .WriteProperty("BorderStyle", muBorderStyle, DEF_BORDERSTYLE)
      Call .WriteProperty("DisplayEllipsis", mbDisplayEllipsis, True)
      Call .WriteProperty("FocusRectMode", muFocusRectMode, DEF_FOCUSRECTMODE)
      Call .WriteProperty("FocusRectColor", mFocusRectColor, vbYellow)
      Call .WriteProperty("FocusRectStyle", muFocusRectStyle, DEF_FOCUSRECTSTYLE)
      Call .WriteProperty("GridLines", muGridLines, lgGridLinesEnum.lgGrid_Both)
      Call .WriteProperty("GridLineWidth", mGridLineWidth, DEF_GRIDLINEWIDTH)
      Call .WriteProperty("ThemeColor", muThemeColor, lgThemeConst.Autodetect) '// MLH
      Call .WriteProperty("ThemeStyle", muThemeStyle, DEF_THEMESTYLE)
      Call .WriteProperty("ShowColumnHeaders", mbColumnHeaders, True) '// MLH
      Call .WriteProperty("CenterRowImage", mbCenterRowImage, True) '// MLH
      Call .WriteProperty("ScrollBars", muSBOrienation, ScrollBarOrienationEnum.Scroll_Both) '// MLH
      Call .WriteProperty("MinVerticalOffset", mMinVerticalOffset, DEF_MinVerticalOffset) '// MLH
      Call .WriteProperty("ColumnHeaderLines", mColumnHeaderLines, 1) '// MLH
      Call .WriteProperty("Caption", msCaption, "") '// MLH
      Call .WriteProperty("CaptionAlignment", muCaptionAlignment, lgCaptionAlignmentEnum.lgAlignCenter) '// MLH
      Call .WriteProperty("Appearance", UserControl.Appearance, lgAppearanceEnum.Appear_3D) '// MLH
      Call .WriteProperty("ColumnHeaderSmall", mblnColumnHeaderSmall, False) '// MLH
      Call .WriteProperty("ScrollBarStyle", muScrollBarStyle, ScrollBarStyleEnum.Style_Flat) '// MLH
      Call .WriteProperty("TotalsLineShow", mbTotalsLineShow, False) '// MLH
      Call .WriteProperty("FocusRowHighlightKeepTextForecolor", mblnKeepForeColor, False) '// MLH
      
      '---------------------------------------------------------------------------------
      '// Behaviour Properties
      Call .WriteProperty("AllowColumnResizing", mbAllowColumnResizing, False) '// MLH
      Call .WriteProperty("AllowRowResizing", mbAllowRowResizing, False) '// MLH
      Call .WriteProperty("AllowWordWrap", mbAllowWordWrap, False) '// MLH
      Call .WriteProperty("Checkboxes", mbCheckboxes, False)
      Call .WriteProperty("ColumnSwap", mbAllowColumnSwap, False) '// MLH
      Call .WriteProperty("ColumnDrag", mbAllowColumnDrag, False) '// MLH - renamed
      Call .WriteProperty("ColumnSort", mbAllowColumnSort, False)
      Call .WriteProperty("Editable", mbAllowEdit, False)
      Call .WriteProperty("AllowDelete", mbAllowDelete, False) '// MLH
      Call .WriteProperty("AllowInsert", mbAllowInsert, False) '// MLH
      Call .WriteProperty("EditTrigger", muEditTrigger, lgEditTriggerEnum.lgAnyF2DblCk)
      Call .WriteProperty("EditMove", muEditMove, lgEditMoveEnum.lgDontNone) '// MLH - New
      Call .WriteProperty("FullRowSelect", mbFullRowSelect, True)
      Call .WriteProperty("FocusRowHighlightStyle", muFocusRowHighlightStyle, lgFocusRowHighlightStyle.Solid) '// MLH - new
      Call .WriteProperty("HideSelection", mbHideSelection, False)
      Call .WriteProperty("HotHeaderTracking", mbAllowColumnHover, True)
      Call .WriteProperty("MultiSelect", muMultiSelect, 0) '// MLH - Added New Options
      '''Call .WriteProperty("Redraw", mbRedraw, DEF_REDRAW) '// MLH - should always start as false
      Call .WriteProperty("ScrollTrack", mbScrollTrack, True)
      Call .WriteProperty("AutoToolTips", mbAutoToolTips, True)
      Call .WriteProperty("Enabled", UserControl.Enabled, True) '// MLH
      Call .WriteProperty("FreezeAtCol", mlngFreezeAtCol, -1) '// MLH

      '---------------------------------------------------------------------------------
      '// Miscellaneous Properties
      Call .WriteProperty("CacheIncrement", mCacheIncrement, DEF_CACHEINCREMENT)
      Call .WriteProperty("Enabled", mbEnabled, True)
      ''' Call .WriteProperty("FormatString", mFormatString, DEF_FORMATSTRING)'// MLH - function removed
      ''' Call .WriteProperty("Locked", mLocked, DEF_LOCKED)'// MLH - Not Used
      Call .WriteProperty("MaxLineCount", mMaxLineCount, DEF_MaxLineCount)
      Call .WriteProperty("MinRowHeight", mMinRowHeightUser, DEF_MinRowHeight)
      Call .WriteProperty("ScaleUnits", muScaleUnits, vbTwips)
      Call .WriteProperty("SearchColumn", mSearchColumn, DEF_SEARCHCOLUMN)
   End With

End Sub

Public Function VisibleHeight() As Long

  Const SM_CXHTHUMB = 10 '// Width of scroll box on horizontal scroll bar
  Const SM_CXBORDER = 5  '// Width of no-sizable borders
  Dim lngBorder As Long

   '// MLH - new
   '// Purpose: return the grid's height minus the horizontal scroll bar height and border (if it is showing).

   If Not mbRedraw Then
      SetScrollBars
   End If

   If muBorderStyle = lgSingle Then
      lngBorder = ScaleY(GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips) * 4
   End If

   If SBVisible(efsHorizontal) Then
      VisibleHeight = UserControl.Height - ScaleY(GetSystemMetrics(SM_CXHTHUMB), vbPixels, vbTwips) - lngBorder
   Else
      VisibleHeight = UserControl.Height - lngBorder
   End If

End Function

Public Function VisibleWidth() As Long

  Const SM_CXHTHUMB = 10 '// Width of scroll box on horizontal scroll bar
  Const SM_CXBORDER = 5  '// Width of no-sizable borders
  Dim lngBorder As Long

   '// MLH - new
   '// Purpose: return the grid's width minus the vertical scroll bar width and border (if it is showing).

   If Not mbRedraw Then
      SetScrollBars
   End If

   If muBorderStyle = lgSingle Then
      lngBorder = ScaleX(GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips) * 4
   End If

   If SBVisible(efsVertical) Then
      VisibleWidth = UserControl.Width - ScaleX(GetSystemMetrics(SM_CXHTHUMB), vbPixels, vbTwips) - lngBorder
   Else
      VisibleWidth = UserControl.Width - lngBorder
   End If
   
   VisibleWidth = VisibleWidth - Screen.TwipsPerPixelX

End Function

'-The following routines are exclusively for the sc_ subclass routines----------------------------
'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)

  Dim nCount As Long                         'Table entry count
  Dim nBase  As Long                         'Remember z_ScMem
  Dim i      As Long                         'Loop index

   nBase = z_ScMem                           'Remember z_ScMem so that we can restore its
   '   value on exit
   z_ScMem = zData(nTable)                   'Map zData() to the specified table

   If uMsg = ALL_MESSAGES Then               'If ALL_MESSAGES are being added to the table...
      nCount = ALL_MESSAGES                  'Set the table entry count to ALL_MESSAGES
   
   Else
      nCount = zData(0)                      'Get the current table entry count
      If nCount >= MSG_ENTRIES Then          'Check for message table overflow
         zError "zAddMsg", "Message table overflow. Either increase" & _
                           " the value of Const MSG_ENTRIES or use ALL_MESSAGES" & _
                           " instead of specific message values"
         GoTo Bail
      End If

      For i = 1 To nCount                    'Loop through the table entries
         If zData(i) = 0 Then                'If the element is free...
            zData(i) = uMsg                  'Use this element
            GoTo Bail                        'Bail
         ElseIf zData(i) = uMsg Then         'If the message is already in the table...
            GoTo Bail                        'Bail
         End If

      Next i                                 'Next message table entry

      nCount = i                             'On drop through: i = nCount + 1, the new table entry count
      zData(nCount) = uMsg                   'Store the message in the appended table entry
   End If

   zData(0) = nCount                         'Store the new table entry count
Bail:
   z_ScMem = nBase                           'Restore the value of z_ScMem

End Sub

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long

  Dim o As Long                                                   'Object pointer
  Dim i As Long                                                   'vTable entry counter
  Dim j As Long                                                   'vTable address
  Dim n As Long                                                   'Method pointer
  Dim b As Byte                                                   'First method byte
  Dim m As Byte                                                   'Known good first method byte

   o = ObjPtr(oCallback)                                          'Get the callback object's address
   GetMem4 o, j                                                   'Get the address of the callback object's vTable
   j = j + &H7A4                                                  'Increment to the the first user entry for a usercontrol
   GetMem4 j, n                                                   'Get the method pointer
   GetMem1 n, m                                                   'Get the first method byte... &H33 if pseudo-code, &HE9 if native
   j = j + 4                                                      'Bump to the next vtable entry
   
   For i = 1 To 511                                               'Loop through a 'sane' number of vtable entries
      GetMem4 j, n                                                'Get the method pointer
      
      If IsBadCodePtr(n) Then                                     'If the method pointer is an invalid code address
         GoTo vTableEnd                                           'We've reached the end of the vTable, exit the for loop
      End If
      
      GetMem1 n, b                                                'Get the first method byte
      
      If b <> m Then                                              'If the method byte doesn't matche the known good value
         GoTo vTableEnd                                           'We've reached the end of the vTable, exit the for loop
      End If
      
      j = j + 4                                                   'Bump to the next vTable entry
   Next i                                                         'Bump counter

   Debug.Assert False                                             'Halt if running under the VB IDE
   Err.Raise vbObjectError, "zAddressOf", "Ordinal not found"     'Raise error if running compiled
   '
vTableEnd:                                                        'We've hit the end of the vTable
   GetMem4 j - (nOrdinal * 4), n                                  'Get the method pointer for the specified
   '   ordinal
   zAddressOf = n                                                 'Address of the callback ordinal

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

   RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)

   RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)

  Dim nCount As Long                         'Table entry count
  Dim nBase  As Long                         'Remember z_ScMem
  Dim i      As Long                         'Loop index

   nBase = z_ScMem                           'Remember z_ScMem so that we can restore its value on exit
   z_ScMem = zData(nTable)                   'Map zData() to the specified table

   If uMsg = ALL_MESSAGES Then               'If ALL_MESSAGES are being deleted from the table...
      zData(0) = 0                           'Zero the table entry count
   Else
      nCount = zData(0)                      'Get the table entry count

      For i = 1 To nCount                    'Loop through the table entries
         If zData(i) = uMsg Then             'If the message is found...
            zData(i) = 0                     'Null the msg value -- also frees the element for re-use
            GoTo Bail                        'Bail
         End If

      Next i                                 'Next message table entry

      zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
   End If

Bail:
   z_ScMem = nBase                           'Restore the value of z_ScMem

End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)

   App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
   MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long

   zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
   Debug.Assert zFnAddr                                     'In the IDE, validate that the procedure address was located

End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

   If z_Funk Is Nothing Then                                   'Ensure that subclassing has been started
      zError "zMap_hWnd", "Subclassing hasn't been started"
   Else
      On Error GoTo Catch                                      'Catch unsubclassed window handles
      z_ScMem = z_Funk("h" & lng_hWnd)                         'Get the thunk address
      zMap_hWnd = z_ScMem
   End If

   Exit Function                                               'Exit returning the thunk address

Catch:
   zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

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
  Dim eBar        As EFSScrollBarConstants
  Dim lV          As Long
  Dim lSC         As Long
  Dim lScrollCode As Long
  Dim tSI         As SCROLLINFO
  Dim lHSB        As Long
  Dim lVSB        As Long
  Dim bRedraw     As Boolean


   Select Case uMsg
   Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL

      miKeyCode = 0
      lScrollCode = (wParam And &HFFFF&)
      lHSB = SBValue(efsHorizontal)
      lVSB = SBValue(efsVertical)
      bRedraw = False

      Select Case uMsg
      Case WM_HSCROLL '//  Get the scrollbar type
         bRedraw = True
         eBar = efsHorizontal
         picTooltip.Visible = False '// MLH

      Case WM_VSCROLL
         bRedraw = True
         eBar = efsVertical
         picTooltip.Visible = False '// MLH

      Case Else     '// WM_MOUSEWHEEL
         bRedraw = True
         eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
         lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
         picTooltip.Visible = False '// MLH
      End Select

      Select Case lScrollCode
      Case SB_THUMBTRACK
         '//  Is vertical/horizontal?
         pSBGetSI eBar, tSI, SIF_TRACKPOS
         SBValue(eBar) = tSI.nTrackPos

         bRedraw = mbScrollTrack

      Case SB_LEFT, SB_BOTTOM
         SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))

      Case SB_RIGHT, SB_TOP
         SBValue(eBar) = SBMin(eBar)

      Case SB_LINELEFT, SB_LINEUP
         If SBVisible(eBar) Then
            lV = SBValue(eBar)                     '// MLH
            'lSC = 1                               '// MLH
            If (eBar = efsHorizontal) Then
               lSC = mCol                          '// MLH
               mCol = SBValue(eBar)                '// MLH
               lV = NavigateLeft                   '// MLH
               If lV > mCol Then lV = 0            '// MLH - Prevent wrapping
               mCol = lSC                          '// MLH
               lSC = 0                             '// MLH

            Else
               lV = SBValue(eBar)                  '// MLH
               If uMsg = WM_MOUSEWHEEL Then        '// MLH
                  'lSC = 3                         '// MLH
                  lSC = mRow                       '// MLH
                  mRow = SBValue(eBar)             '// MLH
                  lV = NavigateUp(3, True)         '// MLH
                  If lV > mRow Then lV = 0         '// MLH - Prevent wrapping
                  mRow = lSC                       '// MLH
                  lSC = 0                          '// MLH
                  
               Else
                  'lSC = 1                         '// MLH
                  lSC = mRow                       '// MLH
                  mRow = SBValue(eBar)             '// MLH
                  lV = NavigateUp(1, True)         '// MLH
                  If lV > mRow Then lV = 0         '// MLH - Prevent wrapping
                  mRow = lSC                       '// MLH
                  lSC = 0                          '// MLH
                  
               End If
            End If

            If (lV - lSC < SBMin(eBar)) Then
               SBValue(eBar) = SBMin(eBar)
            Else
               SBValue(eBar) = lV - lSC
            End If

         End If

      Case SB_LINERIGHT, SB_LINEDOWN
         If SBVisible(eBar) Then
            lV = SBValue(eBar)                     '// MLH
            'lSC = 1                               '// MLH
            If (eBar = efsHorizontal) Then
               lSC = mCol                          '// MLH
               mCol = SBValue(eBar)                '// MLH
               lV = NavigateRight                  '// MLH
               If lV < mCol Then lV = SBMax(eBar)  '// MLH - Prevent wrapping
               mCol = lSC                          '// MLH
               lSC = 0                             '// MLH

            Else
               If uMsg = WM_MOUSEWHEEL Then           '// MLH
                  'lSC = 3 'C_SBSmallChangeVert + 2   '// MLH
                  lSC = mRow                          '// MLH
                  mRow = SBValue(eBar)                '// MLH
                  lV = NavigateDown(3, True)          '// MLH
                  If lV < mRow Then lV = SBMax(eBar)  '// MLH - Prevent wrapping
                  mRow = lSC                          '// MLH
                  lSC = 0                             '// MLH
                  
               Else
                  'lSC = 1  'C_SBSmallChangeVert      '// MLH
                  lSC = mRow                          '// MLH
                  mRow = SBValue(eBar)                '// MLH
                  lV = NavigateDown(1, True)          '// MLH
                  If lV < mRow Then lV = SBMax(eBar)  '// MLH - Prevent wrapping
                  mRow = lSC                          '// MLH
                  lSC = 0                             '// MLH
               End If
            End If

            If (lV + lSC > SBMax(eBar)) Then
               SBValue(eBar) = SBMax(eBar)
            Else
               SBValue(eBar) = lV + lSC
            End If

         End If

      Case SB_PAGELEFT, SB_PAGEUP
         SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)

      Case SB_PAGERIGHT, SB_PAGEDOWN
         SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)

      Case SB_ENDSCROLL
         If Not mbScrollTrack Then
            DrawGrid mbRedraw
            bRedraw = False
         End If
      End Select

      If (lHSB <> SBValue(efsHorizontal)) Or (lVSB <> SBValue(efsVertical)) Then
         Call UpdateCell
         If bRedraw Then
            DrawGrid mbRedraw
         End If
         RaiseEvent Scroll
      End If

   Case WM_MOUSEMOVE
      If Not mbInCtrl Then
         mbInCtrl = True
         Call TrackMouseLeave(hWnd)
         RaiseEvent MouseEnter
      End If

   Case WM_MOUSELEAVE
      If mbInCtrl Then
         mbInCtrl = False
         picTooltip.Visible = False '// MLH
         Call DrawHeaderRow
         UserControl.Refresh
         RaiseEvent MouseLeave
      End If

   Case WM_SETFOCUS
'// MLH
'''      If mbEnabled Then
'''         If Not mbInFocus Then
'''            mbInFocus = True
'''            If mbLockFocusDraw Then
'''               'Debug.Print "WM_SETFOCUS"
'''               DrawGrid mbRedraw
'''            End If
'''         End If
'''      End If
      
      '// mbLockFocusDraw is set in UserControl_Initialize (Language = English)
      If mbHideSelection Then
         If mbInCtrl Or mbLockFocusDraw Then
            'Debug.Print "WM_SETFOCUS: " & Now
            Call DrawGrid(mbRedraw)
         End If
      End If
            
   Case WM_KILLFOCUS
'/ MLH
'''      If hWnd = UserControl.hWnd Then
'''         If mbEnabled Then
'''            If mbInFocus Then
'''               mbInFocus = False
'''               If mbLockFocusDraw Then
'''                  'Debug.Print "WM_KILLFOCUS"
'''                  DrawGrid mbRedraw
'''               End If
'''            End If
'''         End If
'''
'''      ElseIf Not mbInCtrl Then
'''         Call UpdateCell
'''      End If
   
      If Not mbInCtrl Then
         If mbEditPending Then
            Call UpdateCell
         End If
      End If
      '// mbLockFocusDraw is set in UserControl_Initialize (Language = English)
      If mbHideSelection Then
         If Not mbInCtrl Or mbLockFocusDraw Then
            'Debug.Print "WM_KILLFOCUS: " & Now
            Call DrawGrid(mbRedraw, True)
         End If
      End If
      
   Case WM_THEMECHANGED
      RaiseEvent ThemeChanged
      Call SetThemeColor
      DrawGrid mbRedraw

   End Select

End Sub

'''
'// MLH - the following were not used ------------------------------------
'''
'''
'''Public Property Get BackColorFixed() As OLE_COLOR

   '''
   '''   BackColorFixed = mBackColorFixed
   '''

   '''End Property

'''Public Property Let BackColorFixed(ByVal vNewValue As OLE_COLOR)

   '''
   '''   mBackColorFixed = vNewValue
   '''   PropertyChanged "BackColorFixed"
   '''

   '''End Property

'''Public Function ColLeft(ByVal Index As Long) As Long '// MLH - Unused

   '''
   '''  Dim R As RECT
   '''
   '''   SetColRect Index, R
   '''   ColLeft = R.Left
   '''

   '''End Function

'''Public Property Get ColPosition(ByVal Index As Long) As Long

   '''
   '''   ColPosition = mColPtr(Index)
   '''

   '''End Property

'''Public Property Let ColPosition(ByVal Index As Long, ByVal vNewValue As Long)

   '''
   '''  Dim lTemp As Long
   '''
   '''   If Not (mColPtr(Index) = vNewValue) Then
   '''      lTemp = mColPtr(Index)
   '''      mColPtr(Index) = vNewValue
   '''      mColPtr(vNewValue) = lTemp
   '''
   '''      DrawGrid mbRedraw
   '''   End If
   '''

   '''End Property

'''Public Property Let Cols(ByVal vNewValue As Long) '// MLH - use sub AddColumn

   '''
   '''   ReDim mCols(vNewValue) As udtColumn
   '''

   '''End Property

'''Public Sub DebugFormatTable()

   '''
   '''
   '''    Dim lCount As Long
   '''    Dim lTotalRef As Long
   '''
   '''    For lCount = LBound(mCF) To UBound(mCF)
   '''        With mCF(lCount)
   '''            If .lRefCount > 0 Then
   '''                Debug.Print lCount; .lRefCount; .lBackColor; .lForeColor; .sFontName
   '''                lTotalRef = lTotalRef + .lRefCount
   '''            End If
   '''        End With
   '''    Next lCount
   '''
   '''    Debug.Print " = " & lTotalRef
   '''
   '''

   '''End Sub

'''Public Property Get ForeColorFixed() As OLE_COLOR

   '''
   '''   ForeColorFixed = mForeColorFixed
   '''

   '''End Property

'''Public Property Let ForeColorFixed(ByVal vNewValue As OLE_COLOR)

   '''
   '''   mForeColorFixed = vNewValue
   '''   PropertyChanged "ForeColorFixed"
   '''

   '''End Property

'''Public Property Get ForeColorTotals() As OLE_COLOR '// MLH - Not Used

   '''
   '''   ForeColorTotals = mForeColorTotals
   '''

   '''End Property

'''Public Property Let ForeColorTotals(ByVal vNewValue As OLE_COLOR) '// MLH - Not Used

   '''
   '''   mForeColorTotals = vNewValue
   '''   DisplayChange
   '''
   '''   PropertyChanged "ForeColorTotals"
   '''

   '''End Property

'''Public Property Get FormatString() As String

   '''
   '''   FormatString = mFormatString
   '''

   '''End Property

'''Public Property Let FormatString(ByVal vNewValue As String)

   '''
   '''   ''---------------------------------------------------------------------------------
   '''   '// Purpose: Used to create multiple Columns with one string
   '''
   '''   'Each Column is seperated by a "|" char. The Alignment can be specified by
   '''   'using "^" for Centre, "<" for right an ">" for left (default)
   '''   ''---------------------------------------------------------------------------------
   '''
   '''  Dim lcol    As Long
   '''  Dim sCols() As String
   '''
   '''   mFormatString = vNewValue
   '''
   '''   If LenB(mFormatString) Then
   '''      sCols() = Split(vNewValue, "|")
   '''
   '''      If UBound(sCols()) > UBound(mCols) Then
   '''         '// MLH
   '''         '''Cols = UBound(sCols()) + 1
   '''         ReDim mCols(UBound(sCols()) + 1) As udtColumn
   '''      End If
   '''
   '''      For lcol = LBound(sCols) To UBound(sCols)
   '''
   '''         Select Case Mid$(sCols(lcol), 1, 1)
   '''         Case "^"
   '''            mCols(mColPtr(lcol)).sCaption = Mid$(sCols(lcol), 2)
   '''            mCols(mColPtr(lcol)).nAlignment = lgAlignCenterCenter
   '''
   '''         Case "<"
   '''            mCols(mColPtr(lcol)).sCaption = Mid$(sCols(lcol), 2)
   '''            mCols(mColPtr(lcol)).nAlignment = lgAlignLeftCenter
   '''
   '''         Case ">"
   '''            mCols(mColPtr(lcol)).sCaption = Mid$(sCols(lcol), 2)
   '''            mCols(mColPtr(lcol)).nAlignment = lgAlignRightCenter
   '''
   '''         Case Else
   '''            mCols(mColPtr(lcol)).sCaption = sCols(lcol)
   '''         End Select
   '''
   '''         mCols(mColPtr(lcol)).dCustomWidth = 1000
   '''         mCols(mColPtr(lcol)).lWidth = ScaleX(mCols(mColPtr(lcol)).dCustomWidth, muScaleUnits, vbPixels)
   '''         mCols(mColPtr(lcol)).bVisible = True
   '''      Next lcol
   '''
   '''   Else
   '''      ReDim mCols(0) As udtColumn
   '''      Clear
   '''   End If
   '''
   '''   DisplayChange
   '''
   '''   PropertyChanged "FormatString"
   '''

   '''End Property

'''Private Function GetRowHeight() As Long

   '''
   '''   '// Purpose: Return Row Height
   '''   '// MLH - no longer needed - changed mMinRowHeight's scale to vbPixels
   '''
   '''   With UserControl
   '''
   '''      'If mMinRowHeight > 0 Then '// MLH - Is never <= 0
   '''         GetRowHeight = .ScaleY(mMinRowHeight, muScaleUnits, vbPixels)
   '''      'Else
   '''      '   GetRowHeight = DEF_ROWHEIGHT
   '''      'End If
   '''
   '''   End With
   '''

   '''End Function

'''Private Sub pSBSetOrientation() '// Not Used

   '''
   '''   ShowScrollBar mSBhWnd, SB_HORZ, Abs((muSBOrienation = Scroll_Both) Or (muSBOrienation = Scroll_Horizontal))
   '''   ShowScrollBar mSBhWnd, SB_VERT, Abs((muSBOrienation = Scroll_Both) Or (muSBOrienation = Scroll_Vertical))
   '''

   '''End Sub

'''Private Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean

   '''
   '''   If (eBar = efsHorizontal) Then
   '''      SBEnabled = m_bEnabledHorz
   '''   Else
   '''      SBEnabled = m_bEnabledVert
   '''   End If
   '''

   '''End Property

'''Private Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)

   '''
   '''  Dim Lo As Long
   '''  Dim lF As Long
   '''
   '''   Lo = eBar
   '''
   '''   If (bEnabled) Then
   '''      lF = ESB_ENABLE_BOTH
   '''   Else
   '''      lF = ESB_DISABLE_BOTH
   '''   End If
   '''
   '''   If (mbSBNoFlatScrollBars) Then
   '''      EnableScrollBar mSBhWnd, Lo, lF
   '''   Else
   '''      FlatSB_EnableScrollBar mSBhWnd, Lo, lF
   '''   End If
   '''

   '''End Property

'''Private Sub SBRefresh()

   '''
   '''   EnableScrollBar mSBhWnd, SB_VERT, ESB_ENABLE_BOTH
   '''

   '''End Sub

'''Private Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long

   '''// MLH - Changed mSBSmallChangeHorz to a Const
   '''
   '''   If (eBar = efsHorizontal) Then
   '''      SBSmallChange = C_SBSmallChangeHorz
   '''   Else
   '''      SBSmallChange = C_SBSmallChangeVert
   '''   End If
   '''

   '''End Property

'''Private Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)

   '''// MLH - Changed mSBSmallChangeHorz to a Const
   '''
   '''   If (eBar = efsHorizontal) Then
   '''      mSBSmallChangeHorz = lSmallChange
   '''   Else
   '''      mSBSmallChangeVert = lSmallChange
   '''   End If
   '''

   '''End Property

'''Public Property Let TopRow(ByVal vNewValue As Long)

   '''
   '''   If vNewValue > SBMax(efsVertical) Then
   '''      SBValue(efsVertical) = SBMax(efsVertical)
   '''   Else
   '''      SBValue(efsVertical) = vNewValue
   '''   End If
   '''
   '''   SetRowCol vNewValue, mCol, True
   '''   Call DrawGrid(mbRedraw)
   '''

   '''End Sub

