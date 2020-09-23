Attribute VB_Name = "mod_XPStyle"
'// Author: Morgan Haueisen (morganh@hartcom.net)
'// Copyright (c) 2005
'// Version 2

'// ** REQUIRES RESOURCE FILE resXP.res **

Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC  As Long
End Type

'// Operating system version information

Private Type OSVersionInfo
   OSVSize       As Long
   dwVerMajor    As Long
   dwVerMinor    As Long
   dwBuildNumber As Long
   PlatformID    As Long
   szCSDVersion  As String * 128
End Type

Private Declare Function CreateMutex Lib "kernel32" _
      Alias "CreateMutexA" ( _
      ByRef lpMutexAttributes As Any, _
      ByVal bInitialOwner As Long, _
      ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVersionInfo) As Long

Private mlngMutex As Long
Public gblnOsIsXp As Boolean

Public Sub EndApp(Optional CallingForm As Form)

   '// Call from the closing Form's Form_Unload event
   '//
   '// Example:
   '//   Call EndApp(Me)
   '//   Set FormName = Nothing
   '//   '// End Program

  Dim Frm As Form
  Const SEM_NOGPFAULTERRORBOX As Long = &H2&

   On Error Resume Next

   '// free memory
   If mlngMutex Then
      Call ReleaseMutex(mlngMutex)
      Call CloseHandle(mlngMutex)
   End If

   '// Close all open Forms
   For Each Frm In Forms

      If Not (Frm.Name = CallingForm.Name) Then
         Unload Frm
         Set Frm = Nothing
      End If

   Next Frm

   '// Some versions of ComCtl32.DLL version 6.0 cause a crash at shutdown
   '// when you enable XP Visual Styles in an application that has a VB User Control.
   '// This instructs Windows to not display the UAE message box that invites you to send
   '// Microsoft information about the problem.
   If Not IsInIDE Then '// Not running in IDE
      Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
   End If

End Sub

Public Sub IsAppRunning()

   ''' Const ERROR_ALREADY_EXISTS = 183&

   If Not IsInIDE Then '// Ignore if running within IDE

      '// Is this application already open?
      '// (If it is open then end program)
      mlngMutex = CreateMutex(ByVal 0&, 1, App.Title)

      If (Err.LastDllError = 183&) Then
         '// free memory
         Call ReleaseMutex(mlngMutex)
         Call CloseHandle(mlngMutex)
         MsgBox App.Title & " is already running.", vbExclamation
         Call EndApp
         End
      End If

   End If

End Sub

Public Function IsInIDE() As Boolean

   '// Return whether we're running in the IDE.

   '// Assert invocations work only within the development environment and
   '// conditionally suspends execution (if set to False) at the line on which
   '// the method appears.
   '// When the module is compiled into an executable, the method calls on the
   '// Debug object are omitted.

   Debug.Assert zSetTrue(IsInIDE)

End Function

Public Sub ManifestWrite(Optional ByVal vblnOnlyIfXP As Boolean = True)

  Dim OSV As OSVersionInfo

   On Error Resume Next

   '// Get OS compatability flag
   OSV.OSVSize = Len(OSV)

   If GetVersionEx(OSV) = 1 Then
      If OSV.PlatformID = 2 Then
         If OSV.dwVerMajor = 5 Then
            If OSV.dwVerMinor = 1 Then
               gblnOsIsXp = True '// OS is XP
            End If

         ElseIf OSV.dwVerMajor > 5 Then
            gblnOsIsXp = True '// OS is Vista
         End If

      End If
   End If

   '// If OS is XP or force always write then continue
   If gblnOsIsXp Or Not vblnOnlyIfXP Then

      '// Link XP themes to application
      Dim iccex                As tagInitCommonControlsEx
      Const ICC_USEREX_CLASSES As Long = &H200

      With iccex
         .lngSize = LenB(iccex)
         .lngICC = ICC_USEREX_CLASSES
      End With

      Call InitCommonControlsEx(iccex)

   End If 'gblnOsIsXp

End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean

   '// Worker function for IsInIDE
   zSetTrue = True
   bValue = True

End Function

