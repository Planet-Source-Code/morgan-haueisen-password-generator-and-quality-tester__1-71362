VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlaySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//***********************************
'// Author: Morgan Haueisen
'// Copyright (c) 1997-2004
'//***********************************
Option Explicit

'// flag values for dwFlags parameter
Private Const SND_SYNC                   As Long = &H0        '// play synchronously (default)
Private Const SND_ASYNC                  As Long = &H1        '// play asynchronously
Private Const SND_NODEFAULT              As Long = &H2        '// silence not default, if sound not found
Private Const SND_MEMORY                 As Long = &H4        '// lpszSoundName points to a memory file
Private Const SND_LOOP                   As Long = &H8        '// loop the sound until next sndPlaySound
Private Const SND_NOWAIT                 As Long = &H2000     '// don't wait if the driver is busy

'Private Const SND_ALIAS                  As Long = &H10000    '// name is a WIN.INI [sounds] entry
'Private Const SND_FILENAME               As Long = &H20000    '// name is a file name
'Private Const SND_RESOURCE               As Long = &H40004    '// name is a resource name or atom
'Private Const SND_ALIAS_ID               As Long = &H110000   '// name is a WIN.INI [sounds] entry identifier
'Private Const SND_ALIAS_START            As Long = 0          '// must be > 4096 to keep strings in same section of resource file
'Private Const SND_NOSTOP                 As Long = &H10       '// don't stop any currently playing sound
'Private Const SND_VALID                  As Long = &H1F       '// valid flags
'Private Const SND_VALIDFLAGS             As Long = &H17201F   '// Set of valid flag bits.
'                                                              '// Anything outside this range will raise an error
'Private Const SND_RESERVED               As Long = &HFF000000 '// In particular these flags are reserved
'Private Const SND_TYPE_MASK              As Long = &H170007   '// waveform audio error return values

'// Error values
'Private Const WAVERR_BASE                As Long = 32
'Private Const WAVERR_BADFORMAT           As Long = (WAVERR_BASE + 0) '// unsupported wave format
'Private Const WAVERR_STILLPLAYING        As Long = (WAVERR_BASE + 1) '// still something playing
'Private Const WAVERR_UNPREPARED          As Long = (WAVERR_BASE + 2) '// header not prepared
'Private Const WAVERR_SYNC                As Long = (WAVERR_BASE + 3) '// device is synchronous
'Private Const WAVERR_LASTERROR           As Long = (WAVERR_BASE + 3) '// last error in range

Public Enum enuSystemSound
   SS_OK = &H0&
   SS_IconAsterisk = &H10&
   SS_IconQuestion = &H20&
   SS_IconExclamation = &H30&
   SS_IconInformation = &H40&
End Enum

Private Declare Function PlaySound Lib "winmm.dll" _
      Alias "PlaySoundA" ( _
      ByVal lpszName As String, _
      ByVal hModule As Long, _
      ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" _
      Alias "PlaySoundA" ( _
      ByRef lpData As Any, _
      ByVal hModule As Long, _
      ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function MessageBeep Lib "user32" ( _
      ByVal wType As Long) As Long
      
Private m_Sound()            As Byte
Private m_dwFlags            As Long
Private mstrLastSoundPlayed  As String

Public Function DevicesPresent() As Boolean
   
   On Error Resume Next
   '// Can the system play sound files?
   If waveOutGetNumDevs() = 0 Then
      DevicesPresent = False
   Else
      DevicesPresent = True
   End If
   
End Function

Public Sub PlaySoundFile(ByVal PathWAVname As String, _
                         Optional ByVal vAsync As Boolean = True, _
                         Optional ByVal vLoop As Boolean = False)
   
   
   On Local Error Resume Next
   If DevicesPresent Then
      m_dwFlags = SND_SYNC
      If vAsync Then m_dwFlags = SND_ASYNC
      If vLoop Then m_dwFlags = m_dwFlags Or SND_LOOP
      
      If LenB(Dir$(PathWAVname)) Then
         If PlaySound(CStr(PathWAVname), App.hInstance, m_dwFlags) = 0 Then
            Beep
         End If
      Else
         Beep
      End If
   Else
      Beep
   End If
   On Local Error GoTo 0
   
End Sub

Public Function PlaySoundResource(ByVal SndID As Long, _
                                  Optional ByVal SndType As String = "SOUND", _
                                  Optional ByVal vAsync As Boolean = False, _
                                  Optional ByVal vLoop As Boolean = False) As Long
   
   On Local Error Resume Next
   If DevicesPresent Then
      m_dwFlags = SND_MEMORY Or SND_NODEFAULT Or SND_NOWAIT
      If vAsync Then m_dwFlags = m_dwFlags Or SND_ASYNC
      If vLoop Then m_dwFlags = m_dwFlags Or SND_LOOP
      
      If StrComp(CStr(SndID) & SndType, mstrLastSoundPlayed, vbTextCompare) Then
         m_Sound = LoadResData(SndID, SndType)
      End If
      
      PlaySoundResource = PlaySoundData(m_Sound(0), 0&, m_dwFlags)
      mstrLastSoundPlayed = CStr(SndID) & SndType
   Else
      Beep
   End If
   On Local Error GoTo 0
   
End Function

Public Sub PlaySystemSound(ByVal vSysSound As enuSystemSound)

   Call MessageBeep(vSysSound)

End Sub

