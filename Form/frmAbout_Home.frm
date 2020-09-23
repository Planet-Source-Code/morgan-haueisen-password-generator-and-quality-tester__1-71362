VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1365
   ClientLeft      =   2895
   ClientTop       =   3015
   ClientWidth     =   3525
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout_Home.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompanyName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   100
      TabIndex        =   2
      Top             =   390
      UseMnemonic     =   0   'False
      Width           =   3000
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   100
      TabIndex        =   1
      Top             =   645
      UseMnemonic     =   0   'False
      Width           =   3000
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   100
      TabIndex        =   0
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

   On Error Resume Next
   lblTitle.Caption = App.ProductName
   lblCompanyName.Caption = "MorganWare™" 'App.CompanyName

   lblVersion.Caption = "By: Morgan Haueisen" & vbCrLf & "Version " & App.Major & "." & App.Minor & "." & _
      App.Revision & vbCrLf & App.LegalCopyright

   Me.Show
   DoEvents
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAbout = Nothing
End Sub

Private Sub lblDisclaimer_Click()
   Unload Me
End Sub

Private Sub lblCompanyName_Click()
   Unload Me
End Sub

Private Sub lblTitle_Click()
   Unload Me
End Sub

Private Sub lblVersion_Click()
   Unload Me
End Sub


