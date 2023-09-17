VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer IconLoader 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface.
Option Explicit



'This procedure requests the user the select a file containing icons.
Private Sub Form_Activate()
On Error Resume Next
   FileDialog.ShowOpen
   
   If FileDialog.FileName = vbNullString Then
      Unload Me
   Else
      IconLoader.Enabled = True
   End If
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error Resume Next
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   Me.Caption = ProgramInformation()
End Sub


Private Sub IconLoader_Timer()
On Error Resume Next
Dim IconH As Long
Static Count As Long
Static Index As Long

   If Count = 0 Then
      Count = ExtractIconA(App.hInstance, FileDialog.FileName, GET_ICON_COUNT)
      Index = 0
   Else
      Me.Cls
   
      IconH = ExtractIconA(App.hInstance, FileDialog.FileName, Index)
      If Not IconH = NO_HANDLE Then
         DrawIcon Me.hdc, CLng(16), CLng(16), IconH
         DestroyIcon IconH
      End If
   
      Me.Caption = ProgramInformation() & " Icon: " & CStr(Index + 1) & "/" & CStr(Count)
   
      Index = Index + 1
      If Index = Count Then
         Count = 0
         IconLoader.Enabled = False
      End If
   End If
End Sub


