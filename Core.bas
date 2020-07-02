Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API functions used by this program:
Public Declare Function DestroyIcon Lib "User32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DrawIcon Lib "User32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function ExtractIconA Lib "Shell32.dll" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

'The constants used by this program:
Public Const GET_ICON_COUNT As Long = -1  'Defines the "Get the icon count." message.
Public Const NO_HANDLE As Long = -1       'Defines a null handle.

'This procedure returns information about this program.
Public Function ProgramInformation() As String
   With App
      ProgramInformation = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With
End Function


'This procedure is executed when this program is started.
Public Sub Main()
On Error Resume Next
   InterfaceWindow.Show
End Sub


