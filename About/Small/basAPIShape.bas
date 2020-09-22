Attribute VB_Name = "basAPIShape"
Option Explicit

''----------------------------------------------------------
'' support for mouse_over and mouse_leave events
''----------------------------------------------------------
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub FormMove(F As Form)    ' or any object with an hWnd
  Const WM_NCLBUTTONDOWN = &HA1
  ReleaseCapture
  Call SendMessage(F.hWnd, WM_NCLBUTTONDOWN, 2, 0&)
End Sub


