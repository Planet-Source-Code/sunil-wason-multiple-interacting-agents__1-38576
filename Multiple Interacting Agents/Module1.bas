Attribute VB_Name = "Module1"
Option Explicit

' public declarations required by SetWindowPos function
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40

' function to set window position on top
' to enable "always on top"
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal _
hWndInsertAfter As Long, ByVal x As Long, ByVal y As _
Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
As Long) As Long

