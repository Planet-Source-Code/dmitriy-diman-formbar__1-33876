Attribute VB_Name = "modDeclares"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const BM_SETSTATE = &HF3

Public Const BUTTON_PRESSED = 1
Public Const BUTTON_UNPRESSED = 0
