Attribute VB_Name = "modAPIStuff"
Option Explicit

' ----------------------------------------
' Declare the API Functions
' ----------------------------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' ----------------------------------------
' Constants used
' ----------------------------------------
Public Const LB_SETTABSTOPS = &H192
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXTLEN = &H18A
Public Const SB_HORZ = 0

' ----------------------------------------
' Types
' ----------------------------------------
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
