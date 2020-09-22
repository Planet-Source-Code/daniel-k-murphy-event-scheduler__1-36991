Attribute VB_Name = "modMath"
Option Explicit

' ----------------------------------------
' Math Module
' ----------------------------------------

' ----------------------------------------
' Conversions between Pixels and Twips
' ----------------------------------------
Public Function CPixelX(Number As Long) As Long
    ' ----------------------------------------
    ' Converts from Twip X to Pixel X
    ' ----------------------------------------
    CPixelX = Number \ Screen.TwipsPerPixelX
End Function

Public Function CPixelY(Number As Long) As Long
    ' ----------------------------------------
    ' Converts from Twip Y to Pixel Y
    ' ----------------------------------------
    CPixelY = Number \ Screen.TwipsPerPixelY
End Function

Public Function CTwipX(Number As Long) As Long
    ' ----------------------------------------
    ' Converts from Pixel X to Twip X
    ' ----------------------------------------
    CTwipX = Screen.TwipsPerPixelX * Number
End Function

Public Function CTwipY(Number As Long) As Long
    ' ----------------------------------------
    ' Converts from Pixel Y to Twip Y
    ' ----------------------------------------
    CTwipY = Screen.TwipsPerPixelY * Number
End Function

