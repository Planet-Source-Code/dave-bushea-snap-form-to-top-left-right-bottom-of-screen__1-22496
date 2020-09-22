Attribute VB_Name = "SnapMod"
'Snaping example by Dave Bushea (http://www.rapta.net)
Public Type POINTAPI
    X As Long
    Y As Long
    End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public mouseX As Integer
Public mouseY As Integer
