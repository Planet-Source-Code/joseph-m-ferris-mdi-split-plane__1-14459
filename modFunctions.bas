Attribute VB_Name = "modFunctions"
'*' The GetCursorPos API Call expects a structure to return it's values to.

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'*' The POINTAPI Type contains two properties.  The first is the X Property, which returns
'*' the Horizontal Postion of the mouse as a long value, and the Y Property, which returns
'*' the Vertical Postion of the mouse as a long value.
'*'
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Function GetX() As Long

Dim POS As POINTAPI

'*' Call the GetCursorPos Function with the POINTAPI object that has been created.
'*'
GetCursorPos POS

'*' The value of X is now assigned to the X Property of the POS Object.  Assign this value
'*' to the function and return the value.
'*'
GetX = POS.X

End Function

