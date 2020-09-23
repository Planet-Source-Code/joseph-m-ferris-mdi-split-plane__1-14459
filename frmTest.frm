VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Form"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

'*' Quick and Dirty, print the width of this form as a visual aid to demonstrate the impact
'*' of the Split Plane on a child form.
'*'
Cls
Me.Print "Width = " & Me.Width

End Sub
