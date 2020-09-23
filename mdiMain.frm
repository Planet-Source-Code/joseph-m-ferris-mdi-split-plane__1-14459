VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Split Plane Test"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timResize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6420
      Top             =   135
   End
   Begin VB.PictureBox picPlane 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8310
      Left            =   7035
      ScaleHeight     =   554
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin VB.PictureBox picInternalWindow 
         BackColor       =   &H80000005&
         Height          =   8280
         Left            =   120
         ScaleHeight     =   8220
         ScaleWidth      =   2820
         TabIndex        =   2
         Top             =   30
         Width           =   2880
      End
      Begin VB.PictureBox picHandle 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   8310
         Left            =   15
         MousePointer    =   9  'Size W E
         ScaleHeight     =   8310
         ScaleWidth      =   120
         TabIndex        =   1
         Top             =   15
         Width           =   120
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

'*' Load the other form into the MDI Parent.  When the Child form is Maximized, it will be
'*' resized by the Parent.  If it is in the "Normal" Window State, the Desktop of the Parent
'*' form will be resized (with scrollbars appearing as needed).
'*'
frmTest.Show

End Sub

Private Sub picHandle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' The MouseDown event in the picHandle object is the trigger that will allow the tracking
'*' of the mouse position to begin.  All of the calculations are done in a timer to allow
'*' for the repetative and constant tracking of the mouse position and object sizes.
'*'
timResize.Enabled = True

End Sub

Private Sub picHandle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' The MouseUp event kills the trigger that was set on the MouseDown event.  This will mean
'*' that the user has release the mouse button and does not wish to resize the split plane
'*' any further.
'*'
timResize.Enabled = False

End Sub

Private Sub picPlane_Resize()

On Error Resume Next

'*' Resize the picturebox inside of the Split Plane
'*'
picInternalWindow.Width = picPlane.Width - 75

End Sub

Private Sub timResize_Timer()

Dim MinWidth As Long            '*' Minimum Width of the Split Plane
Dim MaxWidth As Long            '*' Maximum Width of the Split Plane
Dim lngRelX As Long             '*' Calculated value of the width of the Split Plane
Dim CurrentX As Long            '*' Current XValue of the Mouse

Static LastMousePosX As Long    '*' Static Variable to Track the Mouse's Last Known Position

'*' Get the current value of the X based upon the location of the mouse.
'*'
CurrentX = GetX

'*' If the value has not changed since the last time, the user has not moved position.
'*' Exit the subroutine to skip redundant calculations.
'*'
If CurrentX = LastMousePosX Then
    Exit Sub
Else
    '*' If the value is different, set the value to the current value of X for the next time
    '*' this event fires.
    '*'
    LastMousePosX = CurrentX
End If

'*' The minimum and maximum width of the splitter plane can be set to either an absolute value
'*' or to an equation.  Values are represented in Twips.
'*'
MinWidth = 150                  '*' On some machines, smaller values cause jumpiness.
MaxWidth = mdiMain.Width / 2   '*' Limit the maximum to be one half of the form's size.

'*' Equation for Determining the width of the PictureBox (Solve for i)
'*'
'*' Mx = Left of the MDI Form
'*' Px = Left of the Mouse Pointer
'*' S = Scale (Width/ScaleWidth)
'*' Mw = Width of the MDI Form
'*'
'*' i = Mx - (Px * S) + Mw
'*'
'*' Yields: Anticipated width of the Split Plane
'*'
intCalc = mdiMain.Left - (CurrentX * (picPlane.Width / picPlane.ScaleWidth)) + mdiMain.Width

'*' Bounds Checking.  Make sure that the value returned is within bounds.  If it is not, set
'*' it to the proper value and exit the sub.
'*'
If intCalc <= MinWidth Then
  intCalc = MinWidth
ElseIf intCalc >= MaxWidth Then
  intCalc = MaxWidth
End If

'*' Set the width of the Split Plane to be equal to that
picPlane.Width = intCalc

End Sub
