VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move The Red Dot!"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Line Line4 
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Planet-Source-Code.com Submission:
'
' How to move objects!
' Joe Burman, 2003
' E-MAIL: webmaster@jburman.com




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' This is to define what the keys do WHEN they are pressed. For movement cases
' it first checks to see if the dot is at its boundary.

Select Case KeyCode ' Starts the KeyCode Cases

Case (vbKeyLeft) 'This is Left key
If Shape1.Left >= 220 Then 'This checks to make sure the dot isn't at its boundary on the left side.
Shape1.Left = Shape1.Left - 100 'This allows the dot to move 100 twips to the left.
End If

Case (vbKeyUp) 'This is Up key
If Shape1.Top >= 220 Then 'This checks to make sure the dot isn't at its boundary on the top.
Shape1.Top = Shape1.Top - 100 'This allows the dot to move 100 twips upward.
End If

Case (vbKeyDown) 'This is Down key
If Shape1.Top <= 1820 Then 'This checks to make sure the dot isn't at its boundary on the bottom.
Shape1.Top = Shape1.Top + 100 'This allows the dot to move 100 twips downward.
End If

Case (vbKeyRight) 'This is Right key
If Shape1.Left <= 2180 Then 'This checks to make sure the dot isn't at its boundary on the right side.
Shape1.Left = Shape1.Left + 100 'This allows the dot to move 100 twips to the right.
End If

End Select
End Sub
