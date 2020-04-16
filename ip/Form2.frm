VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Print "jhgjhg"
X$ = Dir$("c:\*.txt")
If X$ = "" Then
Print "no file"
Else
Print "first file is "; X$
End If
Print "1st"; Spc(50); "50th"
Dim number As Integer
number = 5
For i = 1 To 10
Print "5    *   " & i; Tab(15); "="; Tab(20); 5 * i
Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Print Chr$(KeyAscii);
'Beep
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form3
Form3.Show
End Sub
