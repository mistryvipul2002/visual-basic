VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartX As Integer, StartY As Integer
Dim TempX As Integer, TempY As Integer
Dim CircleRadius As Integer

Private Sub Form_Load()
CircleRadius = 3
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartX = X
TempX = X
StartY = Y
TempY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Line (TempX, TempY)-(StartX, StartY), BackColor
    Line (X, Y)-(StartX, StartY)
    TempX = X
    TempY = Y
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Circle (StartX, StartY), CircleRadius
Circle (TempX, TempY), CircleRadius
End Sub
