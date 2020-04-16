VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim x$
Do
x$ = InputBox$("Password please?")
Loop Until x$ = "vipul"
Show
Print "Hey! u finally got it"
End Sub
