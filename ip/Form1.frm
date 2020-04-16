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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Command1.Font.Bold = True
Command1.Font.Size = 9
Me.Font.Italic = True
Me.Font.Size = 10
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static count As Integer
Cls
count = count + 1
CurrentX = 0
CurrentY = 0
Print X, Y, count
CurrentX = 100
CurrentY = 200
Str1$ = 5655 & 65
Print Str1$
Print CurrentX, CurrentY
Print CurrentX, CurrentY
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form2
Form2.Show
End Sub

