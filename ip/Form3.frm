VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   1440
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Scale (-320, 240)-(320, -240)
For i = -320 To 320
    PSet (i, 0)
Next
For j = -240 To 240
    PSet (0, j)
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static count As Integer
Me.AutoRedraw = False
count = count + 1
CurrentX = 0
CurrentY = 0
Text1.Text = X
Text2.Text = Y
Text3.Text = count
End Sub
