VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   2895
   ClientTop       =   1635
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim vipul As Date

Private Sub Form_Activate()
MsgBox ("form1 active")
End Sub

Private Sub Form_GotFocus()
'Me.Print "the form has got focus" '& vbCrLf
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MsgBox ("(" & CurrentX & "," & CurrentY & ")" & TextWidth("VIpul"))
End Sub

Private Sub Form_Load()
Show
Load Form2
Form2.Show
'Print "now the form has been shown"

End Sub

Private Sub Form_Click()
'Me.Refresh
'MsgBox (Now)
Static vipul As Integer
CurrentX = vipul + 3
CurrentY = vipul + 3
vipul = vipul + 20
Print Now;
Print Spc(30);
Print "VIpul ";
Print "dsgfSDF";
'Dim v As Integer
'v = MsgBox("hjvhhjk", vbYesNo)
End Sub

Private Sub Form_Paint()
Cls
Dim vipul As String
vipul = "Welcome by Vipul Mistry"
CurrentX = ScaleWidth / 2 - TextWidth(vipul) / 2
CurrentY = ScaleHeight / 2 - TextHeight(vipul) / 2
Print vipul
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancel = True
End Sub

Private Sub Form_Resize()
Form_Paint
End Sub
