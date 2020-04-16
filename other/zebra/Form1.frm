VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f1 As New Form1
Dim f2 As New Form1

Private Sub Command1_Click()
MsgBox ("sfgsdgsdf")
End Sub

Private Sub Form_Click()
Dim tex1 As Form1
Set tex1 = Text1
tex1.Text = "gdfgs"
Set tex1 = f1.Text1
tex1.Text = "dsfvsfd"
f1.Show
f2.Show
f1.Move Left - 199, Top + 100
f2.Move Left - 399, Top + 200
End Sub
