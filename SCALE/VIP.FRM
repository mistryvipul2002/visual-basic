VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   2730
   ClientTop       =   2640
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5580
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "scale"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "print"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "scale"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3720
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'Dim Str As String
'Str = InputBox("this is pyschotic", "hello", "default string", 400, 400)
'End Sub
Private Sub Command1_Click()
Me.ScaleMode = Me.ScaleMode + 1
Text1 = Me.ScaleMode
End Sub
Private Sub Command2_Click()
Cls
Me.CurrentX = 40
Me.CurrentY = 10
Print "thtytrtgfd";
End Sub
Private Sub Form_Load()
Show
Text1 = Me.ScaleMode
End Sub
Public Function rand(x) As Integer
    Randomize
    rand = Int(x * Rnd) + 1
End Function
