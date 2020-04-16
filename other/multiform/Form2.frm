VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
MsgBox ("command button1 pressed on form2")
End Sub

Private Sub Form_Activate()
MsgBox ("form2 activate")
End Sub

Private Sub Form_Deactivate()
MsgBox ("form2 deactivate")
End Sub

Private Sub Form_GotFocus()
MsgBox ("form2 gotfocus")
End Sub

Private Sub Form_Initialize()
MsgBox ("form2 initialise")
End Sub

Private Sub Form_Load()
MsgBox ("form2 loaded")
End Sub

Private Sub Form_LostFocus()
MsgBox ("form2 lostfocus")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox ("form2 queryunload")
End Sub

Private Sub Form_Terminate()
MsgBox ("form2 terminate")
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox ("form2 unload")
End Sub

Private Sub Form_Paint()
MsgBox ("form2 paint")
End Sub

Private Sub Form_Resize()
MsgBox ("form2 resize")
End Sub
