VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3030
   ClientTop       =   3645
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
vipul
End Sub

Private Sub Form_Activate()
MsgBox ("form1 activate")
End Sub

Private Sub Form_Click()
Form2.Command1_Click
End Sub

Private Sub Form_Deactivate()
MsgBox ("form1 deactivate")
End Sub

Private Sub Form_GotFocus()
MsgBox ("form1 gotfocus")
End Sub

Private Sub Form_Initialize()
MsgBox ("form1 initialise")
End Sub

Private Sub Form_Load()
Print "Press command button to call global sub vipul in module"
Print "click on form to call command_click sub in form2"
MsgBox ("form1 loaded")
Load Form2
Form2.Show
Show 'vbModal
End Sub

Private Sub Form_LostFocus()
MsgBox ("form1 lostfocus")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox ("form1 queryunload")
End Sub

Private Sub Form_Paint()
MsgBox ("form1 paint")
End Sub

Private Sub Form_Resize()
MsgBox ("form1 resize")
End Sub

Private Sub Form_Terminate()
MsgBox ("form1 terminate")
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox ("form1 unload")
End Sub
