VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim c As Integer
c = 0
Do
MsgBox ("Hello I am Vipul")
MsgBox ("Since we have met for the first time let us know about each other")
c = c + 1
Loop While c <> 5
Show
Cls
Print vbCrLf & "  Cool man now u must have known me very well"
Print vbCrLf & "  Press exit to get rid (or fond) of me..."
End Sub
Private Sub Command1_Click()
MsgBox ("So u want to go away but")
Do
MsgBox ("I won't let u go")
MsgBox ("I want to know more about u")
c = c + 1
Loop While c <> 5
End Sub
Private Sub Form_Unload(Cancel As Integer)
MsgBox ("Oh god u are so frustated")
MsgBox ("Now I know much about u")
Do While MsgBox("Hey listen! Have u read all early mess or not", vbYesNo) = vbNo
    MsgBox "u fool stupid bastard, now read all the stuff again"
    Call Form_Load
    Call Command1_Click
Loop
MsgBox ("Now u can go.......")
Do
MsgBox ("Bye....")
c = c + 1
Loop While c <> 5
End Sub
