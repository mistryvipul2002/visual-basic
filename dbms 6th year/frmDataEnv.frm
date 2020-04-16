VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Main Menu"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "TRAVERSE AND UPDATE DATA"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "FIND DATA"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DELETE DATA "
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "INSERT DATA"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SQL  QUERY"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
Unload Me
TraverseForm.Show
End Sub
Private Sub Command1_Click()
DataEnvironment1.rsCommand1.MoveFirst

End Sub

Private Sub Command2_Click()
DataEnvironment1.rsCommand1.MoveLast
End Sub

Private Sub Command3_Click()

On Error GoTo e:
DataEnvironment1.rsCommand1.MovePrevious
If DataEnvironment1.rsCommand1.BOF Then
DataEnvironment1.rsCommand1.MoveFirst
End If
e: If (Err.Number = 3704) Then Unload frmNavigator
End Sub

Private Sub Command4_Click()
DataEnvironment1.rsCommand1.MoveNext
If DataEnvironment1.rsCommand1.EOF Then
DataEnvironment1.rsCommand1.MoveLast
End If
End Sub

Private Sub Command5_Click()
Unload Me
SQLForm.Show
End Sub

Private Sub Command6_Click()
Unload Me
InsertForm.Show
End Sub

Private Sub Command7_Click()
Unload Me
DeleteForm.Show
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Command8_Click()
Unload Me
SearchForm.Show
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

