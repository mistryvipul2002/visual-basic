VERSION 5.00
Begin VB.Form DeleteForm 
   Caption         =   "Delete Record"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Main Menu"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Employee ID"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "DeleteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub Command1_Click()
On Error GoTo e
Unload DataEnvironment1
Load DataEnvironment1
str = str & " delete from incometax where emp_id= " & Text1.Text
MsgBox str
DataEnvironment1.Commands("command3").CommandText = str
  DataEnvironment1.Commands("command3").CommandType = adCmdText
  DataEnvironment1.Command3
  str = ""
e: If Not (Err.Description = "") Then MsgBox Err.Description, , "error?"
str = ""
End Sub

Private Sub Command2_Click()
Unload Me
MainForm.Show
End Sub
