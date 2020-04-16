VERSION 5.00
Begin VB.Form SearchForm 
   Caption         =   "Employee Search"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form4"
   ScaleHeight     =   4215
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Search By Employee Name"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search By Employee ID"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub Command1_Click()
Unload DataEnvironment2
Load DataEnvironment2
Dim FindEmp
FindEmp = InputBox("Enter the EMPLOYEE  ID. of the Employee", "Enter EMPLOYEE ID.")
If (FindEmp = "") Then
Exit Sub
End If

DataEnvironment2.Commands("Command1").CommandText = "select * from incometax where emp_id = " & FindEmp
DataEnvironment2.Commands("Command1").CommandType = adCmdText
DataEnvironment2.Command1
If (DataEnvironment2.rsCommand1.RecordCount = 0) Then
    MsgBox "No Results"
    Exit Sub
Else
    MsgBox "No Of Entries Found Was::" & DataEnvironment2.rsCommand1.RecordCount, , "*******Search Count******"
    Unload Me
    QueryForm.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
MainForm.Show
End Sub

Private Sub Command3_Click()
Unload DataEnvironment2
Load DataEnvironment2
Dim FindEmp
FindEmp = InputBox("Enter the Name of the Employee", "Enter EMPLOYEE Name")
If (FindEmp = "") Then
Exit Sub
End If

DataEnvironment2.Commands("Command1").CommandText = "select * from incometax where emp_first_name like '%" & FindEmp & "%'"
DataEnvironment2.Commands("Command1").CommandType = adCmdText
DataEnvironment2.Command1
If (DataEnvironment2.rsCommand1.RecordCount = 0) Then
    MsgBox "No Results"
    Exit Sub
Else
    MsgBox "No Of Entries Found Was::" & DataEnvironment2.rsCommand1.RecordCount, , "*******Search Count******"
    Unload Me
    QueryForm.Show
End If
End Sub
