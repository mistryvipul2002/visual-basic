VERSION 5.00
Begin VB.Form SQLForm 
   Caption         =   "SQL Query Window"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   735
      Left            =   4560
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit to Oracle"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your SQL Query Here : -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "SQLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload DataEnvironment1
Load DataEnvironment1
DataEnvironment1.Commands("Command3").CommandText = Text1.Text
DataEnvironment1.Commands("Command3").CommandType = adCmdText
DataEnvironment1.Command3
If (DataEnvironment1.rsCommand3.RecordCount = 0) Then
MsgBox "No Results"
Else
MsgBox "No Of Entries Found Was::" & DataEnvironment1.rsCommand3.RecordCount, , "*******Search Count******"
DataGridSQLForm.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
MainForm.Show
End Sub
