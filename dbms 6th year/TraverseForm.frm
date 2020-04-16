VERSION 5.00
Begin VB.Form TraverseForm 
   Caption         =   "Traverse Records"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form5"
   ScaleHeight     =   7320
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back To Main Menu"
      Height          =   615
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO TO FIRST"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO TO LAST"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<BACK"
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT>>"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtEMP_ID 
      DataField       =   "EMP_ID"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_FIRST_NAME 
      DataField       =   "EMP_FIRST_NAME"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_LAST_NAME 
      DataField       =   "EMP_LAST_NAME"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_INSTI 
      DataField       =   "EMP_INSTI"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_DEPT 
      DataField       =   "EMP_DEPT"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_BASIC_PAY 
      DataField       =   "EMP_BASIC_PAY"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_HA 
      DataField       =   "EMP_HA"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_TA 
      DataField       =   "EMP_TA"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_OTHER_ALLOW 
      DataField       =   "EMP_OTHER_ALLOW"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_REBATES 
      DataField       =   "EMP_REBATES"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_DA 
      DataField       =   "EMP_DA"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_GROSS 
      DataField       =   "EMP_GROSS"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtEMP_INCOME_TAX 
      DataField       =   "EMP_INCOME_TAX"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL RECORDS"
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
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
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
      Left            =   2520
      TabIndex        =   33
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "RECORD  NO."
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
      Left            =   240
      TabIndex        =   32
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_ID:"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   31
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_FIRST_NAME:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_LAST_NAME:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_INSTI:"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   28
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_DEPT:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   27
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_BASIC_PAY:"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   26
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_HA:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_TA:"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   24
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_OTHER_ALLOW:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   23
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_REBATES:"
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   22
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_DA:"
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   21
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_GROSS:"
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   20
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMP_INCOME_TAX:"
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   16
      Top             =   5280
      Width           =   1815
   End
End
Attribute VB_Name = "TraverseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataEnvironment1.rsCommand1.MoveFirst
Label2.Caption = DataEnvironment1.rsCommand1.AbsolutePosition
End Sub

Private Sub Command2_Click()
DataEnvironment1.rsCommand1.MoveLast
Label2.Caption = DataEnvironment1.rsCommand1.AbsolutePosition
End Sub

Private Sub Command3_Click()
DataEnvironment1.rsCommand1.MovePrevious
If DataEnvironment1.rsCommand1.BOF Then
DataEnvironment1.rsCommand1.MoveFirst
End If
Label2.Caption = DataEnvironment1.rsCommand1.AbsolutePosition
End Sub

Private Sub Command4_Click()
DataEnvironment1.rsCommand1.MoveNext
If DataEnvironment1.rsCommand1.EOF Then
DataEnvironment1.rsCommand1.MoveLast
End If
Label2.Caption = DataEnvironment1.rsCommand1.AbsolutePosition
End Sub

Private Sub Command5_Click()
Unload Me
MainForm.Show
End Sub

'Private Sub Command6_Click()
'DataEnvironment1.rsCommand1.Update (emp_id = 7865)
'End Sub

Private Sub Form_Load()
Label2.Caption = DataEnvironment1.rsCommand1.AbsolutePosition
Label4.Caption = DataEnvironment1.rsCommand1.RecordCount
End Sub
