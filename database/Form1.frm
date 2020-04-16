VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1335
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ".\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from publishers order by name"
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Publisher"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveNext
End Sub
