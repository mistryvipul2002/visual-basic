VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":000D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "Form1.frx":0023
      Left            =   2880
      List            =   "Form1.frx":0025
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Publisher ID"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Print "You selected " & Combo1.Text
If List1.ListIndex <> -1 Then
    Print "You selected " & List1.List(List1.ListIndex)
    Print List1.ListIndex
End If
End Sub

Private Sub Form_Load()
List1.AddItem "Vipul Mistry"
List1.AddItem "Pramod Gupta"
List1.AddItem "Sohanveer Goyal"
Combo1.AddItem "Vipul Mistry"
Combo1.AddItem "Pramod"
Combo1.AddItem "Sohan"
Combo1.Text = Combo1.List(0)
End Sub
