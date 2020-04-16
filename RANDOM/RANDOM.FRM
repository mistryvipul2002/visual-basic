VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "10"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random number"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Random number = "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "to"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Range"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize
Text3 = Int((Int(Text2) - Int(Text1) + 1) * Rnd) + Int(Text1)
End Sub
