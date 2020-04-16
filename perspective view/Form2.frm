VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   ScaleHeight     =   4935
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Reset Figure"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      X1              =   4320
      X2              =   4320
      Y1              =   3840
      Y2              =   4320
   End
   Begin VB.Label Label8 
      Caption         =   "x"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "y"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "z"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   4320
      Width           =   135
   End
   Begin VB.Line Line3 
      X1              =   4320
      X2              =   4080
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Label Label5 
      Caption         =   "Z :"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Y :"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "X :"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Transformation"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Scaling Factor"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Form1.Combo1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Changes
End Sub
