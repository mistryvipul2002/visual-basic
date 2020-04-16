VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Tab"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Enter"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Left"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Right"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UP"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
common
SendKeys "{up}"
End Sub

Private Sub Command2_Click()
common
SendKeys "{down}"
End Sub

Private Sub Command3_Click()
common
SendKeys "{right}"
End Sub

Private Sub Command4_Click()
common
SendKeys "{left}"
End Sub

Private Sub Command5_Click()
common
SendKeys "{enter}"
End Sub

Private Sub common()
AppActivate "my", True
End Sub

Private Sub Command6_Click()
common
SendKeys "{tab}"
End Sub

Private Sub Form_Load()
Call Shell("E:\Winnt\explorer.exe", vbNormalFocus)
End Sub
