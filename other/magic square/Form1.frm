VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magic Square"
   ClientHeight    =   5190
   ClientLeft      =   2385
   ClientTop       =   2040
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Size As Integer, Temp As Variant

Private Sub Form_Load()
Size = 2
Do
Temp = InputBox("Enter the size of Magic Square?", "Magic Square")
If IsNumeric(Temp) Then
    Size = CInt(Temp)
    If Size Mod 2 = 0 Then
        MsgBox ("Please enter an odd number!")
    Else: Exit Do
    End If
End If
Loop
ArrangeLabels (Size) 'call function to arrange label positions
Show
Print "The Magic Square of size : - " & Size
MagicSquareDraw (Size)
End Sub

Private Sub MagicSquareDraw(Size As Integer)
Dim Counter As Integer, X As Integer, Y As Integer
Dim NewX As Integer, NewY As Integer
X = Size \ 2 + 1
Y = 1 'set (X,Y) to be middle position on top row of matrix
Label(X + (Y - 1) * Size).Caption = 1 ' set value of matrix as 1

For Counter = 2 To Size * Size 'for remaining numbers
    If X = 1 Then ' move X to left
        NewX = Size
    Else: NewX = X - 1
    End If
    If Y = 1 Then ' move Y to up
        NewY = Size
    Else: NewY = Y - 1
    End If
    
    'if there is a value already present at (NewX,NewY) move down
    If Label(NewX + (NewY - 1) * Size).Caption = 0 Then
        Label(NewX + (NewY - 1) * Size).Caption = Counter
    Else ' move down
        If Y = Size Then
            Y = 1
            NewY = Y
        Else
            Y = Y + 1
            NewY = Y
        End If
        NewX = X
        Label(NewX + (NewY - 1) * Size).Caption = Counter
    End If
    
    X = NewX 'update values of X and Y
    Y = NewY
Next
End Sub

Private Sub ArrangeLabels(Size As Integer)
Dim CellHeight As Integer, CellWidth As Integer
CellHeight = TextHeight("888")
CellWidth = TextWidth("888")
Dim VertSpace As Integer, HoriSpace As Integer
VertSpace = 50
HoriSpace = 50
CellWidth = CellWidth + HoriSpace
CellHeight = CellHeight + VertSpace
Me.Width = (Size + 2) * CellWidth
Me.Height = (Size + 3) * CellHeight
Dim Counter As Integer
For Counter = 1 To Size * Size
    Load Label(Counter)
    If (Counter Mod Size) = 0 Then
        Label(Counter).Left = Size * CellWidth
        Label(Counter).Top = (Int(Counter / Size)) * CellHeight  ' Int function fives floor of the number in argument
    Else:
        Label(Counter).Left = (Counter Mod Size) * CellWidth
        Label(Counter).Top = (Int(Counter / Size) + 1) * CellHeight
    End If
    Label(Counter).Width = CellWidth
    Label(Counter).Height = CellHeight
    Label(Counter).Caption = 0
    Label(Counter).Visible = True
Next
End Sub
