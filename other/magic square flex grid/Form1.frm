VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Magic Square"
   ClientHeight    =   6390
   ClientLeft      =   2400
   ClientTop       =   2055
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Size As Long, Temp As Variant, I As Long
Dim FlexCellHeight As Long, FlexCellWidth As Long

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
Show
Print "The Magic Square of size : - " & Size

MSFlexGrid1.Cols = Size + 1
MSFlexGrid1.Rows = Size + 1
FlexCellWidth = Len(Str$(Size * Size)) * TextWidth("8")
FlexCellHeight = TextHeight("1")
'MSFlexGrid1.Height = (Size + 2) * FlexCellHeight
'MSFlexGrid1.Width = (Size + 2) * FlexCellWidth
For I = 1 To Size + 1
    MSFlexGrid1.ColWidth(I - 1) = FlexCellWidth
    MSFlexGrid1.RowHeight(I - 1) = FlexCellHeight
Next
MagicSquareDraw (Size)
End Sub

Private Sub MagicSquareDraw(Size As Long)
Dim Counter As Long, X As Long, Y As Long
Dim NewX As Long, NewY As Long
X = Size \ 2 + 1
Y = 1 'set (X,Y) to be middle position on top row of matrix
MSFlexGrid1.Col = X
MSFlexGrid1.Row = Y
MSFlexGrid1.Text = 1 ' set value of matrix as 1

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
    MSFlexGrid1.Col = NewX
    MSFlexGrid1.Row = NewY
    
    If MSFlexGrid1.Text = "" Then
        MSFlexGrid1.Text = Counter
    Else ' move down
        If Y = Size Then
            Y = 1
            NewY = Y
        Else
            Y = Y + 1
            NewY = Y
        End If
        NewX = X
        MSFlexGrid1.Col = NewX
        MSFlexGrid1.Row = NewY
        MSFlexGrid1.Text = Counter
    End If
    
    X = NewX 'update values of X and Y
    Y = NewY
Next
End Sub
