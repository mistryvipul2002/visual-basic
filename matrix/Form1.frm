VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Matrix Operations"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   1815
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0022
      TabIndex        =   9
      Text            =   "Select an operation to perform ..."
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "m       x      n"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Output"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "m       x      n"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "m       x      n"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetMatrix(Mat1() As Single, Str As String)
Dim R As Integer, C As Integer, I As Integer
Dim RowEnd As Boolean, TotalNum As Integer

'find total numbers
I = 1
TotalNum = 0
Str = Str & "$"
Do While I <= Len(Str)
    If Mid(Str, I, 1) Like "[0-9,.,e,E,+,-]" Then
        Do While Mid(Str, I, 1) Like "[0-9,.,e,E,+,-]"
            I = I + 1
        Loop
        TotalNum = TotalNum + 1
    Else
        I = I + 1
    End If
Loop

'find total columns as C
I = 1
RowEnd = False
C = 1
Do While I <= Len(Str)
    If Mid(Str, I, 1) Like "[0-9,.,e,E,+,-]" Then
        Do While Mid(Str, I, 1) Like "[0-9,.,e,E,+,-]"
            I = I + 1
        Loop
    ElseIf Mid(Str, I, 1) Like " " Then
            If RowEnd = False Then C = C + 1
            I = I + 1
    Else
        RowEnd = True
        Exit Do
    End If
Loop
R = TotalNum / C

'get matrix
Dim Num As String
Dim X As Integer, Y As Integer
ReDim Mat1(1 To R, 1 To C)
X = 1
Y = 1
I = 1
Do While I <= Len(Str)
    Do While Mid(Str, I, 1) Like "[0-9,.,e,E,+,-]"
        Num = Num & Mid(Str, I, 1)
        I = I + 1
    Loop
    If Num <> "" Then
        Mat1(X, Y) = Val(Num)
        Num = ""
        If Y = UBound(Mat1, 2) Then
            X = X + 1
            Y = 1
        Else
            Y = Y + 1
        End If
    Else
        I = I + 1
    End If
Loop
End Sub

Private Function GetText(Mat() As Single) As String
Dim I As Integer, J As Integer, Str As String
For I = 1 To UBound(Mat, 1)
    For J = 1 To UBound(Mat, 2)
        Str = Str & Mat(I, J)
        If J <> UBound(Mat, 2) Then Str = Str & " "
    Next
    If I <> UBound(Mat, 1) Then Str = Str & vbCrLf
Next
GetText = Str
End Function

Private Sub Calculate_Click()
Dim Mat1() As Single, Mat2() As Single
Call GetMatrix(Mat1, Text1)
Call GetMatrix(Mat2, Text2)
Text4 = UBound(Mat1, 1)
Text5 = UBound(Mat1, 2)
Text6 = UBound(Mat2, 1)
Text7 = UBound(Mat2, 2)

Dim temp() As Single, Success As Boolean
Success = True
Select Case Combo1.ListIndex
    Case 0
        Success = Transpose(Mat1, temp)
    Case 1
        Success = Multiplication(Mat1, Mat2, temp)
    Case 2
        Success = Addition(Mat1, Mat2, temp)
    Case 3
        Success = AtoPowerOfN(Mat1, InputBox("Enter the power to which matrix A has to be raised?"), temp)
    Case 4
        Success = ValueDeterminant(Mat1, temp)
    Case 5
        Success = Adjoint(Mat1, temp)
    Case 6
        Success = Inverse(Mat1, temp)
    Case 7
        Call Combination(Val(InputBox("Enter the value of R?")), Val(InputBox("Enter the value of N?")), temp)
    Case 8
        Success = Rank(Mat1, temp)
    Case 9
        Success = LinearEquation(Mat1, Mat2, temp)
    Case Else
        Success = False
End Select

If Success = True Then
    Text3 = GetText(temp)
    Text8 = UBound(temp, 1)
    Text9 = UBound(temp, 2)
Else
    MsgBox ("The operation cannot be performed")
End If
End Sub

Private Function LinearEquation(Coefficient() As Single, Constant() As Single, Result() As Single) As Boolean
If UBound(Constant, 1) <> UBound(Coefficient, 1) Or UBound(Constant, 2) <> 1 Then
    LinearEquation = False
    Exit Function
End If

Dim RowCoeff As Integer, ColumnCoeff As Integer
RowCoeff = UBound(Coefficient, 1)
ColumnCoeff = UBound(Coefficient, 2)

'Augmented matrix is combination of Coefficient and Constant matrices
Dim Augmented() As Single, I As Integer
Augmented = Coefficient
ReDim Preserve Augmented(1 To RowCoeff, 1 To ColumnCoeff + 1)
For I = 1 To RowCoeff
    Augmented(I, ColumnCoeff + 1) = Constant(I, 1)
Next

'find ranks of Augmented and Coefficient matrix
Dim RankAug() As Single, RankCoeff() As Single
Call Rank(Augmented, RankAug)
Call Rank(Coefficient, RankCoeff)

If RankAug(1, 1) <> RankCoeff(1, 1) Then
    MsgBox ("Given equation(s) are inconsistent and have no solution.")
    ReDim Result(1 To 1, 1 To 1)
    Result(1, 1) = 0
    LinearEquation = True
    Exit Function
ElseIf RowCoeff >= ColumnCoeff Then     'consistent equations
    If RankAug(1, 1) = ColumnCoeff Then 'unique solution
        MsgBox ("Unique Solutions")
        
        'from Augmented matrix find out ColumnCoeff number of rows
        'to form a matrix TempAug whose rank is same as that of Augmented
        Dim Comb() As Single, TempAug() As Single
        Call Combination(RowCoeff, ColumnCoeff, Comb)
        ReDim TempAug(1 To ColumnCoeff, 1 To ColumnCoeff + 1)
        Dim J As Integer, Counter As Integer
        Counter = 1
        
        Dim RankTempAug() As Single
        Do While Counter <= UBound(Comb, 1)
            'fill the TempAug() matrix
            For I = 1 To ColumnCoeff
                For J = 1 To ColumnCoeff + 1
                    TempAug(I, J) = Augmented(Comb(Counter, I), J)
                Next
            Next
            
            Call Rank(TempAug, RankTempAug)
            If RankTempAug(1, 1) = ColumnCoeff Then
                Exit Do
            End If
            Counter = Counter + 1
        Loop
        
        ReDim Coefficient(1 To ColumnCoeff, 1 To ColumnCoeff)
        ReDim Constant(1 To ColumnCoeff, 1 To 1)
        For I = 1 To ColumnCoeff
            For J = 1 To ColumnCoeff
                Coefficient(I, J) = TempAug(I, J)
            Next
            Constant(I, 1) = TempAug(I, ColumnCoeff + 1)
        Next
        
        Dim Temp1() As Single, Temp2() As Single
        Call Inverse(Coefficient, Temp1)
        Call Multiplication(Temp1, Constant, Temp2)
        Result = Temp2
        LinearEquation = True
        Exit Function
    Else
        MsgBox ("Infinite solns..." & ColumnCoeff - RankAug(1, 1) & " arbitrary and " & RankAug(1, 1) & " in terms of them")
    End If
Else
    If RankAug(1, 1) = RowCoeff Then
        MsgBox ("Infinite solns..." & ColumnCoeff - RowCoeff & " arbitrary and " & RowCoeff & " in terms of them")
    Else
        MsgBox ("Infinite solns..." & ColumnCoeff - RankAug(1, 1) & " arbitrary and " & RankAug(1, 1) & " in terms of them")
    End If
End If
ReDim Result(1 To 1, 1 To 1)
Result(1, 1) = 0
LinearEquation = True
End Function

Private Function Rank(Mat() As Single, Result() As Single) As Boolean
ReDim Result(1 To 1, 1 To 1)

Dim Order As Integer, Row As Integer, column As Integer
Row = UBound(Mat, 1)
column = UBound(Mat, 2)
If Row < column Then
    Order = Row
Else
    Order = column
End If

Dim RowMat() As Single, ColumnMat() As Single, Determinant() As Single, temp() As Single
Dim I As Integer, J As Integer, P As Integer, Q As Integer
Do While Order >= 1
    Call Combination(Row, Order, RowMat)
    Call Combination(column, Order, ColumnMat)
    For I = 1 To UBound(RowMat, 1) 'now RowMat[I][] is one combination for row
        For J = 1 To UBound(ColumnMat, 1) 'and column[i][] is one combination for column
            'now fill the determinant
            ReDim Determinant(1 To Order, 1 To Order)  'variable determinant
            For P = 1 To Order
                For Q = 1 To Order
                    Determinant(P, Q) = Mat(RowMat(I, P), ColumnMat(J, Q))
                Next
            Next
            Call ValueDeterminant(Determinant, temp)
            If temp(1, 1) <> 0 Then
                Result(1, 1) = Order
                Rank = True
                Exit Function
            End If
        Next
    Next
    Order = Order - 1
Loop
Result(1, 1) = 0
Rank = False 'if all emements are 0
End Function

Private Sub Combination(R As Integer, N As Integer, Result() As Single)
'we have to select N positions out of R numbers which are 1,2,...,R
'total combinations possible are =                     R!
'                                     C(R,N)   =    ---------
'                                                    N!(R-N)!
Dim Total As Long
Total = Factorial(R) / (Factorial(N) * Factorial(R - N))

ReDim Result(1 To Total, 1 To N)
Dim I As Integer, J As Integer, P As Integer, Arr() As Integer
ReDim Arr(1 To N) As Integer
For I = 1 To N
    Arr(I) = I
Next
P = 1

Label:
    For I = 1 To N
        Result(P, I) = Arr(I)
    Next
    P = P + 1
    For I = N To 1 Step -1
        If Arr(I) <> R - (N - I) Then
            Arr(I) = Arr(I) + 1
            For J = I + 1 To N
                Arr(J) = Arr(J - 1) + 1
            Next
            GoTo Label
        End If
    Next
End Sub

Private Function Factorial(N As Integer) As Long
'factorial of N
Dim Fact As Long, I As Integer
Fact = 1
If N = 0 Then
    Factorial = 1
Else
    For I = 1 To N
        Fact = Fact * I
    Next
End If
Factorial = Fact
End Function

Private Function Inverse(Mat() As Single, Result() As Single) As Boolean
If UBound(Mat, 1) <> UBound(Mat, 2) Then
    Inverse = False
    Exit Function
Else
    Call ValueDeterminant(Mat, Result)
    If Result(1, 1) = 0 Then
        Inverse = False
        Exit Function
    End If
End If
Dim Value As Single, I As Integer, J As Integer
Value = Result(1, 1)
Call Adjoint(Mat, Result)
For I = 1 To UBound(Result, 1)
    For J = 1 To UBound(Result, 2)
        Result(I, J) = Result(I, J) / Value
    Next
Next
Inverse = True
End Function

Private Function Adjoint(Mat() As Single, Result() As Single) As Boolean
If UBound(Mat, 1) <> UBound(Mat, 2) Then
    Adjoint = False
    Exit Function
End If

'cofactor matrice=CoFactor and matrix of adjoint = Adjnt
Dim Cofactor() As Single, Adjnt() As Single, Order As Integer
Order = UBound(Mat, 1)
If Order = 1 Then
    ReDim Adjnt(1 To 1, 1 To 1)
    Adjnt(1, 1) = 1
    Adjoint = True
    Exit Function
End If
 
ReDim Cofactor(1 To Order - 1, 1 To Order - 1)
ReDim Adjnt(1 To Order, 1 To Order)
Dim temp() As Single
Dim P As Integer, Q As Integer, I As Integer, J As Integer, X As Integer, Y As Integer
For I = 1 To Order
  For J = 1 To Order
    'for element (I,J) in Mat find Cofactor matrix
    P = 1
    Q = 1
    For X = 1 To Order
      For Y = 1 To Order
        'traverse each element (X,Y) of Mat
        If X <> I And Y <> J Then   'by curbing Ith row and Jth column
            Cofactor(P, Q) = Mat(X, Y)
            If P = Order - 1 Then
                Q = Q + 1
                P = 1
            Else
                P = P + 1
            End If
        End If
      Next
    Next
    'now we get minor of Mat corresponding to (I,J) in 'CoFactor'
    Call ValueDeterminant(Cofactor, temp)   'store the value of determinant(minor) in 'Temp'
    'Cofactor is [(-1)^(i+j)]*determinant(minor)
    If (I + J) Mod 2 <> 0 Then temp(1, 1) = -temp(1, 1)
    'put this value in main determinant
    Adjnt(I, J) = temp(1, 1)
  Next
Next
Call Transpose(Adjnt, Result) 'Adjoint is transpose of this matrix
Adjoint = True
End Function

Private Function ValueDeterminant(MatOriginal() As Single, Result() As Single) As Boolean
ReDim Result(1 To 1, 1 To 1)

If UBound(MatOriginal, 1) <> UBound(MatOriginal, 2) Then
    ValueDeterminant = False
    Exit Function
ElseIf UBound(MatOriginal, 1) = 1 Then
    Result(1, 1) = MatOriginal(1, 1)
    ValueDeterminant = True
    Exit Function
End If

Dim Order As Integer
Order = UBound(MatOriginal, 1)

'make a new matrix MatMinor of order 1 less than given MatOriginal
Dim MatMinor() As Single
ReDim MatMinor(1 To Order - 1, 1 To Order - 1)

'variables used
Dim Sum As Single, I As Integer, J As Integer
Dim RowMatOriginal As Integer, ColumnMatOriginal As Integer
Dim RowMatMinor As Integer, ColumnMatMinor As Integer
Sum = 0

'this is coloumn 1 on which we are opening determinant MatOriginal
ColumnMatOriginal = 1
For RowMatOriginal = 1 To Order
    RowMatMinor = 1
    ColumnMatMinor = 1
    'corresponding to (RowMatOriginal,ColumnMatOriginal) element of MatOriginal
    'form the minor determinant from MatOriginal to store in MatMinor
    For I = 1 To Order
        For J = 1 To Order
            'leaving the row and column (RowMatOriginal,ColumnMatOriginal)
            If I <> RowMatOriginal And J <> ColumnMatOriginal Then
                MatMinor(RowMatMinor, ColumnMatMinor) = MatOriginal(I, J)
                If ColumnMatMinor = Order - 1 Then
                    RowMatMinor = RowMatMinor + 1
                    ColumnMatMinor = 1
                Else
                ColumnMatMinor = ColumnMatMinor + 1
                End If
            End If
        Next
    Next
    'So we get MatMinor
    'add     ValueDeterminant(MatMinor)*MatOriginal(RowMatOriginal,ColumnMatOriginal)     to Sum;
    Call ValueDeterminant(MatMinor, Result) 'store value of determinant MatMinor in Result
    If RowMatOriginal Mod 2 = 0 Then
        Sum = Sum - MatOriginal(RowMatOriginal, ColumnMatOriginal) * Result(1, 1)
    Else
        Sum = Sum + MatOriginal(RowMatOriginal, ColumnMatOriginal) * Result(1, 1)
    End If
Next
Result(1, 1) = Sum
ValueDeterminant = True
End Function

Private Function AtoPowerOfN(Mat() As Single, N As Integer, Result() As Single) As Boolean
If UBound(Mat, 1) <> UBound(Mat, 2) Then
    AtoPowerOfN = False
    Exit Function
End If
Dim I As Integer, J As Integer, temp() As Single
Result = Mat
temp = Mat
For I = 1 To N - 1
    J = Multiplication(Result, Mat, temp)
    Result = temp
Next
AtoPowerOfN = True
End Function

Private Function Addition(Mat1() As Single, Mat2() As Single, Result() As Single) As Boolean
If UBound(Mat1, 1) <> UBound(Mat2, 1) Or UBound(Mat1, 2) <> UBound(Mat2, 2) Then
    Addition = False
    Exit Function
End If
Result = Mat1
Dim I As Integer, J As Integer
For I = 1 To UBound(Mat1, 1)
    For J = 1 To UBound(Mat1, 2)
            Result(I, J) = Mat1(I, J) + Mat2(I, J)
    Next
Next
Addition = True
End Function

Private Function Multiplication(Mat1() As Single, Mat2() As Single, Result() As Single) As Boolean
If UBound(Mat1, 2) <> UBound(Mat2, 1) Then
    Multiplication = False
    Exit Function
End If

ReDim Preserve Result(1 To UBound(Mat1, 1), 1 To UBound(Mat2, 2))
Dim I As Integer, J As Integer, K As Integer
For I = 1 To UBound(Mat1, 1)
    For J = 1 To UBound(Mat2, 2)
        Result(I, J) = 0
        For K = 1 To UBound(Mat1, 2)
            Result(I, J) = Result(I, J) + Mat1(I, K) * Mat2(K, J)
        Next
    Next
Next
Multiplication = True
End Function

Private Function Transpose(Mat() As Single, Result() As Single) As Boolean
Result = Mat
Dim I As Integer, J As Integer
For I = 1 To UBound(Mat, 1)
    For J = 1 To UBound(Mat, 2)
        Result(I, J) = Mat(J, I)
    Next
Next
Transpose = True
End Function

Private Sub Command2_Click()
Unload Form1
End Sub

Private Sub Form_Load()
Dim v As String
v = vbCrLf
Text1 = "1 2 3" & v & "4 5 6" & v & "7 8 9"
Text2 = "1 2 3" & v & "4 5 6" & v & "7 8 9"
Text4 = "3"
Text5 = "3"
Text6 = "3"
Text7 = "3"
End Sub
