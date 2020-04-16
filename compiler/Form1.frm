VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Lexical Analyser"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter the expression below: -"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'symbol table is used to store only keywords, identifiers and numbers
Dim TokenId(1 To 50) As String   'symbol table
Dim TokenName(1 To 50) As String 'symbol table
Dim NumberofElements As Integer 'no. of tokens in symbol table recognised so far

'input and output strings
Dim OutputString As String
Dim InputString As String

'variables for scaning
Dim CurrentPosition As Integer
Dim State As Integer
Dim InputBuffer As String
Dim StartState As Integer
Dim TokenBeginning As Integer

Private Function InstallId() As Integer
    Dim I As Integer
    For I = 1 To NumberofElements
        If InputBuffer = TokenName(I) Then
            InstallId = I
            'if tokenid is a keyword return 0 otherwise return index of the entry
            If TokenId(I) <> "ID" And TokenId(I) <> "NUM" Then InstallId = 0
            Exit Function
        End If
    Next
    'if the string is not found, make a new entry for identifier
    NumberofElements = NumberofElements + 1
    TokenId(NumberofElements) = "ID"
    TokenName(NumberofElements) = InputBuffer
    InstallId = NumberofElements
End Function

Private Function InstallNum() As Integer
    Dim I As Integer
    For I = 1 To NumberofElements
        If InputBuffer = TokenName(I) Then
            InstallNum = I
            'if tokenid is a keyword return 0 otherwise return index of the entry
            If TokenId(I) <> "ID" And TokenId(I) <> "NUM" Then InstallNum = 0
            Exit Function
        End If
    Next
    'if the string is not found, make a new entry for identifier
    NumberofElements = NumberofElements + 1
    TokenId(NumberofElements) = "NUM"
    TokenName(NumberofElements) = InputBuffer
    InstallNum = NumberofElements
End Function

Private Function GetToken() As String
    Dim I As Integer
    For I = 1 To NumberofElements
        If InputBuffer = TokenName(I) Then
            GetToken = TokenId(I)
            Exit Function
        End If
    Next
End Function

Private Function Fail() As Integer
    CurrentPosition = TokenBeginning
    InputBuffer = ""
    Select Case StartState
        Case 0
            StartState = 9
        Case 9
            StartState = 12
        Case 12
            StartState = 20
        Case 20
            StartState = 25
        Case 25
            StartState = 28
        Case 28
            MsgBox "Compile Error"
            End
    End Select
    Fail = StartState
End Function

Private Sub Retract(I As Integer)
    CurrentPosition = CurrentPosition - I
    InputBuffer = Mid(InputBuffer, 1, Len(InputBuffer) - I)
End Sub

Private Function NextCharacter() As String
    NextCharacter = Mid(InputString, CurrentPosition, 1)
    CurrentPosition = CurrentPosition + 1
    InputBuffer = InputBuffer & NextCharacter
End Function

Private Sub TokenFound()
    TokenBeginning = CurrentPosition
    StartState = 0
    InputBuffer = ""    'empty the buffer when the token is found
End Sub

Private Sub Done_Click()
InputString = Text1.Text & " "

StartState = 0
TokenBeginning = 1

Dim Str As String
Str = NextToken()
Do While Str <> "EndToken"
    Call TokenFound
    OutputString = OutputString & Str
    Str = NextToken()
Loop

Unload Form1
Load Form2
Form2.Label6.Caption = InputString
Form2.Label7.Caption = OutputString

'code to print the symbol table in text3
Dim I As Integer
For I = 1 To NumberofElements
    Form2.Label8.Caption = Form2.Label8.Caption & TokenId(I) & vbCrLf
    Form2.Label9.Caption = Form2.Label9.Caption & TokenName(I) & vbCrLf
Next
Form2.Show
End Sub

Private Function NextToken() As String
Dim Tmp As Integer
Dim Character As String * 1

CurrentPosition = TokenBeginning
State = StartState

'scan characters of input string
Do While CurrentPosition <= Len(InputString)
    Select Case State
        Case 0
            Character = NextCharacter()
            If Character = "<" Then
                State = 1
            ElseIf Character = "=" Then 'token found so returned
                NextToken = "<relop,EQ>"
                Exit Function
            ElseIf Character = ">" Then
                State = 6
            Else
                State = Fail()
            End If
        Case 1
            Character = NextCharacter()
            If Character = "=" Then
                NextToken = "<relop,LE>"
                Exit Function
            ElseIf Character = ">" Then
                NextToken = "<relop,NE>"
                Exit Function
            Else
                Retract (1)
                NextToken = "<relop,LT>"
                Exit Function
            End If
        Case 6
            Character = NextCharacter()
            If Character = "=" Then
                NextToken = "<relop,GE>"
                Exit Function
            Else
                Retract (1)
                NextToken = "<relop,GT>"
                Exit Function
            End If
        Case 9
            Character = NextCharacter()
            If Character Like "[A-Z,a-z]" Then  'letter
                State = 10
            Else
                State = Fail()
            End If
        Case 10
            Character = NextCharacter()
            If Character Like "[A-Z,a-z,0-9]" Then  'letter or digit
                State = 10
            Else
                Retract (1)
                Tmp = InstallId() 'calling the function
                NextToken = "<" & GetToken() & "," & Tmp & ">"
                Exit Function
            End If
        Case 12
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 13
            Else
                State = Fail()
            End If
        Case 13
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 13
            ElseIf Character = "." Then
                State = 14
            ElseIf Character = "E" Then
                State = 16
            Else
                State = Fail()
            End If
        Case 14
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 15
            Else
                State = Fail()
            End If
        Case 15
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 15
            ElseIf Character = "E" Then
                State = 16
            Else
                State = Fail()
            End If
        Case 16
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 18
            ElseIf Character = "+" Or Character = "-" Then
                State = 17
            Else
                State = Fail()
            End If
        Case 17
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 18
            Else
                State = Fail()
            End If
        Case 18
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 18
            Else
                Retract (1)
                Tmp = InstallNum() 'calling the function
                NextToken = "<" & GetToken() & "," & Tmp & ">"
                Exit Function
            End If
        Case 20
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 21
            Else
                State = Fail()
            End If
        Case 21
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 21
            ElseIf Character = "." Then
                State = 22
            Else
                State = Fail()
            End If
        Case 22
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 23
            Else
                State = Fail()
            End If
        Case 23
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 23
            Else
                Retract (1)
                Tmp = InstallNum() 'calling the function
                NextToken = "<" & GetToken() & "," & Tmp & ">"
                Exit Function
            End If
        Case 25
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 26
            Else
                State = Fail()
            End If
        Case 26
            Character = NextCharacter()
            If Character Like "[0-9]" Then  'digit
                State = 26
            Else
                Retract (1)
                Tmp = InstallNum() 'calling the function
                NextToken = "<" & GetToken() & "," & Tmp & ">"
                Exit Function
            End If
        Case 28
            Character = NextCharacter()
            If Character = " " Then
                State = 29
            Else
                State = Fail()
            End If
        Case 29
            Character = NextCharacter()
            If Character = " " Then 'white space
                State = 29
            Else
                Retract (1)
                NextToken = ""
                Exit Function
            End If
    End Select
Loop
NextToken = "EndToken"
End Function

Private Sub Form_Load()
'initialisation of symbol table for keywords
TokenId(1) = "BEGIN"
TokenId(2) = "END"
TokenId(3) = "IF"
TokenId(4) = "THEN"
TokenId(5) = "ELSE"
TokenName(1) = "BEGIN"
TokenName(2) = "END"
TokenName(3) = "IF"
TokenName(4) = "THEN"
TokenName(5) = "ELSE"
NumberofElements = 5
End Sub
