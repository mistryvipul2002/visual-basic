VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Lexical Analyser"
   ClientHeight    =   6825
   ClientLeft      =   1620
   ClientTop       =   1500
   ClientWidth     =   4530
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   4530
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   4335
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Text            =   "BEGIN IF A>=B THEN C=25.56 ELSE D=34 END  "
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Token ID     Token Name"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Input string: -"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Output of Lexical Analyser: -"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Symbol Table: -"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'symbol table is used to store only keywords, identifiers and numbers
Dim TokenId(1 To 50) As String   'symbol table tokenid
Dim TokenName(1 To 50) As String 'symbol table tokenname
Dim NumberofElements As Integer 'no. of tokens in symbol table recognised so far

'input string
Dim InputString As String

'variables for scaning of input string for search of tokens
Dim LookAheadPosition As Integer    'current position of scanning in input string
Dim State As Integer        'state of transition diagram
Dim InputBuffer As String   'whatever chacter is read from inmput string is stored in this and this is cleared when a token is found
Dim StartState As Integer   'starting state for token search generally 0
Dim TokenBeginning As Integer   'position in input string where token search has been started

'function called when a token is found
Private Sub TokenFound()
    TokenBeginning = LookAheadPosition      'set TokenBeginning to next character
    StartState = 0      'set starting state to 0
    InputBuffer = ""    'empty the buffer when the token is found
End Sub

'calculate buttton clicked
Private Sub Command1_Click()
InputString = Text1.Text & " "
Text3 = ""
Text2 = ""
StartState = 0
TokenBeginning = 1

Dim Str As String
Str = NextToken()
Do While Str <> "EndToken"
    Call TokenFound
    If Str <> "" Then Text2 = Text2 & Str & vbCrLf
    Str = NextToken()
Loop

'code to print the symbol table in text3
Dim I As Integer
For I = 1 To NumberofElements
    Text3 = Text3 & I & ".)  " & TokenId(I) & "      " & TokenName(I) & vbCrLf
Next
End Sub

'searches the symbol table to match a TokenName with InputBuffer
'if found returns the index of the entry but in case of keyword returns 0
'if not found makes a new entry for the new Identifier and returns its index
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

'searches the symbol table to match a TokenName with InputBuffer
'if found returns the index of the entry but in case of keyword returns 0
'if not found makes a new entry for the new Number and returns its index
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
    'if the string is not found, make a new entry for number
    NumberofElements = NumberofElements + 1
    TokenId(NumberofElements) = "NUM"
    TokenName(NumberofElements) = InputBuffer
    InstallNum = NumberofElements
End Function

'Returns the TokenId of token by findng the InputBuffer in the TokenName of the symbol table
'Since InstallId is called before this function so it will surely find a entry in the symbol table
Private Function GetToken() As String
    Dim I As Integer
    For I = 1 To NumberofElements
        If InputBuffer = TokenName(I) Then
            GetToken = TokenId(I)
            Exit Function
        End If
    Next
End Function

'If a transition diagram fails in between we start the token recognition
'from the next transition diagram and place the LookAheadPosition to TokenBeginning
Private Function Fail() As Integer
    LookAheadPosition = TokenBeginning
    InputBuffer = ""
    Select Case StartState
        Case 0          'transition diagram for relational operators
            StartState = 9
        Case 9          'transition diagram for identifiers or keywords
            StartState = 12
        Case 12         'transition diagram for numbers of the form a.bE(+/-)c
            StartState = 20
        Case 20         'transition diagram for numbers of the form a.b
            StartState = 25
        Case 25         'transition diagram for numbers of the form a
            StartState = 28
        Case 28         'transition diagram for white spaces
            MsgBox "Compile Error"
            End
    End Select
    Fail = StartState
End Function

'retracts the LookAheadPosition by I steps back
Private Sub Retract(I As Integer)
    LookAheadPosition = LookAheadPosition - I
    InputBuffer = Mid(InputBuffer, 1, Len(InputBuffer) - I)
End Sub

'returns the next character of InputString and put it in InputBuffer
Private Function NextCharacter() As String
    NextCharacter = Mid(InputString, LookAheadPosition, 1)
    LookAheadPosition = LookAheadPosition + 1
    InputBuffer = InputBuffer & NextCharacter
End Function

'returns the token appearing in the InputString starting from the position
'TokenBeginning and starting state to be StartState
Private Function NextToken() As String
Dim Tmp As Integer
Dim Character As String * 1

LookAheadPosition = TokenBeginning
State = StartState

'scan characters of input string
Do While LookAheadPosition <= Len(InputString)
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
'when all the input string is scanned return TokenId as endtoken
NextToken = "EndToken"
End Function

'cancel clicked
Private Sub Command2_Click()
Unload Form1
End Sub

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

