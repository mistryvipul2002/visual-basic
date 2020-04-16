VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "form1.frx":0000
      Left            =   5280
      List            =   "form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Options"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Points() As Single, Join() As Integer
Dim I As Integer, TempX As Integer, TempY As Integer

Public Sub Combo1_Click()
Cls
If Combo1.Text = Combo1.List(0) Then
    Call AssignValuesHexagon(Points())
ElseIf Combo1.Text = Combo1.List(1) Then
    Call AssignValuesTetrahedron(Points())
Else: Call AssignValuesCuboid(Points())
End If
DrawObject Me.ForeColor
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Me.Show
Label1.Width = TextWidth("Mouse :- (" & ScaleWidth & "," & ScaleHeight & ")  ")
Label1.Left = ScaleWidth - Label1.Width
'scaling of the form
Scale (-ScaleWidth / 2, ScaleHeight / 2)-(ScaleWidth / 2, -ScaleHeight / 2)
Combo1.Text = Combo1.List(0)
End Sub

Private Sub AssignValuesHexagon(Points() As Single)
ReDim Points(1 To 12, 1 To 3) As Single
Dim hi As Single, s3 As Single
s3 = Sqr(3)
hi = 5
Call PutData1(Points(), 1, 1, 0, hi / 2)
Call PutData1(Points(), 2, 0.5, s3 / 2, hi / 2)
Call PutData1(Points(), 3, -0.5, s3 / 2, hi / 2)
Call PutData1(Points(), 4, -1, 0, hi / 2)
Call PutData1(Points(), 5, -0.5, -s3 / 2, hi / 2)
Call PutData1(Points(), 6, 0.5, -s3 / 2, hi / 2)
Call PutData1(Points(), 7, 1, 0, -hi / 2)
Call PutData1(Points(), 8, 0.5, s3 / 2, -hi / 2)
Call PutData1(Points(), 9, -0.5, s3 / 2, -hi / 2)
Call PutData1(Points(), 10, -1, 0, -hi / 2)
Call PutData1(Points(), 11, -0.5, -s3 / 2, -hi / 2)
Call PutData1(Points(), 12, 0.5, -s3 / 2, -hi / 2)
For I = 1 To UBound(Points)
    Points(I, 1) = Points(I, 1) * 20
    Points(I, 2) = Points(I, 2) * 20
    Points(I, 3) = Points(I, 3) * 20
Next

ReDim Join(1 To 18, 1 To 2) As Integer
Call PutData2(Join(), 1, 5, 6)
Call PutData2(Join(), 2, 1, 2)
Call PutData2(Join(), 3, 2, 3)
Call PutData2(Join(), 4, 3, 4)
Call PutData2(Join(), 5, 4, 5)
Call PutData2(Join(), 6, 6, 1)
Call PutData2(Join(), 7, 11, 12)
Call PutData2(Join(), 8, 7, 8)
Call PutData2(Join(), 9, 8, 9)
Call PutData2(Join(), 10, 9, 10)
Call PutData2(Join(), 11, 10, 11)
Call PutData2(Join(), 12, 12, 7)
Call PutData2(Join(), 13, 6, 12)
Call PutData2(Join(), 14, 1, 7)
Call PutData2(Join(), 15, 2, 8)
Call PutData2(Join(), 16, 3, 9)
Call PutData2(Join(), 17, 4, 10)
Call PutData2(Join(), 18, 5, 11)
End Sub

Private Sub AssignValuesTetrahedron(Points() As Single)
ReDim Points(1 To 4, 1 To 3) As Single
Dim Size As Integer, s3 As Single
s3 = Sqr(3)
Size = 100
Call PutData1(Points(), 1, Size / s3, 0, 0)
Call PutData1(Points(), 2, -Size / (2 * s3), Size / 2, 0)
Call PutData1(Points(), 3, -Size / (2 * s3), -Size / 2, 0)
Call PutData1(Points(), 4, 0, 0, (Sqr(2#) / s3) * Size)

ReDim Join(1 To 6, 1 To 2) As Integer
Call PutData2(Join(), 1, 1, 2)
Call PutData2(Join(), 2, 1, 3)
Call PutData2(Join(), 3, 1, 4)
Call PutData2(Join(), 4, 2, 3)
Call PutData2(Join(), 5, 2, 4)
Call PutData2(Join(), 6, 3, 4)
End Sub

Private Sub AssignValuesCuboid(Points() As Single)
ReDim Points(1 To 8, 1 To 3) As Single
Dim X As Integer, Y As Integer, Z As Integer
X = 20
Y = 30
Z = 40
Call PutData1(Points(), 1, X / 2, Y / 2, -Z / 2)
Call PutData1(Points(), 2, X / 2, Y / 2, Z / 2)
Call PutData1(Points(), 3, -X / 2, Y / 2, Z / 2)
Call PutData1(Points(), 4, -X / 2, Y / 2, -Z / 2)
Call PutData1(Points(), 5, X / 2, -Y / 2, -Z / 2)
Call PutData1(Points(), 6, X / 2, -Y / 2, Z / 2)
Call PutData1(Points(), 7, -X / 2, -Y / 2, Z / 2)
Call PutData1(Points(), 8, -X / 2, -Y / 2, -Z / 2)

ReDim Join(1 To 12, 1 To 2) As Integer
Call PutData2(Join(), 1, 1, 2)
Call PutData2(Join(), 2, 2, 3)
Call PutData2(Join(), 3, 3, 4)
Call PutData2(Join(), 4, 4, 1)
Call PutData2(Join(), 5, 5, 6)
Call PutData2(Join(), 6, 6, 7)
Call PutData2(Join(), 7, 7, 8)
Call PutData2(Join(), 8, 8, 5)
Call PutData2(Join(), 9, 1, 5)
Call PutData2(Join(), 10, 4, 8)
Call PutData2(Join(), 11, 3, 7)
Call PutData2(Join(), 12, 2, 6)
End Sub

Private Sub PutData1(Pts() As Single, Index As Integer, X As Variant, Y As Variant, Z As Variant)
Pts(Index, 1) = X
Pts(Index, 2) = Y
Pts(Index, 3) = Z
End Sub

Private Sub PutData2(Pts() As Integer, Index As Integer, X As Variant, Y As Variant)
Pts(Index, 1) = X
Pts(Index, 2) = Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TempX = X
TempY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Mouse :- (" & X & "," & Y & ")"
Dim HorizontalAngle As Integer, VerticalAngle As Integer
If Button = vbLeftButton Then
    DrawObject Me.BackColor
    HorizontalAngle = X - TempX
    VerticalAngle = Y - TempY
    Call Rotate(HorizontalAngle, VerticalAngle)
    DrawObject Me.ForeColor
    TempX = X
    TempY = Y
End If
End Sub

Private Sub Rotate(XAngle As Integer, YAngle As Integer)
Dim QX As Single, QY As Single, X As Single, Y As Single, Z As Single
QX = (XAngle / 180) * 3.14
QY = (YAngle / 180) * 3.14
For I = 1 To UBound(Points)
    X = Points(I, 1)
    Z = Points(I, 3)
    'rotation for XAngle i.e. about Y-Axis
    Points(I, 3) = Z * Cos(QX) - X * Sin(QX)
    Points(I, 1) = Z * Sin(QX) + X * Cos(QX)
    Y = Points(I, 2)
    Z = Points(I, 3)
    'rotation for YAngle i.e. about X-Axis
    Points(I, 2) = Y * Cos(-QY) - Z * Sin(-QY)
    Points(I, 3) = Y * Sin(-QY) + Z * Cos(-QY)
Next
End Sub

Private Sub DrawObject(Color As Long)
Line (-ScaleWidth / 2, 0)-(ScaleWidth / 2, 0)
Line (0, ScaleHeight / 2)-(0, -ScaleHeight / 2)
'draw the object
Dim TempX1 As Single, TempY1 As Single, TempX2 As Single, TempY2 As Single
Dim D As Integer
D = 300
For I = 1 To UBound(Join)
    TempX1 = -D * Points(Join(I, 1), 1) / (-D + Points(Join(I, 1), 3))
    TempX2 = -D * Points(Join(I, 2), 1) / (-D + Points(Join(I, 2), 3))
    TempY1 = -D * Points(Join(I, 1), 2) / (-D + Points(Join(I, 1), 3))
    TempY2 = -D * Points(Join(I, 2), 2) / (-D + Points(Join(I, 2), 3))
    If TempX1 = TempX2 And TempY1 = TempY2 Then
        PSet (TempX1, TempX2), Color
    Else: Line (TempX1, TempY1)-(TempX2, TempY2), Color
    End If
Next
'code for orthographic view
'For I = 1 To UBound(Join)
'    If Points(Join(I, 1), 1) = Points(Join(I, 2), 1) And Points(Join(I, 1), 2) = Points(Join(I, 2), 2) Then
'        PSet (Points(Join(I, 2), 1), Points(Join(I, 2), 2)), Color
'    Else: Line (Points(Join(I, 1), 1), Points(Join(I, 1), 2))-(Points(Join(I, 2), 1), Points(Join(I, 2), 2)), Color
'    End If
'Next
End Sub

Public Sub Changes()
Cls
For I = 1 To UBound(Points)
    Points(I, 1) = Points(I, 1) * CSng(Form2.Text1)
    Points(I, 2) = Points(I, 2) * CSng(Form2.Text1)
    Points(I, 3) = Points(I, 3) * CSng(Form2.Text1)
    Points(I, 1) = Points(I, 1) + CSng(Form2.Text2)
    Points(I, 2) = Points(I, 2) + CSng(Form2.Text3)
    Points(I, 3) = Points(I, 3) + CSng(Form2.Text4)
Next
DrawObject Me.ForeColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

