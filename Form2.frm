VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Painting"
   ClientHeight    =   7365
   ClientLeft      =   930
   ClientTop       =   3555
   ClientWidth     =   12585
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   12585
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "显示控制面板"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   1680
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show '显示控制面板
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height - 100
Me.Width = Screen.Width
End Sub
Public Sub PaintCoord()
Dim a As Integer, b As Integer, c As Integer, d As Integer
a = xleft1
b = xright1
c = ytop1
d = ybottom1
If a * b > 0 Then
GoTo a
End If
If c * d > 0 Then
GoTo a
End If
Call reference
Form2.ForeColor = vbBlack
Line (xleft1, 0)-(xright1, 0)
Line (0, ytop1)-(0, ybottom1)
CurrentX = 0.95 * xright1: CurrentY = 0.01 * ybottom1: Print "X轴"
CurrentX = 0.01 * xright1: CurrentY = 0.95 * ytop1: Print "Y轴"
CurrentX = 0.005 * xright1: CurrentY = 0.005 * ybottom1: Print 0 '原点标点
If (paint = 1) Then
For i = a To -1
Line (i, 0.005 * ytop1)-(i, -0.005 * ytop1)  'x轴左侧画竖线以及标点
CurrentX = i: CurrentY = 0.01 * ybottom1: Print i
Next i
For i = 1 To b
Line (i, 0.005 * ytop1)-(i, -0.005 * ytop1) 'x轴右侧画竖线以及标点
CurrentX = i: CurrentY = 0.01 * ybottom1: Print i
Next i
For i = 1 To c
Line (0.005 * xright1, i)-(-0.005 * xright1, i) 'y轴上侧画竖线以及标点
CurrentX = 0.01 * xright1: CurrentY = i: Print i
Next i
For i = d To -1
Line (0.005 * xright1, i)-(-0.005 * xright1, i) 'y轴下侧画竖线以及标点
CurrentX = 0.01 * xright1: CurrentY = i: Print i
Next i
End If
a:
End Sub
Public Sub sinx()
Dim a(3) As Double
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0) 'a(1)a(3)老位置a(0)a(2)新位置
a(0) = j
a(2) = Round(Sin(a(0)), 4)
a(3) = Round(Sin(a(1)), 4)
If linestyle = 1 Then
Line (a(0), a(2))-(a(1), a(3)), RGB(rred, ggreen, bblue)
End If
If linestyle = 2 Then
CurrentX = a(0): CurrentY = a(2): Print linefonts
End If
Next j
End Sub

Public Sub cosx()
Dim a(3) As Double
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0) 'a(1)a(3)老位置a(0)a(2)新位置
a(0) = j
a(2) = Round(Cos(a(0)), 4)
a(3) = Round(Cos(a(1)), 4)
If linestyle = 1 Then
Line (a(0), a(2))-(a(1), a(3)), RGB(rred, ggreen, bblue)
End If
If linestyle = 2 Then
CurrentX = a(0): CurrentY = a(2): Print linefonts
End If
Next j
End Sub

Public Sub tanx()
Dim a(3) As Double
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0) 'a(1)a(3)老位置a(0)a(2)新位置
a(0) = j
a(2) = Round(Tan(a(0)), 4)
a(3) = Round(Tan(a(1)), 4)
If linestyle = 1 Then
Line (a(0), a(2))-(a(1), a(3)), RGB(rred, ggreen, bblue)
End If
If linestyle = 2 Then
CurrentX = a(0): CurrentY = a(2): Print linefonts
End If
Next j
End Sub
Public Sub kx()
Dim a(3) As Double
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0)
a(0) = j
a(2) = Round(Val(Form1.x) * a(0) + Val(Form1.y), 4)
a(3) = Round(Val(Form1.x) * a(1) + Val(Form1.y), 4)
If linestyle = 1 Then
Line (a(0), a(2))-(a(1), a(3)), RGB(rred, ggreen, bblue)
End If
If linestyle = 2 Then
CurrentX = a(0): CurrentY = a(2): Print linefonts
End If
Next j
End Sub


Public Sub reference()
If kind = 1 Then
Form2.Scale (xleft1, ytop1)-(xright1, ybottom1)
End If
End Sub

Public Sub ajimi()
Dim a(4) As Double
Dim charac As String
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0)
a(0) = j
a(2) = Round((1 / 8) * a(0) * Sin(a(0)), 4)
a(3) = Round((1 / 8) * a(1) * Sin(a(1)), 4)
If linestyle = 1 Then
charac = "·"
CurrentX = Round((1 / 8) * a(1) * Cos(a(1)), 4): CurrentY = a(3): Print charac
End If
If linestyle = 2 Then
charac = linefonts
CurrentX = Round((1 / 8) * a(1) * Cos(a(1)), 4): CurrentY = a(3): Print charac
End If
Next j
End Sub

Public Sub kxx()
Dim a(3) As Double
a(0) = ibegin
For j = ibegin To iendup Step linestep
a(1) = a(0)
a(0) = j
a(2) = Round(Val(Form1.x) * a(0) * a(0) + Val(Form1.y), 4)
a(3) = Round(Val(Form1.x) * a(1) * a(1) + Val(Form1.y), 4)
If linestyle = 1 Then
Line (a(0), a(2))-(a(1), a(3)), RGB(rred, ggreen, bblue)
End If
If linestyle = 2 Then
CurrentX = a(0): CurrentY = a(2): Print linefonts
End If
Next j
End Sub
