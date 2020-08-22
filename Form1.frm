VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   5535
   ClientLeft      =   19545
   ClientTop       =   1740
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8595
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7800
      Top             =   4800
   End
   Begin VB.Frame Frame5 
      Caption         =   "关于"
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      Begin VB.Frame Frame21 
         Caption         =   "MY VOICE"
         Height          =   3255
         Left            =   2880
         TabIndex        =   59
         Top             =   840
         Width           =   4575
         Begin VB.Label Label14 
            Caption         =   "时间治愈的是愿意自渡的人"
            BeginProperty Font 
               Name            =   "迷你简书魂"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   360
            TabIndex        =   61
            Top             =   2160
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "开发者：Gancenty    QQ：2539797953 "
            BeginProperty Font 
               Name            =   "微软雅黑 Light"
               Size            =   18
               Charset         =   134
               Weight          =   290
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   600
            TabIndex        =   60
            Top             =   600
            Width           =   3255
         End
      End
      Begin VB.Image Image1 
         Height          =   1620
         Left            =   600
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1620
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "画笔颜色及绘制频度"
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      Begin VB.Timer Timer5 
         Interval        =   100
         Left            =   360
         Top             =   4440
      End
      Begin VB.Frame Frame19 
         Caption         =   "RGB"
         Height          =   3735
         Left            =   480
         TabIndex        =   50
         Top             =   600
         Width           =   3495
         Begin VB.Frame Frame20 
            Caption         =   "展示颜色"
            Height          =   1215
            Left            =   480
            TabIndex        =   54
            Top             =   2280
            Width           =   2415
            Begin VB.Label Label5 
               Height          =   495
               Left            =   360
               TabIndex        =   55
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.HScrollBar bb 
            Height          =   495
            Left            =   960
            TabIndex        =   53
            Top             =   1560
            Width           =   2175
         End
         Begin VB.HScrollBar gg 
            Height          =   495
            Left            =   960
            TabIndex        =   52
            Top             =   960
            Width           =   2175
         End
         Begin VB.HScrollBar rr 
            Height          =   495
            Left            =   960
            TabIndex        =   51
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label13 
            Height          =   375
            Left            =   360
            TabIndex        =   58
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label12 
            Height          =   375
            Left            =   360
            TabIndex        =   57
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label11 
            Height          =   375
            Left            =   360
            TabIndex        =   56
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "重新绘制"
         Height          =   735
         Left            =   4920
         TabIndex        =   49
         Top             =   1920
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "线条样式"
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8295
      Begin VB.Timer Timer4 
         Interval        =   100
         Left            =   600
         Top             =   4080
      End
      Begin VB.Frame Frame18 
         Caption         =   "频度"
         Height          =   1455
         Left            =   2760
         TabIndex        =   46
         Top             =   2160
         Width           =   3975
         Begin VB.HScrollBar HScroll1 
            Height          =   615
            Left            =   240
            TabIndex        =   47
            Top             =   480
            Width           =   3495
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "字"
         Height          =   1335
         Left            =   2760
         TabIndex        =   45
         Top             =   600
         Width           =   3975
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   48
            Text            =   "Better"
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "重绘"
         Height          =   975
         Left            =   2640
         TabIndex        =   44
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Frame Frame16 
         Caption         =   "样式"
         Height          =   1815
         Left            =   600
         TabIndex        =   41
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Option6 
            Caption         =   "字符"
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "线"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "线条函数"
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Command3 
         Caption         =   "清除曲线"
         Height          =   735
         Left            =   4320
         TabIndex        =   39
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Timer Timer3 
         Interval        =   100
         Left            =   7680
         Top             =   4200
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   6720
         Top             =   4680
      End
      Begin VB.CommandButton Command4 
         Caption         =   "绘制曲线"
         Height          =   735
         Left            =   1680
         TabIndex        =   38
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Frame Frame12 
         Caption         =   "函数的参数方程(参数为t)"
         Height          =   3615
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   5655
         Begin VB.TextBox y 
            Height          =   375
            Left            =   840
            TabIndex        =   37
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox x 
            Height          =   375
            Left            =   840
            TabIndex        =   36
            Top             =   360
            Width           =   3255
         End
         Begin VB.Frame Frame13 
            Caption         =   "t的取值范围"
            Height          =   1815
            Left            =   240
            TabIndex        =   25
            Top             =   1560
            Width           =   5175
            Begin VB.Frame Frame15 
               Caption         =   "自定义范围"
               Height          =   1455
               Left            =   2520
               TabIndex        =   27
               Top             =   240
               Width           =   2535
               Begin VB.TextBox endup 
                  Height          =   375
                  Left            =   840
                  TabIndex        =   31
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.TextBox begin 
                  Height          =   375
                  Left            =   840
                  TabIndex        =   30
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "终止值："
                  Height          =   180
                  Left            =   120
                  TabIndex        =   33
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "起始值："
                  Height          =   180
                  Left            =   120
                  TabIndex        =   32
                  Top             =   480
                  Width           =   720
               End
            End
            Begin VB.Frame Frame14 
               Caption         =   "快速选择"
               Height          =   1455
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   2175
               Begin VB.OptionButton Option4 
                  Caption         =   "自由选择"
                  Height          =   375
                  Left            =   360
                  TabIndex        =   29
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.OptionButton Option3 
                  Caption         =   "横坐标范围"
                  Height          =   375
                  Left            =   360
                  TabIndex        =   28
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1335
               End
            End
         End
         Begin VB.Label Label10 
            Caption         =   "y="
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "x="
            Height          =   255
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "曲线函数"
         Height          =   3615
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1815
         Begin VB.ListBox List1 
            Columns         =   1
            Height          =   2580
            ItemData        =   "Form1.frx":80B2E
            Left            =   120
            List            =   "Form1.frx":80B30
            TabIndex        =   40
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "坐标轴设定"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Command2 
         Caption         =   "清除坐标轴"
         Height          =   735
         Left            =   4320
         TabIndex        =   12
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "绘制坐标轴"
         Height          =   735
         Left            =   1680
         TabIndex        =   11
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Frame Frame8 
         Caption         =   "坐标系尺寸及方向设定"
         Height          =   3495
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   4095
         Begin VB.Frame Frame10 
            Caption         =   "窗体右下角横坐标以及纵坐标"
            Height          =   1215
            Left            =   240
            TabIndex        =   16
            Top             =   2040
            Width           =   3615
            Begin VB.TextBox ytop 
               Height          =   270
               Left            =   1680
               TabIndex        =   21
               Text            =   "8"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox ybottom 
               Height          =   270
               Left            =   1680
               TabIndex        =   17
               Text            =   "-8"
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "窗体上纵坐标："
               Height          =   180
               Left            =   360
               TabIndex        =   20
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "窗体下纵坐标："
               Height          =   180
               Left            =   360
               TabIndex        =   18
               Top             =   720
               Width           =   1260
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "窗体左上角以及右下角坐标"
            Height          =   1215
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   3615
            Begin VB.TextBox xright 
               Height          =   270
               Left            =   1680
               TabIndex        =   22
               Text            =   "8"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox xleft 
               Height          =   270
               Left            =   1680
               TabIndex        =   14
               Text            =   "-8"
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "窗体右横坐标："
               Height          =   255
               Left            =   360
               TabIndex        =   19
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "窗体左横坐标："
               Height          =   255
               Left            =   360
               TabIndex        =   15
               Top             =   360
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "目标窗体属性"
         Height          =   1575
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   3375
         Begin VB.CheckBox Check1 
            Caption         =   "绘制坐标刻度"
            Height          =   180
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "坐标轴定义方式"
         Height          =   1575
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "窗体左上角以及右下角坐标"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton Option2 
            Caption         =   "窗体左上角和窗体矢量尺寸"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   960
            Width           =   2895
         End
      End
   End
   Begin VB.Menu setting 
      Caption         =   "坐标轴设定"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
   End
   Begin VB.Menu line 
      Caption         =   "线条函数"
   End
   Begin VB.Menu style 
      Caption         =   "线条样式"
   End
   Begin VB.Menu color 
      Caption         =   "画笔颜色及绘制频度"
   End
   Begin VB.Menu about 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public choose As Integer
Public first As Integer
Private Sub about_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End Sub
Private Sub Check1_Click()
If Check1.Value = 0 Then '是否显示坐标刻度
paint = 0
Else
paint = 1
End If
Call Command1_Click
End Sub
Private Sub color_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
End Sub
Private Sub Command1_Click()
Form2.Cls
paintcoordstatus = True
Timer3.Enabled = True '绘制坐标轴
Call Form2.PaintCoord
If functionstatus = True Then
Call Command4_Click
End If
End Sub

Private Sub Command2_Click()
paintcoordstatus = False '清除坐标系
Form2.Cls
If functionstatus = True Then
Call Command4_Click
End If
End Sub

Private Sub Command3_Click()
functionstatus = False '清除曲线
Form2.Cls
If paintcoordstatus = True Then
Call Command1_Click
End If
End Sub

Private Sub Command4_Click()
functionstatus = True '绘制曲线
Call Form2.reference
Select Case choose
Case 1: Call Form2.sinx
Case 2: Call Form2.cosx
Case 3: Call Form2.tanx
Case 4: Call Form2.kx
Case 5: Call Form2.kxx
Case 6: Call Form2.ajimi
End Select
Timer1.Enabled = True
first = 0
End Sub

Private Sub Command5_Click()
Call Command4_Click
End Sub

Private Sub Command6_Click()
Call Command5_Click
End Sub

Private Sub Form_Load()
Frame1.Visible = True '坐标轴设定
Frame2.Visible = False '线条函数
Frame3.Visible = False '线条样式
Frame4.Visible = False '画笔颜色及绘制频度
Frame5.Visible = False '泡妹模式
Timer3.Enabled = True '实时刷新坐标轴
Form2.Show '绘图窗口
paint = 1 '是否绘制坐标刻度 1为是 0为不是
kind = 1 'kind是指坐标轴的定义方式 1为窗体左上角以及右下角的坐标 2为窗体左上角和窗体矢量尺寸
paintcoordstatus = False '坐标轴绘制状态
functionstatus = False '曲线绘制状态
Timer1.Enabled = True
first = 0
HScroll1.Value = 50
linestyle = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("是否退出本程序？", vbQuestion + vbOKCancel, "信息提示") = vbCancel Then
Cancel = True
Else
Unload Form2
End If
End Sub

Private Sub line_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub
Private Sub Option1_Click()
kind = 1
End Sub

Private Sub Option2_Click()
kind = 2
End Sub

Private Sub setting_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub style_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Timer1_Timer()
List1.Clear '初始化函数
List1.AddItem "正弦函数", 0
List1.AddItem "余弦函数", 1
List1.AddItem "正切函数", 2
List1.AddItem "直线函数ax+b", 3
List1.AddItem "抛物线ax^2+b", 4
List1.AddItem "阿基米德螺线", 5
Label9.Caption = "x="
Label10.Caption = "y="
x.Text = ""
y.Text = ""
Option3.Value = True
Option4.Value = False
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer() '实时更新参数方程取值范围
If List1.ListIndex = 0 Then
x.Text = "i"
y.Text = "Sin(i)"
choose = 1
End If
If List1.ListIndex = 1 Then
x.Text = "i"
y.Text = "Cos(i)"
choose = 2
End If
If List1.ListIndex = 2 Then
x.Text = "i"
y.Text = "Tan(i)"
choose = 3
End If
If List1.ListIndex = 3 Then
Label9.Caption = "a="
Label10.Caption = "b="
If first = 0 Then
x.Text = "1"
y.Text = "0"
End If
first = 1
choose = 4
End If
If List1.ListIndex = 4 Then
Label9.Caption = "a="
Label10.Caption = "b="
If first = 0 Then
x.Text = "1"
y.Text = "0"
End If
first = 1
choose = 5
End If
If List1.ListIndex = 5 Then
x.Text = "1/8*i*cos(i)"
y.Text = "1/8*i*sin(i)"
Option3.Value = False
Option4.Value = True
begin.Text = 0
endup.Text = 80
choose = 6
End If
If Option3.Value = True Then
begin.Text = xleft1
endup.Text = xright1
begin.Enabled = False
endup.Enabled = False
End If
begin.Enabled = True
endup.Enabled = True
ibegin = Val(begin.Text)
iendup = Val(endup.Text)
End Sub

Private Sub Timer3_Timer() '实时更新坐标信息
xleft1 = Val(xleft.Text)
xright1 = Val(xright.Text)
ytop1 = Val(ytop.Text)
ybottom1 = Val(ybottom.Text)
End Sub

Private Sub Timer4_Timer()
linefonts = Text1.Text
If Option5.Value = True Then
linestyle = 1
HScroll1.Max = 100: HScroll1.Min = 0
HScroll1.LargeChange = 10: HScroll1.SmallChange = 1
linestep = HScroll1.Value / 1000
End If
If Option6.Value = True Then
Form2.ForeColor = RGB(rred, ggreen, bblue)
linestyle = 2
HScroll1.Max = 100: HScroll1.Min = 0
HScroll1.LargeChange = 10: HScroll1.SmallChange = 1
linestep = HScroll1.Value / 100
End If
End Sub

Private Sub Timer5_Timer()
rr.Max = 255: rr.Min = 0
rr.LargeChange = 20: rr.SmallChange = 10
rred = rr.Value

gg.Max = 255: gg.Min = 0
gg.LargeChange = 20: gg.SmallChange = 10
ggreen = gg.Value

bb.Max = 255: bb.Min = 0
bb.LargeChange = 20: bb.SmallChange = 10
bblue = bb.Value
Label5.BackColor = RGB(rred, ggreen, bblue)

Label11.BackColor = vbRed
Label12.BackColor = vbGreen
Label13.BackColor = vbBlue
End Sub
