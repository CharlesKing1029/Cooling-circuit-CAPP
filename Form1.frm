VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "冷却水道的尺寸确定"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   9000
      TabIndex        =   6
      Top             =   8760
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   1095
      Left            =   14400
      TabIndex        =   5
      Text            =   "模具成型表面温度"
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   8760
      TabIndex        =   4
      Text            =   "冷却水流速 单位m/s"
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   14280
      TabIndex        =   3
      Text            =   "单位时间注入模具中树脂的质量 单位kg/h"
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   8760
      TabIndex        =   2
      Text            =   "塑件壁厚 单位mm"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Text            =   "树脂种类"
      Top             =   4560
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FORM_LOAD()
Combo1.AddItem ("ABS")
Combo1.AddItem ("PVC")
Combo1.AddItem ("PP")
Combo1.AddItem ("PE")
Combo1.AddItem ("PC")
End Sub
Private Sub command1_click()
If Combo1.Text = "ABS" Then q = 3.5
If Combo1.Text = "PVC" Then q = 2.65
If Combo1.Text = "PP" Then q = 5.9
If Combo1.Text = "PE" Then q = 6.9
If Combo1.Text = "PC" Then q = 2.9
e = Text1.Text
v = Text3.Text
m = Text2.Text
t = Text4.Text
Select Case e
 Case e = 2
 d = 9
 Case e >= 2 And e < 4
 d = 11
 Case e >= 4 And e <= 6
 d = 12
 Case e < 2
 MsgBox ("壁厚过小，会造成充填阻力增大，复杂制件将难于成型，请重新设置壁厚")
 Case e > 6
 d = 14
End Select
a = 7.5 * (1000 * v) ^ 0.8 / d ^ 0.2
s = q * m / (3600 * a * (t - 20))
l = a / (3.14 * d)
Text5.Text = l
End Sub
