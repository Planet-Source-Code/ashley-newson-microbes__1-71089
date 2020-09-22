VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "KILL"
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Resize Form/Magnification if Magnification/Form Size Changes"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Rolling Sugar Plant Seed"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Rolling Sugar"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Sugar Plant Thing"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   1575
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Wall"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Fungi"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Spore"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Virus"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Text            =   "8"
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Corpse"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "0"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Sugar"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Bacteria"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Value           =   -1  'True
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      Min             =   1
      Max             =   500
      SelStart        =   50
      Value           =   50
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      Min             =   1
      Max             =   500
      SelStart        =   50
      Value           =   50
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Infected Bacterium"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Magnification"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Hz"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "50"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "50"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Y Size"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "X Size"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sad = 0
If Option2.Value = True Then sad = 1
If Option3.Value = True Then sad = 2
If Option4.Value = True Then sad = 3
If Option5.Value = True Then sad = 4
If Option6.Value = True Then sad = 5
If Option7.Value = True Then sad = 6
If Option8.Value = True Then sad = 7
If Option9.Value = True Then sad = 255
If Option10.Value = True Then sad = 8
If Option11.Value = True Then sad = 9
If Option12.Value = True Then sad = 10
If sad <> 0 Then
For vy = 1 To Form2.Slider2.Value
For vx = 1 To Form2.Slider1.Value
If maptypen(vx, vy) = sad Then
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
maptype(vx, vy) = 0
map(vx, vy) = 0
End If
Next
Next
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Form2.WindowState = 1
End Sub

Private Sub Slider1_Change()
Label3.Caption = Slider1.Value
mag = Val(Text3.Text)
Form1.Picture1.Width = Slider1.Value * mag * 15
Form1.Picture1.Height = Slider2.Value * mag * 15
If Form1.WindowState = 0 Then
If Check1.Value = 1 Then
Form1.Width = (Slider1.Value * mag * 15) + 120
Form1.Height = (Slider2.Value * mag * 15) + 810
End If
Form1.Cls
End If
End Sub

Private Sub Slider2_Change()
Label4.Caption = Slider2.Value
mag = Val(Text3.Text)
Form1.Picture1.Width = Slider1.Value * mag * 15
Form1.Picture1.Height = Slider2.Value * mag * 15
If Form1.WindowState = 0 Then
If Check1.Value = 1 Then
Form1.Width = (Slider1.Value * mag * 15) + 120
Form1.Height = (Slider2.Value * mag * 15) + 810
End If
Form1.Cls
End If
End Sub

Private Sub Text1_Change()
If Str$(Val(Text2.Text)) = Text2.Text Then
Text1.Text = "0"
MsgBox ("Number Please")
Exit Sub
End If
End Sub

Private Sub text2_Validate(Cancel As Boolean)
If Str$(Val(Text2.Text)) = Text2.Text Then
MsgBox ("Number Please")
Exit Sub
End If
If Val(Text2.Text) < 0 Then
MsgBox ("Positive or Zero Please")
Exit Sub
End If
If Val(Text2.Text) = 0 Then
Form1.Timer1.Enabled = False
Exit Sub
End If
Form1.Timer1.Enabled = True
Form1.Timer1.Interval = Int(1000 / Val(Text2.Text))
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Str$(Val(Text3.Text)) = Text3.Text Then
MsgBox ("Zero Or Positive Number Please")
Text3.Text = "1"
Exit Sub
End If
If Val(Text3.Text) < 0 Then
MsgBox ("Positive or Zero Please")
Text3.Text = "1"
Exit Sub
End If
mag = Val(Text3.Text)
Form1.Picture1.Width = Slider1.Value * mag * 15
Form1.Picture1.Height = Slider2.Value * mag * 15
If Form1.WindowState = 0 Then
If Check1.Value = 1 Then
Form1.Width = (Slider1.Value * mag * 15) + 120
Form1.Height = (Slider2.Value * mag * 15) + 810
End If
Form1.Cls
End If
End Sub
