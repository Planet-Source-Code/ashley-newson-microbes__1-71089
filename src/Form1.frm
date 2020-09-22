VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Microbes"
   ClientHeight    =   6000
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   6000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmap(*.bmp)|*.bmp|All Files(*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Microbes Save File(*.msf)|*.msf|All Files(*.*)|*.*"
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Menu save 
      Caption         =   "Save"
   End
   Begin VB.Menu open 
      Caption         =   "Open"
   End
   Begin VB.Menu screenshot 
      Caption         =   "ScreenShot"
   End
   Begin VB.Menu License 
      Caption         =   "License"
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
MsgBox ("Microbes - Version 1.2.2" + Chr$(13) + "Original Program by Ashley Newson." + Chr$(13) + "This adaption by Ashley Newson." + Chr$(13) + "To see the license agreement click the 'License' button on the menu.")
End Sub

Private Sub Form_Load()
On Error GoTo errorform1load
If 0 = 1 Then
errorform1load:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
If Command$ <> "" Then
Dim mx As Integer
Dim my As Integer
fto$ = Mid$(Command$, 2, Len(Command$) - 2)
Open fto$ For Random As #1
Form1.Caption = "Microbes - " + fto$
Get #1, 1, mx
Get #1, 2, my
Form2.Slider1.Value = mx
Form2.Slider2.Value = my
For vy = 1 To Form2.Slider2.Value
For vx = 1 To Form2.Slider1.Value
Get #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 1, maptype(vx, vy)
Get #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 2, map(vx, vy)
maptypen(vx, vy) = maptype(vx, vy)
mapn(vx, vy) = map(vx, vy)
Next
Next
Close #1
End If
End Sub

Private Sub Form_Resize()
On Error GoTo errorform1resize
If 0 = 1 Then
errorform1resize:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
If Form1.WindowState <> 1 Then
If Form2.Check1.Value = 1 Then
If (Form1.Width - 120) / Picture1.Width > (Form1.Height - 810) / Picture1.Height Then
Form2.Text3.Text = (Form1.Height - 810) / Form2.Slider2.Value / 15
Else
Form2.Text3.Text = (Form1.Width - 120) / Form2.Slider1.Value / 15
End If
mag = Form2.Text3.Text
Form1.Picture1.Width = Form2.Slider1.Value * mag * 15
Form1.Picture1.Height = Form2.Slider2.Value * mag * 15
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub License_Click()
MsgBox ("Program: Microbes" + Chr$(13) + "Version: 1.2.2" + Chr$(13) + "Microbes (C) Ashley Newson 2008" + Chr$(13) + "By ditributing this program, using the program and/or changing the source code of this program you are agreeing to this license agreement:" + Chr$(13) + "1. You are not allowed to sell this software." + Chr$(13) + "2. You are allowed to distribute this software FOR FREE ONLY." + Chr$(13) + "3. You are allow to edit the source code, HOWEVER you must make clear the original author (Ashley Newson) and you must release the source code under the same license." + Chr$(13) + "4. You may use this source code as an example!" + Chr$(13) + "5. If you wish to get extra permissions please e-mail the original author at master0060@hotmail.co.uk" + Chr$(13) + "6. The original author cannot be harmed by legal damage if the original or an adaptation does something bad or incorrect!")
End Sub

Private Sub open_Click()
On Error GoTo erroropenclick
If 0 = 1 Then
erroropenclick:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
Dim mx As Integer
Dim my As Integer
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Random As #1
Form1.Caption = "Microbes - " + CommonDialog1.FileName
Get #1, 1, mx
Get #1, 2, my
Form2.Slider1.Value = mx
Form2.Slider2.Value = my
For vy = 1 To Form2.Slider2.Value
For vx = 1 To Form2.Slider1.Value
Get #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 1, maptype(vx, vy)
Get #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 2, map(vx, vy)
maptypen(vx, vy) = maptype(vx, vy)
mapn(vx, vy) = map(vx, vy)
Next
Next
Close #1
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorpicmd
If 0 = 1 Then
errorpicmd:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
mag = Val(Form2.Text3.Text)
vx = Int((X / 15 / mag) + 0.999)
vy = Int((Y / 15 / mag) + 0.999)
Form2.Label6.Caption = Str$(vx) + Str$(vy)
If Button = 2 And vx <= 500 And vy <= 500 And vx > 0 And vy > 0 Then
Form2.Show
Form2.Label6.Caption = Str$(map(vx, vy)) + Str$(maptype(vx, vy)) + Str$(vx) + Str$(vy)
End If
If Button = 1 And vx <= Form2.Slider1.Value And vy <= Form2.Slider2.Value And vx > 0 And vy > 0 Then
If Form2.Option1.Value = True Then
maptype(vx, vy) = 0
maptypen(vx, vy) = 0
map(vx, vy) = 0
mapn(vx, vy) = 0
End If
If Form2.Option2.Value = True Then
maptype(vx, vy) = 1
maptypen(vx, vy) = 1
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option3.Value = True Then
maptype(vx, vy) = 2
maptypen(vx, vy) = 2
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option4.Value = True Then
maptype(vx, vy) = 3
maptypen(vx, vy) = 3
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option5.Value = True Then
maptype(vx, vy) = 4
maptypen(vx, vy) = 4
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option6.Value = True Then
maptype(vx, vy) = 5
maptypen(vx, vy) = 5
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option7.Value = True Then
maptype(vx, vy) = 6
maptypen(vx, vy) = 6
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option8.Value = True Then
maptype(vx, vy) = 7
maptypen(vx, vy) = 7
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option9.Value = True Then
maptype(vx, vy) = 255
maptypen(vx, vy) = 255
map(vx, vy) = 0
mapn(vx, vy) = 0
End If
If Form2.Option10.Value = True Then
maptype(vx, vy) = 8
maptypen(vx, vy) = 8
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option11.Value = True Then
maptype(vx, vy) = 9
maptypen(vx, vy) = 9
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If Form2.Option12.Value = True Then
maptype(vx, vy) = 10
maptypen(vx, vy) = 10
map(vx, vy) = Val(Form2.Text1.Text)
mapn(vx, vy) = Val(Form2.Text1.Text)
End If
If maptypen(vx, vy) = 0 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 0), BF
If maptypen(vx, vy) = 255 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 255), BF
If maptypen(vx, vy) = 1 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 255, 0), BF
If maptypen(vx, vy) = 2 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 128, 0), BF
If maptypen(vx, vy) = 3 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 128, 128), BF
If maptypen(vx, vy) = 4 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 255), BF
If maptypen(vx, vy) = 5 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 128, 128), BF
If maptypen(vx, vy) = 6 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 255, 128), BF
If maptypen(vx, vy) = 7 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 64, 0), BF
If maptypen(vx, vy) = 8 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 0, 0), BF
If maptypen(vx, vy) = 9 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 0), BF
If maptypen(vx, vy) = 10 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 0, 0), BF
End If
End Sub

Private Sub save_Click()
On Error GoTo errorsaveclick
If 0 = 1 Then
errorsaveclick:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
Dim mx As Integer
Dim my As Integer
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Random As #1
Form1.Caption = "Microbes - " + CommonDialog1.FileName
mx = Form2.Slider1.Value
my = Form2.Slider2.Value
Put #1, 1, mx
Put #1, 2, my
For vy = 1 To Form2.Slider2.Value
For vx = 1 To Form2.Slider1.Value
Put #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 1, maptype(vx, vy)
Put #1, ((((vy * Form2.Slider1) + vx) - Form2.Slider1) * 2) + 2, map(vx, vy)
Next
Next
Close #1
End If
End Sub

Private Sub screenshot_Click()
On Error GoTo errorssclick
If 0 = 1 Then
errorssclick:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to resume.")
Exit Sub
End If
Picture1.Picture = LoadPicture()
If Val(Form2.Text3.Text) > 0 Then
mag = Val(Form2.Text3.Text)
For vy = 1 To Form2.Slider2
For vx = 1 To Form2.Slider1
If maptypen(vx, vy) = 0 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 0), BF
If maptypen(vx, vy) = 255 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 255), BF
If maptypen(vx, vy) = 1 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 255, 0), BF
If maptypen(vx, vy) = 2 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 128, 0), BF
If maptypen(vx, vy) = 3 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 128, 128), BF
If maptypen(vx, vy) = 4 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 255), BF
If maptypen(vx, vy) = 5 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 128, 128), BF
If maptypen(vx, vy) = 6 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 255, 128), BF
If maptypen(vx, vy) = 7 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 64, 0), BF
If maptypen(vx, vy) = 8 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 0, 0), BF
If maptypen(vx, vy) = 9 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 0), BF
If maptypen(vx, vy) = 10 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 0, 0), BF
Next
Next
End If
CommonDialog2.FileName = ""
CommonDialog2.ShowSave
If CommonDialog2.FileName <> "" Then SavePicture Picture1.Image, CommonDialog2.FileName
Picture1.Picture = LoadPicture()
End Sub

Private Sub Timer1_Timer()
On Error GoTo errorrun
If 0 = 1 Then
errorrun:
MsgBox ("Error:" + Chr$(13) + "Error Code:" + Str$(Err.Number) + Chr$(13) + "Error Description:" + Err.Description + Chr$(13) + Chr$(13) + "!Aborting this operation!" + Chr$(13) + "Press OK to pause the simulation and return to the option screen.")
Timer1.Enabled = False
Exit Sub
End If
Dim moveme As Boolean
For vy = 1 To Form2.Slider2
For vx = 1 To Form2.Slider1
moveme = True
'===================================Bacterium======================================
If maptype(vx, vy) = 1 Then
If map(vx, vy) <= 10 Then
maptypen(vx, vy) = 3
Else
For vy0 = -1 To 1
For vx0 = -1 To 1
If (maptypen(vx + vx0, vy + vy0) = 2 Or maptypen(vx + vx0, vy + vy0) = 3) And (vx0 <> 0 Or vy0 <> 0) Then
If mapn(vx + vx0, vy + vy0) > 10 Then
mapn(vx + vx0, vy + vy0) = mapn(vx + vx0, vy + vy0) - 10
mapn(vx, vy) = mapn(vx, vy) + 10
moveme = False
End If
If mapn(vx + vx0, vy + vy0) <= 10 Then
mapn(vx, vy) = mapn(vx, vy) + mapn(vx + vx0, vy + vy0)
mapn(vx + vx0, vy + vy0) = 0
maptypen(vx + vx0, vy + vy0) = 0
End If
End If
If (maptypen(vx + vx0, vy + vy0) = 7) And (vx0 <> 0 Or vy0 <> 0) Then
If mapn(vx + vx0, vy + vy0) > 1 Then
mapn(vx + vx0, vy + vy0) = mapn(vx + vx0, vy + vy0) - 1
mapn(vx, vy) = mapn(vx, vy) + 1
moveme = False
End If
If mapn(vx + vx0, vy + vy0) = 1 Then
mapn(vx, vy) = mapn(vx, vy) + 1
mapn(vx + vx0, vy + vy0) = 0
maptypen(vx + vx0, vy + vy0) = 0
End If
End If
Next
Next
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If maptypen(vx + vx0, vy + vy0) = 0 Then
If mapn(vx, vy) < 300 Then
If moveme = True Then
mapn(vx + vx0, vy + vy0) = mapn(vx, vy) - 1
mapn(vx, vy) = 0
maptypen(vx + vx0, vy + vy0) = 1
maptypen(vx, vy) = 0
End If
Else
mapn(vx + vx0, vy + vy0) = mapn(vx, vy) / 2
mapn(vx, vy) = mapn(vx, vy) / 2
maptypen(vx + vx0, vy + vy0) = 1
End If
End If
End If
End If
'======================================Virri=======================================
If maptype(vx, vy) = 4 Then
If map(vx, vy) < 1 Then
maptypen(vx, vy) = 0
Else
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If vx0 <> 0 Or vy0 <> 0 Then
If maptypen(vx + vx0, vy + vy0) = 1 Or maptypen(vx + vx0, vy + vy0) = 5 Then
maptypen(vx + vx0, vy + vy0) = 5
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
GoTo endpartv
End If
If maptypen(vx + vx0, vy + vy0) = 0 Then
maptypen(vx + vx0, vy + vy0) = 4
mapn(vx + vx0, vy + vy0) = mapn(vx, vy) - 1
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
End If
End If
endpartv:
End If
End If
'===============================Infected Bacterium=================================
If maptype(vx, vy) = 5 Then
If map(vx, vy) <= 10 Then
maptypen(vx, vy) = 4
Else
For vy0 = -1 To 1
For vx0 = -1 To 1
If (maptypen(vx + vx0, vy + vy0) = 2 Or maptypen(vx + vx0, vy + vy0) = 3 Or maptypen(vx + vx0, vy + vy0) = 7) And (vx0 <> 0 Or vy0 <> 0) Then
If mapn(vx + vx0, vy + vy0) > 10 Then
mapn(vx + vx0, vy + vy0) = mapn(vx + vx0, vy + vy0) - 10
mapn(vx, vy) = mapn(vx, vy) + 10
moveme = False
End If
If mapn(vx + vx0, vy + vy0) <= 10 Then
mapn(vx, vy) = mapn(vx, vy) + mapn(vx + vx0, vy + vy0)
mapn(vx + vx0, vy + vy0) = 0
maptypen(vx + vx0, vy + vy0) = 0
End If
End If
Next
Next
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If maptypen(vx + vx0, vy + vy0) = 0 Then
If mapn(vx, vy) < 600 Then
If moveme = True Then
mapn(vx + vx0, vy + vy0) = mapn(vx, vy) - 1
mapn(vx, vy) = 0
maptypen(vx + vx0, vy + vy0) = 5
maptypen(vx, vy) = 0
End If
Else
For vy1 = -1 To 1
For vx1 = -1 To 1
If maptypen(vx + vx1, vy + vy1) = 0 Then
maptypen(vx + vx1, vy + vy1) = 4
mapn(vx + vx1, vy + vy1) = 15
End If
If maptypen(vx + vx1, vy + vy1) = 1 Then
maptypen(vx + vx1, vy + vy1) = 5
End If
Next
Next
maptypen(vx, vy) = 4
mapn(vx, vy) = 15
End If
End If
End If
End If
'======================================Spore=======================================
If maptype(vx, vy) = 6 Then
If map(vx, vy) < 1 Then
maptypen(vx, vy) = 0
Else
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If vx0 <> 0 Or vy0 <> 0 Then
If maptypen(vx + vx0, vy + vy0) = 2 Or maptypen(vx + vx0, vy + vy0) = 3 Then
maptypen(vx + vx0, vy + vy0) = 7
mapn(vx + vx0, vy + vy0) = mapn(vx + vx0, vy + vy0) + 100
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
GoTo endparts
End If
If maptypen(vx + vx0, vy + vy0) = 0 Then
maptypen(vx + vx0, vy + vy0) = 6
mapn(vx + vx0, vy + vy0) = mapn(vx, vy) - 1
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
End If
End If
endparts:
End If
End If
'=====================================Fungus=======================================
If maptype(vx, vy) = 7 Then
If map(vx, vy) <= 5 Then
If map(vx, vy) <= 0 Then
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
Else
maptypen(vx, vy) = 3
End If
Else
For vy0 = -1 To 1
For vx0 = -1 To 1
If (maptypen(vx + vx0, vy + vy0) = 2 Or maptypen(vx + vx0, vy + vy0) = 3) And (vx0 <> 0 Or vy0 <> 0) Then
If mapn(vx + vx0, vy + vy0) > 100 Then
mapn(vx + vx0, vy + vy0) = mapn(vx + vx0, vy + vy0) - 100
mapn(vx, vy) = mapn(vx, vy) + 100
moveme = False
End If
If mapn(vx + vx0, vy + vy0) <= 100 Then
mapn(vx, vy) = mapn(vx, vy) + mapn(vx + vx0, vy + vy0)
maptypen(vx + vx0, vy + vy0) = 0
End If
End If
Next
Next
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
mapn(vx, vy) = mapn(vx, vy) - 1
If maptypen(vx + vx0, vy + vy0) = 0 Then
If mapn(vx, vy) >= 5000 Then
mapn(vx + vx0, vy + vy0) = 100
mapn(vx, vy) = mapn(vx, vy) - 4000
maptypen(vx + vx0, vy + vy0) = 6
End If
End If
End If
End If
'================================Sugar Plant Thing=================================
If maptype(vx, vy) = 8 Then
If mapn(vx, vy) <= 0 Then
maptypen(vx, vy) = 3
mapn(vx, vy) = 250
Else
For vy0 = -1 To 1
For vx0 = -1 To 1
If (maptypen(vx + vx0, vy + vy0) <> 0) And (vx0 <> 0 Or vy0 <> 0) Then
mapn(vx, vy) = mapn(vx, vy) - 1
End If
Next
Next
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
mapn(vx, vy) = mapn(vx, vy) + 7
If maptypen(vx + vx0, vy + vy0) = 0 Then
If mapn(vx, vy) >= 750 Then
v = Int(Rnd * 20) + 1
If v >= 1 And v <= 18 Then
mapn(vx + vx0, vy + vy0) = 10
mapn(vx, vy) = 1
maptypen(vx + vx0, vy + vy0) = 9
End If
If v >= 18 And v <= 19 Then
mapn(vx + vx0, vy + vy0) = 20
mapn(vx, vy) = 1
maptypen(vx + vx0, vy + vy0) = 10
End If
If v = 20 Then
maptypen(vx, vy) = 3
End If
End If
End If
End If
End If
'==================================Rolling Sugar===================================
If maptype(vx, vy) = 9 Then
mapn(vx, vy) = mapn(vx, vy) - 1
If map(vx, vy) <= 1 Then
maptypen(vx, vy) = 2
mapn(vx, vy) = 10000
Else
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If vx0 <> 0 Or vy0 <> 0 Then
If maptypen(vx + vx0, vy + vy0) = 0 Then
maptypen(vx + vx0, vy + vy0) = 9
mapn(vx + vx0, vy + vy0) = mapn(vx, vy)
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
End If
End If
End If
End If
'===================================Rolling Seed===================================
If maptype(vx, vy) = 10 Then
mapn(vx, vy) = mapn(vx, vy) - 1
If map(vx, vy) <= 1 Then
maptypen(vx, vy) = 8
mapn(vx, vy) = 1
Else
vx0 = Int(Rnd * 3) - 1
vy0 = Int(Rnd * 3) - 1
If vx0 <> 0 Or vy0 <> 0 Then
If maptypen(vx + vx0, vy + vy0) = 0 Then
maptypen(vx + vx0, vy + vy0) = 10
mapn(vx + vx0, vy + vy0) = mapn(vx, vy)
maptypen(vx, vy) = 0
mapn(vx, vy) = 0
End If
End If
End If
End If
'==================================================================================
Next
Next
'==================================FINALISING======================================
For vy = 1 To Form2.Slider2
For vx = 1 To Form2.Slider1
map(vx, vy) = mapn(vx, vy)
maptype(vx, vy) = maptypen(vx, vy)
Next
Next
'====================================DISPLAY=======================================
If Val(Form2.Text3.Text) > 0 Then
mag = Val(Form2.Text3.Text)
For vy = 1 To Form2.Slider2
For vx = 1 To Form2.Slider1
If maptypen(vx, vy) = 0 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 0), BF
If maptypen(vx, vy) = 255 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 255), BF
If maptypen(vx, vy) = 1 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 255, 0), BF
If maptypen(vx, vy) = 2 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 128, 0), BF
If maptypen(vx, vy) = 3 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 128, 128), BF
If maptypen(vx, vy) = 4 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 0, 255), BF
If maptypen(vx, vy) = 5 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(0, 128, 128), BF
If maptypen(vx, vy) = 6 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 255, 128), BF
If maptypen(vx, vy) = 7 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 64, 0), BF
If maptypen(vx, vy) = 8 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 0, 0), BF
If maptypen(vx, vy) = 9 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(255, 255, 0), BF
If maptypen(vx, vy) = 10 Then Picture1.Line (vx * 15 * mag, vy * 15 * mag)-((vx * 15 * mag) - ((mag * 15) - 1), (vy * 15 * mag) - ((mag * 15) - 1)), RGB(128, 0, 0), BF
Next
Next
End If
End Sub
