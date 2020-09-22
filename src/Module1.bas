Attribute VB_Name = "SubMain"
Public map(501, 501) As Long
Public maptype(501, 501) As Byte
Public mapn(501, 501) As Long
Public maptypen(501, 501) As Byte

Sub Main()
Randomize Timer
Load Form2
Form2.Show
Load Form1
Form1.Show
End Sub
