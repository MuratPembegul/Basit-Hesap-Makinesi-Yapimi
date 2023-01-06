Dim x As Single
Dim y As Single
Dim z As Double


toplama kodu

Private Sub Command1_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x + y
Text3.Text = z
End Sub



çıkartma kodu

Private Sub Command2_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x - y
Text3.Text = z
End Sub


çarpma kodu

Private Sub Command3_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x * y
Text3.Text = z
End Sub


bölme kodu

Private Sub Command4_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x / y
Text3.Text = z
End Sub


temizle kodu

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub