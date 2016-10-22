Public Class Form1
    Dim V, I, R As Decimal

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles button3.Click
        Close()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        radioButton3.Checked = True
        radioButton5.Checked = True
        radioButton8.Checked = True
        textBox1.Text = ""
        textBox2.Text = ""
        textBox3.Text = ""
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles button2.Click
        If radioButton8.Checked = True Then
            If radioButton3.Checked = True Then
                R = Val(textBox1.Text)
            ElseIf radioButton2.Checked = True Then
                R = Val(textBox1.Text) * 1000
            ElseIf radioButton1.Checked = True Then
                R = Val(textBox1.Text) * 1000000000
            End If
            If radioButton5.Checked = True Then
                I = Val(textBox2.Text) * 0.001
            ElseIf radioButton4.Checked = True Then
                I = Val(textBox2.Text)
            End If
            If R <= 0 Or I <= 0 Then
                MsgBox("Please enter only numric values larger than 0")
                Exit Sub
            End If
            If IsNumeric(textBox1.Text) And IsNumeric(textBox2.Text) Then
                If R > 0 And I > 0 Then
                    V = I * R
                    textBox3.Text = V
                End If
            End If
        End If
        If radioButton7.Checked = True Then
            V = Val(textBox3.Text)
            If radioButton3.Checked = True Then
                R = Val(textBox1.Text)
            ElseIf radioButton2.Checked = True Then
                R = Val(textBox1.Text) * 1000
            ElseIf radioButton1.Checked = True Then
                R = Val(textBox1.Text) * 1000000000
            End If
            If R <= 0 Or V <= 0 Then
                MsgBox("Please enter only numric values larger than 0")
                Exit Sub
            End If
            If IsNumeric(textBox1.Text) And IsNumeric(textBox3.Text) Then
                If R > 0 And V > 0 Then
                    I = V / R
                    If radioButton5.Checked = True Then
                        I = I * 1000
                    End If
                    textBox2.Text = I
                End If
            End If
        End If
        If radioButton6.Checked = True Then
            If radioButton5.Checked = True Then
                I = Val(textBox2.Text) * 0.001
            ElseIf radioButton4.Checked = True Then
                I = Val(textBox2.Text)
            End If
            V = Val(textBox3.Text)
            If V <= 0 Or I <= 0 Then
                MsgBox("Please enter only numric values larger than 0" & V & I)
                Exit Sub
            End If
            If V > 0 And I > 0 Then
                If IsNumeric(textBox2.Text) And IsNumeric(textBox3.Text) Then
                    R = V / I
                    If radioButton2.Checked = True Then
                        R = R * 0.001
                    ElseIf radioButton1.Checked = True Then
                        R = R * 0.000001
                    End If
                    textBox1.Text = R
                End If
            End If
        End If
    End Sub
End Class
