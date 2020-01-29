Imports osuCalculator.AccCalc
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hit300, hit100, hit50, miss, topCalc, bottomCalc, result As Double
        Dim accuracy As AccCalc = New AccCalc

        If TextBox1.Text = "" Then
            hit300 = 0
        Else
            hit300 = CDbl(TextBox1.Text)
        End If

        If TextBox2.Text = "" Then
            hit100 = 0
        Else
            hit100 = CDbl(TextBox2.Text)
        End If

        If TextBox3.Text = "" Then
            hit50 = 0
        Else
            hit50 = CDbl(TextBox3.Text)
        End If
        If TextBox4.Text = "" Then
            miss = 0
        Else
            miss = CDbl(TextBox4.Text)
        End If

        topCalc = accuracy.TopCalculation(hit300, hit100, hit50)
        bottomCalc = accuracy.BottomCalculation(hit300, hit100, hit50, miss)

        If bottomCalc = 0 Then
            bottomCalc = 1
        End If


        result = topCalc / bottomCalc

        TextBox5.Clear()
        TextBox5.Text = Math.Round(CDbl(result * 100), 2) & "%"

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()

    End Sub
End Class
