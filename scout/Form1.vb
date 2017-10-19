Public Class Form1
    Dim people As Integer
    Dim cnt As Integer
    Dim Type(50) As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        people = InputBox("請輸入人數", "Hi")
        Label2.Text = "參加人數" & people & "人"
        Randomize()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim num As Integer
        For i = 1 To people
            num = Int(Rnd() * people + 1)
            If num = Type(i) Then
                num = Int(Rnd() * people + 1)
            End If
            If num = Type(i) Then
                num = Int(Rnd() * people + 1)
            End If
            If num = Type(i) Then
                num = Int(Rnd() * people + 1)
            End If
            Me.Text = num.ToString
            If Type(num) = 0 Then
                Exit For
            End If
        Next
        cnt += 1
        If cnt <= people Then
            Select Case num / people
                Case 0.0 To 0.1
                    Label1.Text = "歡迎來到空軍"
                    Label3.Text += "空"
                Case 0.11 To 0.4
                    Label1.Text = "歡迎來到海軍"
                    Label3.Text += "海"
                Case Is > 0.4
                    Label1.Text = "歡迎來到陸軍"
                    Label3.Text += "陸"
            End Select

        Else
                MessageBox.Show("抽完啦")
            End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i = 1 To people
            Button1_Click(sender, e)
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Application.Restart()
    End Sub
End Class
