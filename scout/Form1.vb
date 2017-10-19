Public Class Form1
    Dim people As Integer
    Dim cnt, i As Integer
    Dim Type(50) As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        people = InputBox("請輸入人數", "Hi")
        Randomize()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim num As Integer
        num = Int(Rnd() * people + 1)
        Me.Text = num.ToString
        Select Case num
            Case 1
                Label1.Text = "歡迎來到空軍"
            Case 2 To 4
                Label1.Text = "歡迎來到海軍"
            Case Else
                Label1.Text = "歡迎來到陸軍"
        End Select
    End Sub
End Class
