Public Class Form1
    Dim random As New Random
    Dim qnum As String
    Dim ansnum As Integer
    Dim hit As Integer
    Dim blow As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        qnum = "0000"
        ansnum = 0

        '回答作成
        While (qnum.Substring(0, 1) = qnum.Substring(1, 1) Or
               qnum.Substring(0, 1) = qnum.Substring(2, 1) Or
               qnum.Substring(0, 1) = qnum.Substring(3, 1) Or
               qnum.Substring(1, 1) = qnum.Substring(2, 1) Or
               qnum.Substring(1, 1) = qnum.Substring(3, 1) Or
               qnum.Substring(2, 1) = qnum.Substring(3, 1))

            qnum = random.Next(1, 9999).ToString("0000")

        End While
        Debug.Print(qnum)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If 0 = Inputcheck(TextBox1.Text) Then

            Call Match(TextBox1.Text)

            ansnum += 1
            Label11.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If 0 = Inputcheck(TextBox2.Text) Then

            Call Match(TextBox2.Text)

            ansnum += 1
            Label12.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If 0 = Inputcheck(TextBox3.Text) Then

            Call Match(TextBox3.Text)

            ansnum += 1
            Label13.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If 0 = Inputcheck(TextBox4.Text) Then

            Call Match(TextBox4.Text)

            ansnum += 1
            Label14.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If 0 = Inputcheck(TextBox5.Text) Then

            Call Match(TextBox5.Text)

            ansnum += 1
            Label15.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If 0 = Inputcheck(TextBox6.Text) Then

            Call Match(TextBox6.Text)

            ansnum += 1
            Label16.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If 0 = Inputcheck(TextBox7.Text) Then

            Call Match(TextBox7.Text)

            ansnum += 1
            Label17.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If 0 = Inputcheck(TextBox8.Text) Then

            Call Match(TextBox8.Text)

            ansnum += 1
            Label18.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If 0 = Inputcheck(TextBox9.Text) Then

            Call Match(TextBox9.Text)

            ansnum += 1
            Label19.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If 0 = Inputcheck(TextBox10.Text) Then

            Call Match(TextBox10.Text)

            ansnum += 1
            Label20.Text = "hit:" & hit & "     " & "blow:" & blow

            Call Judge()

        End If

    End Sub

    Private Sub Match(ByRef num As String)

        Dim i As Integer
        Dim j As Integer

        hit = 0
        blow = 0
        For i = 0 To 3
            For j = 0 To 3
                If num.Substring(i, 1) = qnum.Substring(j, 1) Then
                    If i = j Then
                        hit += 1
                    Else
                        blow += 1
                    End If
                End If
            Next
        Next

    End Sub

    Private Function Inputcheck(ByRef indata As String) As Integer

        Dim a As Integer

        a = 0

        If indata = "" Then
            a = MsgBox("数字が未入力です", vbCritical)
        ElseIf 4 <> indata.Length Then
            a = MsgBox("4桁の数字を入力してください", vbCritical)
        Else
        End If

        Inputcheck = a

    End Function

    Private Sub Judge()

        Dim a As Integer

        If hit = 4 Then
            a = MsgBox("正解！", vbOKOnly)
        Else
            If ansnum >= 10 Then
                a = MsgBox("ゲームオーバー", vbCritical)
            End If
        End If

    End Sub

End Class
