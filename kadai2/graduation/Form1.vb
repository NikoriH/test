Public Class Form1
    Private Sub lstFileName_DragEnter(sender As Object, e As DragEventArgs) Handles lstFileName.DragEnter

        'ドラッグされているものがファイルであるか確認します。
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'ファイルであればコピー可能であることを宣言します。
            e.Effect = DragDropEffects.Copy
        End If

    End Sub

    Private Sub lstFileName_DragDrop(sender As Object, e As DragEventArgs) Handles lstFileName.DragDrop

        'ドロップされたファイルのフルパスを取得します。
        Dim fileName As String
        fileName = CType(e.Data.GetData(DataFormats.FileDrop), String())(0)

        'フルパスからファイル情報を生成してlstFileNameに格納します。
        lstFileName.Items.Add(New IO.FileInfo(fileName))

        '保ボタンを有効化します。
        btnSave.Enabled = True

    End Sub

    Private Sub lstFileName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFileName.SelectedIndexChanged

        Dim file As IO.FileInfo = DirectCast(lstFileName.SelectedItem, IO.FileInfo)
        WebBrowser1.Navigate(file.FullName)
        Me.Text = Application.ProductName & " - " & file.FullName

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        '▼保存するファイルの指定
        Dim dialog As New SaveFileDialog
        dialog.DefaultExt = ".txt" '既定の拡張子を .txt とします。

        'ファイルを保存するダイアログを開き、閉じられるまで待ちます。OKが押されたか確認します。
        If dialog.ShowDialog <> DialogResult.OK Then
            'OK以外で閉じられた場合何もしません。
            Return
        End If

        '▼ファイルへの書き込み
        Using writer As New IO.StreamWriter(dialog.FileName)

            'lstFileNameに格納されているファイル情報を１つずつ取り出します。
            For Each fileInfo As IO.FileInfo In lstFileName.Items
                'ファイル情報からフルパスに書き込みます。
                writer.WriteLine(fileInfo.FullName)
            Next

        End Using

    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click

        '▼開くファイルの指定
        Dim dialog As New OpenFileDialog

        'ファイルを開くダイアログを開き、閉じられるまで待ちます。OKが押されたか確認します。
        If dialog.ShowDialog <> DialogResult.OK Then
            'OK以外で閉じられた場合何もしません。
            Return
        End If

        '▼ファイルからの読み込み
        'lstFileNameに格納している内容をすべてクリアします。
        lstFileName.Items.Clear()

        '指定されたファイルの内容を１行ずつ処理します。
        For Each line As String In IO.File.ReadAllLines(dialog.FileName)
            'フルパスからファイル情報を生成してlstFileNameに格納します。
            lstFileName.Items.Add(New IO.FileInfo(line))
        Next

        '保ボタンを有効化します。
        btnSave.Enabled = True

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.Text = Application.ProductName

    End Sub

End Class
