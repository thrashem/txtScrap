Imports System.IO
Imports System.Windows.Forms

Module Module1
    <STAThread>
    Sub Main()
        Try
            ' クリップボードにテキストがあるか確認
            If Not My.Computer.Clipboard.ContainsText() Then
                ' Console.WriteLine("クリップボードにテキストがありません")
                Exit Sub
            End If

            ' クリップボードの内容を取得し設定
            Dim clipboardText As String = My.Computer.Clipboard.GetText()
            My.Computer.Clipboard.SetText(clipboardText, TextDataFormat.Text)

            ' 現在時刻とファイルパスを設定
            Dim dt As DateTime = DateTime.Now
            Dim txtScrapDirectory As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "txtScrap")
            Dim txtScrapFile As String = Path.Combine(txtScrapDirectory, dt.ToString("yyyyMMdd") & ".txt")
            Dim utf8 As System.Text.Encoding = System.Text.Encoding.UTF8

            ' ディレクトリが存在しない場合は作成
            If Not Directory.Exists(txtScrapDirectory) Then
                Directory.CreateDirectory(txtScrapDirectory)
            End If

            ' ファイルに書き込み
            Using sw As New StreamWriter(txtScrapFile, True, utf8)
                sw.Write(dt.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf & clipboardText & vbCrLf)
            End Using

        Catch ioEx As IOException
            Console.WriteLine("ファイル操作中にエラーが発生しました: " & ioEx.Message)
        Catch formatEx As FormatException
            Console.WriteLine("データ形式エラーが発生しました: " & formatEx.Message)
        Catch ex As Exception
            Console.WriteLine("予期しないエラーが発生しました: " & ex.Message)
        End Try
    End Sub
End Module
