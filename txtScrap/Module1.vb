Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing ' これを追加

Module Module1
    Private notify As NotifyIcon

    <STAThread>
    Sub Main()
        ' NotifyIconの初期化
        notify = New NotifyIcon()
        notify.Icon = System.Drawing.SystemIcons.Information '
        notify.Visible = True

        Try
            ' クリップボードにテキストがあるか確認
            If Not My.Computer.Clipboard.ContainsText() Then
                notify.ShowBalloonTip(3000, "エラー", "クリップボードにテキストがありません", ToolTipIcon.Error)
                Exit Sub
            End If

            ' クリップボードの内容を取得し設定
            Dim clipboardText As String = My.Computer.Clipboard.GetText()
            My.Computer.Clipboard.SetText(clipboardText, TextDataFormat.Text)

            ' 現在時刻とファイルパスを設定
            Dim dt As DateTime = DateTime.Now
            Dim txtScrapDirectory As String = "G:\マイドライブ\txtScrap"
            Dim txtScrapFile As String = Path.Combine(txtScrapDirectory, dt.ToString("yyyyMMdd") & ".md")
            Dim utf8 As System.Text.Encoding = System.Text.Encoding.UTF8

            ' ディレクトリが存在しない場合は作成
            If Not Directory.Exists(txtScrapDirectory) Then
                Directory.CreateDirectory(txtScrapDirectory)
            End If

            ' ファイルに書き込み
            Using sw As New StreamWriter(txtScrapFile, True, utf8)
                sw.Write("#### " & dt.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf & clipboardText & vbCrLf & "***" & vbCrLf)
            End Using

            ' 書き込み成功の通知
            notify.ShowBalloonTip(3000, "成功", "テキストが正常に保存されました", ToolTipIcon.Info)

        Catch ioEx As IOException
            notify.ShowBalloonTip(3000, "エラー", "ファイル操作中にエラーが発生しました: " & ioEx.Message, ToolTipIcon.Error)
        Catch formatEx As FormatException
            notify.ShowBalloonTip(3000, "エラー", "データ形式エラーが発生しました: " & formatEx.Message, ToolTipIcon.Error)
        Catch ex As Exception
            notify.ShowBalloonTip(3000, "エラー", "予期しないエラーが発生しました: " & ex.Message, ToolTipIcon.Error)
        Finally
            System.Threading.Thread.Sleep(3000)
            notify.Visible = False
            notify.Dispose()
        End Try
    End Sub
End Module
