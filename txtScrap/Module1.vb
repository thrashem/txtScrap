Imports System.IO
Module Module1
    Sub Main()
        If My.Computer.Clipboard.ContainsText() Then
            My.Computer.Clipboard.SetText(My.Computer.Clipboard.GetText(), System.Windows.Forms.TextDataFormat.Text)
            Dim dt As DateTime = DateTime.Now
            Dim sw As System.IO.StreamWriter = Nothing
            Dim txtScrapDirectory As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\txtScrap\"
            Dim txtScrapFile As String = txtScrapDirectory & dt.ToString("yyyyMMdd") & ".txt"
            Dim utf_8 As System.Text.Encoding = System.Text.Encoding.UTF8
            If System.IO.Directory.Exists(txtScrapDirectory) Then
            Else
                System.IO.Directory.CreateDirectory(txtScrapDirectory)
            End If
            sw = New System.IO.StreamWriter(txtScrapFile, True, utf_8)
            sw.Write(dt.ToString("yyyy/MM/dd HH:mm:ss") & vbCrLf & My.Computer.Clipboard.GetText() & vbCrLf)
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If
        End If
    End Sub
End Module