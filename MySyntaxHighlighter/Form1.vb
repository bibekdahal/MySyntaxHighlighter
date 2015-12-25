Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SyntaxHighlighter1.ColorAllRTB()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim firstChar As Integer = SyntaxHighlighter1.SelectionStart
        SyntaxHighlighter1.Paste(DataFormats.GetFormat(DataFormats.Text))
        SyntaxHighlighter1.SelectionColor = Color.Black

        SyntaxHighlighter1.ColorLines(SyntaxHighlighter1.GetLineFromCharIndex(firstChar), SyntaxHighlighter1.GetLineFromCharIndex(SyntaxHighlighter1.SelectionStart))

        SyntaxHighlighter1.Focus()
    End Sub

End Class
