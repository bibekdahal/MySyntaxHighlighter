Public Class SyntaxHighlighter
    Private Words As New DataTable

    Private Declare Function GetScrollPos Lib "user32" Alias "GetScrollPos" (ByVal hWnd As IntPtr, ByVal NBar As Integer) As Integer
    Private Declare Function SetScrollPos Lib "user32" Alias "SetScrollPos" (ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Boolean) As Integer
    Private Declare Function PostMessageA Lib "user32" Alias "PostMessageA" (ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Boolean

    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer

    Private keywords As String()
    'Contains Windows Messages for the SendMessage API call
    Private Enum EditMessages
        LineIndex = 187
        LineFromChar = 201
        GetFirstVisibleLine = 206
        CharFromPos = 215
        PosFromChar = 1062
    End Enum


    Private Const SB_HORZ As Integer = 0
    Private Const SB_VERT As Integer = 1
    Private Const WM_HSCROLL As Integer = 276
    Private Const WM_VSCROLL As Integer = 277
    Private Const SB_THUMBPOSITION As Integer = 4


    Private Property HScrollPos() As Integer
        Get
            Return GetScrollPos(Me.Handle.ToInt32, SB_HORZ)
        End Get
        Set(ByVal value As Integer)
            SetScrollPos(DirectCast(Me.Handle, IntPtr), SB_HORZ, value, True)
            PostMessageA(DirectCast(Me.Handle, IntPtr), WM_HSCROLL, SB_THUMBPOSITION + 65536 * value, 0)
        End Set
    End Property

    Private Property VScrollPos() As Integer
        Get
            Return GetScrollPos(Me.Handle.ToInt32, SB_VERT)
        End Get
        Set(ByVal value As Integer)
            SetScrollPos(DirectCast(Me.Handle, IntPtr), SB_VERT, value, True)
            PostMessageA(DirectCast(Me.Handle, IntPtr), WM_VSCROLL, SB_THUMBPOSITION + 65536 * value, 0)
        End Set
    End Property

    Public Function AddSyntaxWord(ByVal strWord As String, ByVal clrColor As Color)
        Dim MyRow As DataRow
        MyRow = Words.NewRow()
        MyRow("Word") = strWord
        MyRow("Color") = clrColor.Name
        Words.Rows.Add(MyRow)
        Return True
    End Function
    Public Function ClearSyntaxWords()
        Words = New DataTable
        'Load all the keywords and the colors to make them 
        Words.Columns.Add("Word")
        Words.PrimaryKey = New DataColumn() {Words.Columns(0)}
        Words.Columns.Add("Color")
        Return True
    End Function

    Public Sub ColorAllRtb()

        ColorCoding = False
        Dim vvv As Long = VScrollPos
        Dim hhh As Long = HScrollPos

        RefreshAll()

        VScrollPos = vvv
        HScrollPos = hhh
        ColorCoding = True
    End Sub
    Private Sub AddSyntax()
        ClearSyntaxWords()

        AddSyntaxWord("\b(addhandler|addressof|alias|and|andalso|as|boolean|byref|byte|byval|call|case|catch|cbool|cbyte|cchar|" _
               & "cdate|cdec|cdbl|char|cint|class|clng|cobj|const|continue|csbyte|cshort|csng|cstr|ctype|cuint|culng|" _
               & "cushort|date|decimal|declare|default|delegate|dim|directcast|do|double|each|else|elseif|" _
               & "end|endif|enum|erase|error|event|exit|false|finally|for|friend|function|get|gettype|global|" _
               & "gosub|goto|handles|if|implements|imports|in|inherits|integer|interface|is|isnot|let|lib|" _
               & "like|long|loop|me|mod|module|mustinherit|mustoverride|mybase|myclass|namespace|narrowing|" _
               & "my|new|next|not|nothing|notinheritable|notoverridable|object|of|on|operator|option|optional|" _
               & "or|orelse|overloads|overridable|overrides|paramarray|partial|private|property|protected|public|" _
               & "raiseevent|readonly|redim|rem|removehandler|resume|return|sbyte|select|" _
               & "set|shadows|shared|short|single|static|step|stop|string|structure|sub|synclock|then|throw|to|true|" _
               & "try|trycast|typeof|variant|wend|uinteger|ulong|ushort|using|when|while|widening|with|withevents|" _
               & "writeonly|xor|#const|#else|#elseif|#end|#if)\b", Color.Blue) '|-|&|&=|*|*=|/|/=|\|\=|^|^=|+|+=|=|-=|" _

        AddSyntaxWord("'.*", Color.Green)
        AddSyntaxWord("(?<!'.*)\""[^""]*\""", Color.DarkRed)

        keywords = New String("addhandler|addressof|alias|and|andalso|as|boolean|byref|byte|byval|call|case|catch|cbool|cbyte|cchar|" _
               & "cdate|cdec|cdbl|char|cint|class|clng|cobj|const|continue|csbyte|cshort|csng|cstr|ctype|cuint|culng|" _
               & "cushort|date|decimal|declare|default|delegate|dim|directcast|do|double|each|else|elseif|" _
               & "end|endif|enum|erase|error|event|exit|false|finally|for|friend|function|get|gettype|global|" _
               & "gosub|goto|handles|if|implements|imports|in|inherits|integer|interface|is|isnot|let|lib|" _
               & "like|long|loop|me|mod|module|mustinherit|mustoverride|mybase|myclass|namespace|narrowing|" _
               & "my|new|next|not|nothing|notinheritable|notoverridable|object|of|on|operator|option|optional|" _
               & "or|orelse|overloads|overridable|overrides|paramarray|partial|private|property|protected|public|" _
               & "raiseevent|readonly|redim|rem|removehandler|resume|return|sbyte|select|" _
               & "set|shadows|shared|short|single|static|step|stop|string|structure|sub|synclock|then|throw|to|true|" _
               & "try|trycast|typeof|variant|wend|uinteger|ulong|ushort|using|when|while|widening|with|withevents|" _
               & "writeonly|xor|#const|#else|#elseif|#end|#if").Split("|")
    End Sub
    Private Sub ColorLine(ByVal LineNumber As Integer)
        On Error Resume Next

        Dim SelectionAt As Integer = Me.SelectionStart
        Dim SelectionLen As Integer = Me.SelectionLength
        Dim SelColor As Color = Me.SelectionColor

        Dim MyRow As DataRow

        ' Lock the update
        LockWindowUpdate(Me.Handle.ToInt32)

        Me.SelectionStart = Me.GetFirstCharIndexFromLine(LineNumber)
        Me.SelectionLength = Me.Lines(LineNumber).Length
        Me.SelectionColor = Color.Black
        'Check for matches in a particular line number
        Dim rm As System.Text.RegularExpressions.MatchCollection
        Dim m As System.Text.RegularExpressions.Match

        For Each MyRow In Words.Rows

            rm = System.Text.RegularExpressions.Regex.Matches(Me.Lines(LineNumber).ToLower, MyRow("Word"))
            For Each m In rm
                Me.SelectionStart = Me.GetFirstCharIndexFromLine(LineNumber) + m.Index
                Me.SelectionLength = m.Length
                Me.SelectionColor = Color.FromName(MyRow("color"))
            Next
        Next

        ' Restore the selectionstart
        Me.SelectionStart = SelectionAt
        Me.SelectionLength = SelectionLen
        Me.SelectionColor = SelColor
        ' Unlock the update
        LockWindowUpdate(0)
    End Sub

    Private Sub ColorMultipleLines(ByVal FirstLineNumber As Integer, ByVal LastLineNumber As Integer)
        On Error Resume Next

        Dim SelectionAt As Integer = Me.SelectionStart
        Dim SelectionLen As Integer = Me.SelectionLength
        Dim SelColor As Color = Me.SelectionColor

        Dim MyRow As DataRow

        ' Lock the update
        LockWindowUpdate(Me.Handle.ToInt32)

        Dim textToCheck As String = ""
        For ii As Integer = FirstLineNumber To LastLineNumber
            textToCheck = textToCheck + Me.Lines(ii) + Chr(10)
        Next

        Me.SelectionStart = Me.GetFirstCharIndexFromLine(FirstLineNumber)
        Me.SelectionLength = textToCheck.Length
        Me.SelectionColor = Color.Black

        'Check for matches in a particular line number
        Dim rm As System.Text.RegularExpressions.MatchCollection
        Dim m As System.Text.RegularExpressions.Match

        For Each MyRow In Words.Rows

            rm = System.Text.RegularExpressions.Regex.Matches(textToCheck.ToLower, MyRow("Word"))
            For Each m In rm
                Me.SelectionStart = Me.GetFirstCharIndexFromLine(FirstLineNumber) + m.Index
                Me.SelectionLength = m.Length
                Me.SelectionColor = Color.FromName(MyRow("color"))
            Next
        Next

        ' Restore the selectionstart
        Me.SelectionStart = SelectionAt
        Me.SelectionLength = SelectionLen
        Me.SelectionColor = SelColor
        ' Unlock the update
        LockWindowUpdate(0)
    End Sub

    Private Sub ColorTextPart(ByVal FirstCharIndex As Integer, ByVal LastCharIndex As Integer)
        On Error Resume Next

        Dim SelectionAt As Integer = Me.SelectionStart
        Dim SelectionLen As Integer = Me.SelectionLength
        Dim SelColor As Color = Me.SelectionColor
        Dim MyRow As DataRow

        ' Lock the update
        LockWindowUpdate(Me.Handle.ToInt32)

        Dim texttocheck As String
        Me.SelectionStart = FirstCharIndex
        Me.SelectionLength = LastCharIndex - FirstCharIndex
        texttocheck = Me.SelectedText

        'Check for matches in a particular line number
        Dim rm As System.Text.RegularExpressions.MatchCollection
        Dim m As System.Text.RegularExpressions.Match

        For Each MyRow In Words.Rows

            rm = System.Text.RegularExpressions.Regex.Matches(texttocheck.ToLower, MyRow("Word"))
            For Each m In rm
                Me.SelectionStart = FirstCharIndex + m.Index
                Me.SelectionLength = m.Length
                Me.SelectionColor = Color.FromName(MyRow("color"))
            Next
        Next

        ' Restore the selectionstart
        Me.SelectionStart = SelectionAt
        Me.SelectionLength = SelectionLen
        Me.SelectionColor = SelColor
        ' Unlock the update
        LockWindowUpdate(0)
    End Sub

    Private Sub RefreshAll()
        On Error Resume Next

        Dim SelectionAt As Integer = Me.SelectionStart
        Dim SelectionLen As Integer = Me.SelectionLength
        Dim SelColor As Color = Me.SelectionColor

        Dim MyRow As DataRow

        ' Lock the update
        LockWindowUpdate(Me.Handle.ToInt32)

        Me.SelectionStart = 1
        Me.SelectionLength = Me.Text.Length
        Me.SelectionColor = Color.Black

        'Check for matches in a particular line number
        Dim rm As System.Text.RegularExpressions.MatchCollection
        Dim m As System.Text.RegularExpressions.Match

        For Each MyRow In Words.Rows
            rm = System.Text.RegularExpressions.Regex.Matches(Me.Text.ToLower, MyRow("Word"))
            For Each m In rm
                Me.SelectionStart = m.Index
                Me.SelectionLength = m.Length
                Me.SelectionColor = Color.FromName(MyRow("color"))
            Next
        Next

        ' Restore the selectionstart
        Me.SelectionStart = SelectionAt
        Me.SelectionLength = SelectionLen
        Me.SelectionColor = SelColor
        ' Unlock the update
        LockWindowUpdate(0)
    End Sub

    Public Sub ColorLines(ByVal FirstLineNumber As Integer, ByVal LastLineNumber As Integer)

        ColorCoding = False
        Dim vvv As Long = VScrollPos
        Dim hhh As Long = HScrollPos
        If FirstLineNumber = LastLineNumber Then
            ColorLine(FirstLineNumber)
        Else
            ColorMultipleLines(FirstLineNumber, LastLineNumber)
        End If
        VScrollPos = vvv
        HScrollPos = hhh
        ColorCoding = True
    End Sub

    Public Sub ColorPartOfText(ByVal FirstCharIndex As Integer, ByVal LastCharIndex As Integer)

        ColorCoding = False
        Dim vvv As Long = VScrollPos
        Dim hhh As Long = HScrollPos
        ColorTextPart(FirstCharIndex, LastCharIndex)
        VScrollPos = vvv
        HScrollPos = hhh
        ColorCoding = True
    End Sub

    Private ColorCoding As Boolean = True

    Private Sub SyntaxHighlighter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ColorCoding = True Then
            Me.SelectionColor = Color.Black
        End If
    End Sub


    Private Sub SyntaxHighlighter_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If ColorCoding = True Then
            If Me.Text <> "" Then
                ColorLines(Me.GetLineFromCharIndex(Me.SelectionStart), Me.GetLineFromCharIndex(Me.SelectionStart + Me.SelectionLength))
            End If
        End If
    End Sub

End Class
