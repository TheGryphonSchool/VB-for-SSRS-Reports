    Private Function StartsWithRegex(start As String) As _
                                 System.Text.RegularExpressions.Regex
        Return New _
    System.Text.RegularExpressions.Regex("^" & EscapeRegexString(start))
    End Function

    Private Function EscapeRegexString(unescaped As String) As String
        ' Escape regex meta-characters in user-supplied string so that a regex can
        ' be built from the string that matches the supplied characters literally
        Dim escRgx As System.Text.RegularExpressions.Regex
        escRgx = New System.Text.RegularExpressions.Regex("[|^$.()?+*\[\]\\]")
        Return escRgx.Replace(unescaped, "\$&")
    End Function

    Private Sub ThrowUnlessSearchesAreSearchable(searches As Object(), _
                                                 searchItem As Object)
        Dim ADVICE As String = "to use the match strategies 'Contains'('C') or 'Starts-with' " & _
        "('S'). Omit the matchStrategy argument to use exact matching."
        If Not TypeOf searchItem Is String Then
            Throw New ArgumentException( _
                "The search item must be a string " & ADVICE)
        ElseIf Not TypeOf searches(0) Is String Then
            Throw New ArgumentException( _
                "The parameter must have string values " & ADVICE)
        End If
    End Sub
