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

    Private Sub ThrowIfMatchStrategyTypeConflict(searches As Object(), _
                                                 searchItem As Object, _
                                                 matchStrategy As Char)
        ThrowUnlessSearchIsString(searchItem, matchStrategy)
        ThrowUnlessSearchesAreStrings(searches, matchStrategy)
    End Sub

    Private Sub ThrowUnlessSearchIsString(searchItem As Object, _
                                          matchStrategy As Char)
        If TypeOf searchItem Is String Then Exit Sub
        Throw New ArgumentException(MatchStrategyExceptionMessage( _
            "The search item must be a string", matchStrategy))
    End Sub

    Private Sub ThrowUnlessSearchesAreStrings(searches As Object(), _
                                              matchStrategy As Char)
        If TypeOf searches(0) Is String Then Exit Sub
        Throw New ArgumentException(MatchStrategyExceptionMessage( _
            "The parameter must have string values", matchStrategy))
    End Sub

    Private Function MatchStrategyExceptionMessage(problemStatement As String, _
                                                   matchStrategy As Char) As String
        Dim strategyDescription As String
        Select Case matchStrategy
            Case "C"C
                strategyDescription = "'Contains' ('C')"
            Case "S"C
                strategyDescription = "'Starts-with' ('S')"
            Case Else
                strategyDescription = "'Regular Expression' ('R')"
        End Select
        Return problemStatement & " to use the match strategy " & _
            strategyDescription & _
            ". Omit the matchStrategy argument to use exact matching."
    End Function
