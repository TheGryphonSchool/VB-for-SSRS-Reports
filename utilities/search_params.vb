Public Function CountMatchingParams(value_or_label As String, _
                                    search_item As Object, _
                                    param As Object,
                                    Optional match_strategy As Integer = 0) _
                As Integer
    Dim searches As Object()
    Dim found_count As Integer = 0
    Dim starts_with_regex As System.Text.RegularExpressions.Regex

    ThrowIfSearchAndStrategyMismatched(search_item, match_strategy)
    value_or_label = value_or_label.toLower()
    If Not param.IsMultiValue Then
        Return CountInSingleValueParam(value_or_label, search_item, _
                                       param, match_strategy)
    End If
    searches = IIf(value_or_label = "value", param.Value, param.Label)
    ThrowUnlessSearchesAreSearchable(searches, match_strategy)
    Select Case match_strategy
        Case 0
            For i As Integer = 0 To param.Count -1
                If searches(i).Contains(search_item) Then
                    found_count += 1
                End If
            Next i
        Case 1
            starts_with_regex = StartsWithRegex(search_item)
            For i As Integer = 0 To param.Count -1
                If starts_with_regex.IsMatch(searches(i)) Then
                    found_count += 1
                End If
            Next i
        Case Else
            For i As Integer = 0 To param.Count -1
                If search_item.Equals(searches(i)) Then
                    found_count += 1
                End If
            Next i
    End Select
    Return found_count
End Function

Private Function CountInSingleValueParam(value_or_label As String, _
                                         search_item As String, _
                                         param As Object, _
                                         match_strategy As Integer) _
                                         As Integer
    Dim search As Object

    search = IIf(value_or_label = "value", param.Value, param.Label)
    If Not TypeOf search Is String Then
        Throw New ArgumentException("The parameter must be a string")
    End If
    Select Case match_strategy
        Case 0
            Return IIf(search.Contains(search_item), 1, 0)
        Case 1
            Return IIF(StartsWithRegex(search_item).IsMatch(search), 1, 0)
        Case Else
            Return IIf(search_item.Equals(search), 1, 0)
    End Select
End Function

Private Function StartsWithRegex(start As String) As _
                                 System.Text.RegularExpressions.Regex
    Return New _
    System.Text.RegularExpressions.Regex("^" & start)
End Function

Private Sub ThrowIfSearchAndStrategyMismatched(search_item As Object, _
                                               strategy As Integer)
    If strategy < 2 And Not TypeOf search_item Is String Then
        Throw New ArgumentException(
            "The search item must be a string to use the match strategies " & _
            "'Contains' (0) or 'Begins-with' (1). Pass 2 as the fourth" & _
            "parameter to use exact matching")
    End If
End Sub

Private Sub ThrowUnlessSearchesAreSearchable(searches As Object(), _
                                             strategy As Integer)
    If searches.Length < 1 Then
        Throw New ArgumentException("The parameter you passed is empty!")
    ElseIf strategy < 2 And Not TypeOf searches(0) Is String Then
        Throw New ArgumentException(
            "The parameter must have string values to use the " & _
            "match strategies, 'Contains' or 'Begins-with'. Pass " & _
            "2 as the fourth parameter to use exact matching")
    End If
End Sub

Public Function isInParam(value_or_label As String, _
                          search_item As Object, _
                          param As Object) As Boolean
    Dim lookups As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    Return Array.IndexOf(lookups, search_item) >= 0
End Function
