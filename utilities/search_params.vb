Public Function CountMatchingParams(value_or_label As String, _
                                    search_item As Object, _
                                    param As Object,
                                    Optional match_strategy As Integer = 0) _
                As Integer
    Dim searches As Object()
    Dim found_count As Integer = 0
    Dim search As Object

    If match_strategy < 2 And Not TypeOf search_item Is String Then
        Throw New ArgumentException(
            "The search item must be a string to use the match strategies " & _
            "'Contains' or 'Begins-with'. Pass 2 as the fourth parameter " & _
            "to use exact matching")
    End If
    value_or_label = value_or_label.toLower()
    If param.IsMultiValue Then
        If value_or_label = "value" Then
            searches = param.Value
        Else
            searches = param.Label
        End If
        If searches.Length < 1 Then
            Throw New ArgumentException("The parameter you passed is empty!")
        ElseIf match_strategy < 2 And Not TypeOf searches(0) Is String Then
            Throw New ArgumentException(
                "The parameter must have string values to use the " & _
                "match strategies, 'Contains' or 'Begins-with'. Pass " & _
                "2 as the fourth parameter to use exact matching")
        End If
        For i As Integer = 0 To param.Count -1
            If matches(searches(i), search_item, match_strategy) Then
                found_count += 1
            End If
        Next i
    Else
        If value_or_label = "value" Then
            search = param.Value
        Else
            search = param.Label
        End If
        If Not TypeOf search Is String Then
            Throw New ArgumentException("The parameter must be a string")
        End If
        found_count = IIf(search.Contains(search_item), 1, 0)
    End If
    Return found_count
End Function

Private Function matches(candidate As Object, _
                         criterion As Object, _
                         strategy As Integer) _
                 As Boolean
    Select Case strategy
        Case 0
            Return candidate.Contains(criterion)
        Case 1
            Return Left(candidate, criterion.Length) = criterion
        Case Else
            Return criterion.Equals(candidate)
    End Select
End Function


Public Function isInParam(value_or_label As String, _
                          search_item As Object, _
                          param As Object) As Boolean
    Dim lookups As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    Return Array.IndexOf(lookups, search_item) >= 0
End Function
