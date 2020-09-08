Public Function lookupNthMatchingParam(value_or_label As String, _
                                       search_item As Object, _
                                       number As Integer, _
                                       param As Object) As Object
    Dim searches As Object()
    Dim results As Object()
    Dim found_count As Integer = 1
    value_or_label = value_or_label.toLower()
    searches = IIf(value_or_label = "value", param.Value, param.Label)
    results = IIf(value_or_label = "value", param.Label, param.Value)
    For i As Integer = 0 To param.Count -1
        If searches(i) = search_item Then
            If found_count = number Then
                Return results(i)
            End If
            found_count += 1
        End If
    Next i
    Return Nothing
End Function

Public Function lookupNthParam(value_or_label As String, _
                               number As Integer, _
                               param As Object) As Object
    Dim results As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    If number <= param.Count Then
        Return results(number - 1)
    End If
    'Return nothing if parameter doesn't have that number of items
End Function
