    ' Dependent on utilities/param_helpers.vb
    ' It must be combined if this file is

    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object) As Integer
        Return CountMatchingParams(valueOrLabel, searchItem, param)
    End Function

    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object, _
                                        matchStrategy As Char) As Integer
        Dim searches As Object()
        Dim foundCount As Integer = 0
        Dim regexForStartsWith As System.Text.RegularExpressions.Regex

        valueOrLabel = valueOrLabel.ToLower()
        If Not param.IsMultiValue Then
            Return CountInSingleValueParam(valueOrLabel, searchItem, _
                                       param, matchStrategy)
        End If
        searches = IIf(valueOrLabel = "value", param.Value, param.Label)
        Select Case matchStrategy
            Case "C"C ' Contains
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                For i As Integer = 0 To param.Count - 1
                    If searches(i).Contains(searchItem) Then
                        foundCount += 1
                    End If
                Next i
            Case "S"C ' Starts-with
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                regexForStartsWith = StartsWithRegex(searchItem)
                For i As Integer = 0 To param.Count - 1
                    If regexForStartsWith.IsMatch(searches(i)) Then
                        foundCount += 1
                    End If
                Next i
            Case Else ' Equals
                For i As Integer = 0 To param.Count - 1
                    If searchItem.Equals(searches(i)) Then
                        foundCount += 1
                    End If
                Next i
        End Select
        Return foundCount
    End Function

    Private Function CountInSingleValueParam(valueOrLabel As String, _
                                             searchItem As String, _
                                             param As Object, _
                                             matchStrategy As Char) As Integer
        Dim search As Object

        search = IIf(valueOrLabel = "value", param.Value, param.Label)
        If Not TypeOf search Is String Then
            Throw New ArgumentException("The parameter must be a string")
        End If
        Select Case matchStrategy
            Case "C" ' Contains
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(search.Contains(searchItem), 1, 0)
            Case "S" ' Starts-with
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(StartsWithRegex(searchItem).IsMatch(search), 1, 0)
            Case Else ' Equals
                Return IIf(searchItem.Equals(search), 1, 0)
        End Select
    End Function

    Public Function IsInParam(valueOrLabel As String, _
                              searchItem As Object, _
                              param As Object) As Boolean
        Dim lookups As Object() = _
            IIf(valueOrLabel.ToLower() = "value", param.Value, param.Label)
        Return Array.IndexOf(lookups, searchItem) >= 0
    End Function
