Public NotInheritable Class ParamLookups
    Private Shared singleton_instance As ParamLookups
    Private caches As New _
        System.Collections.Generic.Dictionary(Of Object, Object)
        'SSRS_parameter => Dict(search_item => result)
    
    Public Shared Function getInstance() As ParamLookups
        If (singleton_instance Is Nothing) Then
            singleton_instance = New ParamLookups()
        End If
        Return singleton_instance
    End Function

    Public Function searchCaches(param As Object, _
                                 search_item As Object) As Object
        Dim cache As Object = caches(param)
        If cache Is Nothing Then
            Return Nothing
        End If
        Return cache(search_item)
    End Function
    
    Public Sub cacheResult(param As Object, _
                           search_item As Object, _
                           result As Object)
        Dim new_cache As Object = caches(param)
        If new_cache Is Nothing Then
            new_cache = New _
                System.Collections.Generic.Dictionary(Of Object, Object)
            caches.Add(param, new_cache)
        End If
        new_cache.Add(search_item, result)
    End Sub
End Class

' Return the first param that matches the search_item
Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object) As Object
    Return _lookupParam(value_or_label, search_item, param)
End Function

' Return the nth param that matches the search_item
Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object, _
                            nth_match As Integer) As Object
    Return _lookupParam(value_or_label, search_item, param, nth_match)
End Function

' Basic, with caching
Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object, _
                            caching As Boolean) As Object
    Return _lookupParam(value_or_label, search_item, param, 1, caching)
End Function

' Workhorse delegated to by all the overloads
Private Function _lookupParam(value_or_label As String, _
                              search_item As Object, _
                              param As Object, _
                              Optional nth_match As Integer = 1, _
                              Optional caching As Boolean = False) As Object
    Dim searches As Object()
    Dim results As Object()
    Dim found_count = 0
    If param.IsMultiValue Then
        If caching Then
            _lookupParam = _
                ParamLookups.getInstance().searchCaches(param, search_item)
            If _lookupParam IsNot Nothing Then
                Exit Function
            End If
        End If
        value_or_label = value_or_label.toLower()
        searches = IIf(value_or_label = "value", param.Value, param.Label)
        results = IIf(value_or_label = "value", param.Label, param.Value)
        For i As Integer = 0 To param.Count -1
            If search_item.Equals(searches(i)) Then
                found_count += 1
                If found_count.Equals(nth_match) Then
                    _lookupParam = results(i)
                    If caching Then
                        ParamLookups.getInstance().cacheResult(param, _
                                                               search_item, _
                                                               _lookupParam)
                    End If
                    Exit Function
                End If
            End If
        Next i
    ElseIf search_item = IIf(value_or_label = "value", param.Value, param.Label)
        Return IIf(value_or_label = "value", param.Label, param.Value)
    End If
    Return Nothing' if value is not found in parameter
End Function

Public Function lookupNthParam(value_or_label As String, _
                               number As Integer, _
                               param As Object) As Object
    Dim results As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    If number <= param.Count Then
        Return results(number - 1)
    End If
    Return Nothing 'if parameter doesn't have that number of items
End Function

Public Function lookupAllMatchingParams(value_or_label As String, _
                                        search_item As Object, _
                                        param As Object, _
                                        Optional contains As Boolean = False) _
                As Object()
    Dim searches As Object()
    Dim results As Object()
    Dim finds As Object() = {}
    Dim found_count As Integer = 0
    Dim is_match As Boolean
    Dim search As Object
    Dim result As Object

    value_or_label = value_or_label.toLower()
    If param.IsMultiValue Then
        searches = IIf(value_or_label = "value", param.Value, param.Label)
        results = IIf(value_or_label = "value", param.Label, param.Value)
        If contains AndAlso (Not TypeOf search_item Is String OrElse _
                            searches.Length > 0 OrElse _
                            Not TypeOf searches(0) Is String) Then
            contains = False
        End If
        For i As Integer = 0 To param.Count -1
            If contains Then
                is_match = searches(i).Contains(search_item)
            Else
                is_match = search_item.Equals(searches(i))
            End If
            If is_match Then
                ReDim Preserve finds(found_count)
                finds(found_count) = results(i)
                found_count += 1
            End If
        Next i
    Else
        search = IIf(value_or_label = "value", param.Value, param.Label)
        result = IIf(value_or_label = "value", param.Label, param.Value)
        If contains AndAlso Not TypeOf search Is String Then
            contains = False
        End If
        If contains Then
            is_match = search.Contains(search_item)
        Else
            is_match = search_item.Equals(search)
        End If
        If is_match
            Redim Preserve finds(0)
            finds(0) = result
        End If
    End If
    Return finds
End Function
