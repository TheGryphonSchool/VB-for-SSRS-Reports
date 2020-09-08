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

Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object) As Object
    Return _lookupParam(value_or_label, search_item, param)
End Function

Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object, _
                            nth_match As Integer) As Object
    Return _lookupParam(value_or_label, search_item, param, nth_match)
End Function

Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object, _
                            caching As Boolean) As Object
    Return _lookupParam(value_or_label, search_item, param, 1, caching)
End Function

Overloads Public Function lookupParam(value_or_label As String, _
                            search_item As Object, _
                            param As Object, _
                            nth_match As Integer, _
                            caching As Boolean) As Object
    Return _lookupParam(value_or_label, search_item, param, nth_match, caching)
End Function

Private Function _lookupParam(value_or_label As String, _
                             Optional search_item As Object = Nothing, _
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
            If searches(i) = search_item Then
                found_count += 1
                If found_count = nth_match Then
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
    'Return nothing if parameter doesn't have that number of items
End Function

Public Function isInParam(value_or_label As String, _
                           search_item As Object, _
                           param As Object) As Boolean
    Dim lookups As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    Return Array.IndexOf(lookups, search_item) >= 0
End Function
