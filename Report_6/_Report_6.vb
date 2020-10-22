'File produced by combining files using the Combine Files VScode extension
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\MISCELLANEOUS.VB
Public Function roundIfFloat(float As String) As String
    If Not float.Contains(".0") Then
        Return float
    End If
    Return CStr(CInt(float))
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\ARRAYS.VB
Public Function removeDuplicates(items As Object()) As Object()
    'maps original indicies to new indicies
    Dim index_map As New _
        System.Collections.Generic.Dictionary(Of Integer, Integer)
    Dim shift As Integer = 0
    Dim unique_items As New System.Collections.Generic.List(Of Object)
    Dim index As Integer = 0
    For Each item As Object In items
        If unique_items.contains(item) Then
            shift += 1
        Else
            unique_items.Add(item)
            index_map.Add(index, index - shift)
        End If
        index += 1
    Next item
    For Each old_index As Integer In index_map.Keys
        items(index_map(old_index)) = items(old_index)
    Next old_index
    ReDim Preserve items(index - 1 - shift)
    Return items
End Function

Public Function SumArray(nums() As Object) As Integer
    Dim num As Integer
	SumArray = 0
	For Each o As Object In nums
        If Integer.TryParse(o, num)
            SumArray += num
        End If
	Next o
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\COLOUR_SCALE.VB
Public Class ColourScale
    Private scale As New System.Collections.Generic.List(Of Integer())

    Public Sub New(first As String, _
                   second As String, _
                   Optional third As String = "", _
                   Optional fourth As String = "", _
                   Optional fifth As String = "")
        For Each nth As String In New String(4) {first, second, third, fourth, fifth}  
            If nth Is "" Then
                Exit For
            End If
            addToScale(nth)
        Next nth
    End Sub

    Public Function getColour(fraction As Double)
        Dim last_index As Integer = scale.Count - 1
        Dim start As Integer
        If fraction >= 1.0 Then
            Return mixTwoColours(1.0, last_index - 1)
        End If
        start = CInt(Math.Floor(fraction * last_index))
        Return mixTwoColours(fraction * last_index - start, start)
    End Function

    Private Sub addToScale(hexColour As String)
        Dim rgb(2) As Integer
        hexColour = hexColour.Replace("#", "")
        For i As Integer = 0 To 2
            rgb(i) = Convert.ToInt32(hexColour.Substring(i * 2, 2), 16)
        Next i
        scale.Add(rgb)
    End Sub

    Private Function mixTwoColours(fraction As Double, _
                                   start_index As Integer) As String
        Dim starts As Integer
        Dim ends As Integer
        mixTwoColours = "#"
        For i As Integer = 0 To 2
            starts = scale.Item(start_index)(i)
            ends = scale.Item(start_index + 1)(i)
            mixTwoColours += _
                Hex(CInt(starts + fraction * (ends - starts))).PadLeft(2,  "0")
        Next i
    End Function
End Class

Dim header_colour_scale As ColourScale

Public Function colourFromScale(fraction As Double, _
                                first As String, _
                                second As String, _
                                third As String) As String
    If header_colour_scale Is Nothing Then
        header_colour_scale = New ColourScale(first, second, third)
    End If
    Return header_colour_scale.getColour(fraction)
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\LOOKUP_PARAMS.VB
Public NotInheritable Class ParamLookups
    Private Shared singleton_instance As ParamLookups
    Private caches As New _
        System.Collections.Generic.Dictionary(Of Object, Object)
        'SSRS_parameter => Dict(search_item => result)
    
    Public Shared Function GetInstance() As ParamLookups
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
                ParamLookups.GetInstance().searchCaches(param, search_item)
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
                        ParamLookups.GetInstance().cacheResult(param, _
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

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\SEARCH_PARAMS.VB
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
    System.Text.RegularExpressions.Regex("^" & EscapeRegexString(start))
End Function

Private Function EscapeRegexString(unescaped As String) As String
    ' Escape regex meta-characters in user-supplied string so that a regex can
    ' be built from the string that matches the supplied characters literally
    Dim esc_rgx As System.Text.RegularExpressions.Regex
    esc_rgx = New System.Text.RegularExpressions.Regex("[|^$.()?+*\[\]\\]")
    Return esc_rgx.Replace(unescaped, "\$&")
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
