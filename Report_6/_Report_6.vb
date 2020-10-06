'File produced by combining files using the Combine Files VScode extension
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GRADE_CHECKER.VB
Public NotInheritable Class GradeChecker
    Private latest_group_id As Integer = 0
    Private latest_position As Integer = 0
    Private Shared singleton_grade_checker As GradeChecker

    Public Shared Function getInstance() As GradeChecker
        If (singleton_grade_checker Is Nothing) Then
            singleton_grade_checker = New GradeChecker()
        End If
        Return singleton_grade_checker
    End Function

    Public Function isOk(group_id As Integer, position As Object) As Boolean
        If group_id <> latest_group_id Then
            'First grade in group (i.e. top rank)
            latest_group_id = group_id
        ElseIf position < latest_position Then
            Return False
        End If
        latest_position = position
        Return True
    End Function
End Class

Public Function highlightGrade(group_id As Integer, _
                               position As Object, _
                               bad As String, _
                               ok As String) As String
    If IsNothing(position) OrElse _
        GradeChecker.getInstance().isOk(group_id, CInt(position)) Then
        'If position is nothing, these aren't grades, so don't highlight them
        Return ok
    End If
    Return bad
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\MISCELLANEOUS.VB
Public Function roundIfFloat(float As String) As String
    If float.Contains(".0") Then
        Return CStr(CInt(float))
    End If
    Return float
End Function

Public Function averageIfNumeric(ranks As String) As Integer
    If ranks is Nothing Then
        Return 0
    End If
    Return averageArray(ranks.Split(", "))
End Function

Public Function lookupCleanedJoinedGrades(value_or_label As String, _
                                          group_learner_column As Object, _
                                          grades_param As Object) As String
    Return cleanAndJoin(lookupAllMatchingParams(value_or_label, _
                                                group_learner_column, _
                                                grades_param))
End Function

Private Function cleanAndJoin(items As Object()) As String
    Dim output As String = ""
    Dim output_length As Integer
    Dim unique_items As New System.Collections.Generic.List(Of Object)
    If items Is Nothing Then
        Return output
    End If
    For Each item As Object In items
        If Not TypeOf item Is String Then
            item = TryCast(item, String)
            item = IIf(item, item, "can't convert to string")
        End If
        If item.Contains(".0") Then
            item = item.Substring(0, item.IndexOf(".0"))
        End If
        If Not unique_items.contains(item) Then
            unique_items.Add(item)
            output += item & ", "
        End If
    Next item
    output_length = output.Length()
    If output_length = 0 Then
        Return output
    End If
    Return output.Substring(0, output_length - 2)
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\RANK_CHECKER.VB
Public Class RankChecker
    Public badRanks As New System.Collections.Generic.List(Of Integer)

    Public Sub addBad(badRank As Integer)
        badRanks.Add(badRank)
    End Sub

    Public Function isOk(rank As Integer) As Boolean
        Return Not badRanks.Contains(rank)
    End Function
End Class

Public NotInheritable Class RankCheckers
    Private Shared singleton_rank_checker As RankCheckers
    Private rank_checkers_dict As New _
        System.Collections.Generic.Dictionary(Of String, RankChecker)

    Public Shared Function getInstance() As RankCheckers
        If (singleton_rank_checker Is Nothing) Then
            singleton_rank_checker = New RankCheckers()
        End If
        Return singleton_rank_checker
    End Function

    Public Function countOkRanks(group_code As String, _
                                 ranks As Object()) As Integer
        Dim rank_checker As RankChecker = New RankChecker
        Dim prev_rank As Int64 = 0
        countOkRanks = 0
        Array.sort(ranks)
        For Each rank As Object In ranks
            If TypeOf rank Is Integer Then
                If rank - prev_rank = 1 Then
                    countOkRanks += 1
                Else
                    rank_checker.addBad(rank)
                End If
                prev_rank = rank
            End If
        Next rank
        rank_checkers_dict.Add(group_code, rank_checker)
    End Function

    Public Function isOk(group_code As String, rank As String) As Boolean
        Dim rank_checker As RankChecker = rank_checkers_dict(group_code)
        Return rank_checker Is Nothing OrElse rank_checker.isOk(CInt(rank))
    End Function
End Class

Public Function countOkRanks(group_code As String, ranks As Object()) As Integer
    Return RankCheckers.getInstance().countOkRanks(group_code, ranks)
End Function

Public Function highlightRank(group_code As String, _
                              rank As String, _
                              empty As String, _
                              bad As String, _
                              ok As String, _
                              differing As String) As String
    'rank will be a string unless it's nothing
    If IsNothing(rank) Then
        Return empty
    ElseIf rank.Contains(",") Then
        'Two teachers have entered different grades
        Return differing
    End If
    Return IIf(RankCheckers.getInstance().isOk(group_code, rank), ok, bad)
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

Public Function averageArray(items As Object()) As Double
	Dim sum as Double = 0
	Dim count as Integer = 0
	For Each item As Double In items
		If Not IsNumeric(item) Then
            Try
                item = CDbl(item) 
            Catch _ex As Exception
                item = 0
            End Try
        End If
        sum += item
		count += 1
	Next item
	If count = 0 Then
        Return 0
    End If
    Return sum / count
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
        If fraction = 1.0 Then
            Return scale(last_index)
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
