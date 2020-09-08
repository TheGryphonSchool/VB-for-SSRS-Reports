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

Public Function AverageCollection(items As Object()) As Double
	Dim sum as Double = 0
	Dim count as Integer = 0
	For Each item As Double In items
		sum += item
		count += 1
	Next item
	If count = 0 Then
        Return 0
    End If
    Return sum / count
End Function

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

Public Function isInParam(value_or_label As String, _
                           search_item As Object, _
                           param As Object) As Boolean
    Dim lookups As Object() = _
        IIf(value_or_label.toLower() = "value", param.Value, param.Label)
    Return Array.IndexOf(lookups, search_item) >= 0
End Function

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

Public Function lookupNthMatchingParam(value_or_label As String, _
                                       search_item As Object, _
                                       param As Object) As Boolean
    Return True
End Function


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

Public Function roundIfFloat(float As String) As String
    If Not float.Contains(".0") Then
        Return float
    End If
    Return CStr(CInt(float))
End Function
