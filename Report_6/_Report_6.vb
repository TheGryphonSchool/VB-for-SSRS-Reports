'File produced by combining files using the Combine Files VScode extension
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\FUNCTIONS.VB
Public Function roundIfFloat(float As String) As String
    If float.Contains(".0") Then
        Return CStr(CInt(float))
    End If
    Return float
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

Public Function storeAndRankGroupLearner(groupID As Integer, _
                                         learnerID As Integer, _
                                         gradePointOnScale As Integer, _
                                         ranks As String) As Double
    Return _
        RankChecker.getInstance().storeAndRankGroupLearner(groupID, _
                                                           learnerID, _
                                                           gradePointOnScale, _
                                                           ranks)
End Function

Public Function getGroupLearnerProblems(groupID As Integer, _
                                        learnerID As Integer) As Integer
    Return _
        RankChecker.getInstance().getGroupLearnerProblems(groupID, learnerID)
End Function

Public Function getGroupLearnerRankDelta(groupID As Integer, _
                                         learnerID As Integer) As Double
    Return _
        RankChecker.getInstance().getGroupLearnerRankDelta(groupID, learnerID)
End Function


' Public Function highlightRank(group_code As String, _
'                               rank As String, _
'                               empty As String, _
'                               bad As String, _
'                               ok As String, _
'                               differing As String) As String
'     'rank will be a string unless it's nothing
'     If IsNothing(rank) Then
'         Return empty
'     ElseIf rank.Contains(",") Then
'         'Two teachers have entered different grades
'         Return differing
'     End If
'     Return IIf(RankChecker.getInstance().isOk(group_code, rank), ok, bad)
' End Function

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

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GROUP.VB
Public Class Group
    Private learnerList As New System.Collections.Generic.List(Of GroupLearner)
    Private learnerDict As New _
        System.Collections.Generic.Dictionary(Of Integer, GroupLearner)
    Public sortedAndAnalysed As Boolean = False

    Public Function storeAndRankLearner(learnerID As Integer, _
                                        gradePointOnScale As Integer, _
                                        ranks As String) As Double
        Dim learner As New GroupLearner(gradePointOnScale, ranks)
        learnerList.Add(learner)
        learnerDict.Add(learnerID, learner)
        Return learner.effectiveRank
    End Function

    Public Function getLearnerProblems(learnerID As Integer) As Integer
        Dim learner As GroupLearner
        sortAndAnalyse()
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            return 0
        End If
        return learner.problemCode
    End Function

    Public Function getLearnerRankDelta(learnerID As Integer) As Integer
        Dim learner As GroupLearner
        sortAndAnalyse()
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            return 0
        End If
        return learner.rankDelta
    End Function

    Private Sub sortAndAnalyse()
        If sortedAndAnalysed Then
            Return
        End If
        learnerList.Sort(AddressOf compareUsingRanks)
        findSkippedRanks()
        findAndStoreRankGradeIncompatibilities()
        ' Find grade/rank invompatibilites & store in learners' rankDelta var
        setRankDeltas(rankRangesForGrades(countEachGrade()))
        sortedAndAnalysed = True
    End Sub

    Private Shared Function compareUsingRanks(learner1 As GroupLearner, _
                                              learner2 As GroupLearner) _
                                              As Integer
        ' Pass method to List#sort() to sort learners by their ranks
        ' Side-effect: If learners have the same rank, sets 2nd bit of problem
        '   code to 1 for each learner
        Dim comparison As Integer =
            learner1.effectiveRank.compareTo(learner2.effectiveRank)
        If comparison = 0 Then
            learner1.problemCode = learner1.problemCode Or 2
            learner2.problemCode = learner2.problemCode Or 2
        End If
        Return comparison
    End Function

    Private Sub findSkippedRanks()
        ' Find learners with ranks more than 1 greater than the learner before
        '   and set 3rd bit of their problem code to 1
        Dim lastEffectiveRank As Double
        For Each learner As GroupLearner In learnerList
            If learner.effectiveRank - lasteffectiveRank > 1.0 Then
                ' 1 or more ranks have been skipped
                learner.problemCode = learner.problemCode Or 4
            End If
            lastEffectiveRank = learner.effectiveRank
        Next learner
    End Sub

    Private Function countEachGrade() As new _
            System.Collections.Generic.Dictionary(of Integer, Integer)
        Dim currentGrade As Integer
        For Each learner As GroupLearner In learnerList
            currentGrade = learner.gradePoints
            If countEachGrade.containsKey(currentGrade) Then
                countEachGrade(currentGrade) += 1
            Else
                countEachGrade.Add(currentGrade, 1)
            End If
        Next learner
    End Function
    

    Private Function rankRangesForGrades(gradeCounts As _
            System.Collections.Generic.Dictionary(of Integer, Integer)) As new _
            System.Collections.Generic.Dictionary(of Integer, Integer(1))
        Dim highestRank As Integer = 0
        Dim newHighestRank As Integer
        For each grade As Integer in gradeCounts.Keys
            newHighestRank = highestRank + gradeCounts(grade)
            rankRangesForGrades.Add(grade, { highestRank + 1, newHighestRank })
            highestRank = newHighestRank
        Next grade
    End Function
    
    Private Sub setRankDeltas(gradeRankRanges As _
            System.Collections.Generic.Dictionary(of Integer, Integer(1)))
        Dim rankRange(1) As Integer
        Dim learnerRank As Double
        For Each learner As GroupLearner In learnerList
            rankRange = gradeRankRanges(learner.gradePoints)
            learnerRank = learner.effectiveRank
            If learnerRank < rankRange(0) Then
                learner.rankDelta = rankRange(0) - learnerRank
            Else If learnerRank > rankRange(1) Then
                learner.rankDelta = rankRange(1) - learnerRank
            Else
                learner.rankDelta = 0
            End If
        Next learner
    End Sub
End Class


'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GROUPLEARNER.VB
Public Class GroupLearner
    Public gradePoints As Integer
    Public effectiveRank As Double
    Public problemCode As Integer
    ' problem code is a pseudo-bitfield whose bits have these meanings:
    '   1st bit => has multiple, conflicting ranks
    '   2nd bit => has same ranks as another learner
    '   3rd bit => has skipped a rank 
    Public rankDelta As Double
    ' The minimum magnitude (signed) change in rank needed for this learner to
    '   have a rank appropriate for their grade. Note that this may place them
    '   in the same rank as another learner
 
    Sub new(gradePointOnScale As Integer, ranks As String)
        problemCode = 0
        gradePoints = gradePointOnScale
        effectiveRank = averageIfNumeric(ranks)
    End Sub

    Private Sub setEffectiveRank(ranks As String)
        ' Side effect: Sets 1st bit of problem code to 1 if ranks conflict
        Dim sum as Double = 0
        Dim count as Integer = 0
        Dim prev As Double = 0
        Dim conflict As Boolean = False
        If ranks is Nothing Then
           effectiveRank 0.0
           Return
        End If
        For Each item As Object In averageArray(ranks.Split(", "))
            Try
                item = CDbl(item) 
                Sum += item
                count += 1
                If prev AndAlso item <> prev Then
                    conflict = True
                End If
                prev = item
            Catch _ex As Exception
                ' Ignore here
            End Try
        Next item
        If conflict Then 
            problemCode = problemCode Or 1
        End If
        If count = 0 Then
            effectiveRank 0.0
        Else If
            effectiveRank sum / count
        End If
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\RANKCHECKER.VB
Public NotInheritable Class RankChecker
    Private Shared singleton_rank_checker As RankChecker
    Private groups As New _
        System.Collections.Generic.Dictionary(Of Integer, Group)

    Public Shared Function getInstance() As RankChecker
        If (singleton_rank_checker Is Nothing) Then
            singleton_rank_checker = New RankChecker()
        End If
        Return singleton_rank_checker
    End Function

    Public Function storeAndRankGroupLearner(groupID As Integer, _
                                             learnerID As Integer, _
                                             gradePointOnScale As Integer, _
                                             ranks As String) As Double
        Dim group As Group
        If Not groups.TryGetValue(groupID, group) Then
            group = new group
            groups.Add(groupID, group)
        End If
        group.storeAndRankLearner(learnerID, gradePointOnScale, ranks)
    End Function

    Public Function getGroupLearnerProblems(groupID As Integer, _
                                            learnerID As Integer) As Integer
        Dim group As Group
        If Not groups.TryGetValue(groupID, group) Then
            return 0
        End If
        group.getLearnerProblems(learnerID)
    End Function

    Public Function getGroupLearnerRankDelta(groupID As Integer, _
                                             learnerID As Integer) As Double
        Dim group As Group
        If Not groups.TryGetValue(groupID, group) Then
            return 0
        End If
        group.getLearnerRankDelta(learnerID)
    End Function
End Class

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
