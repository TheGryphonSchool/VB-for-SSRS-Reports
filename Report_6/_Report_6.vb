'File produced by combining files using the Combine Files VScode extension
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\FUNCTIONS.VB
Public Function StoreColumn(columnID As Integer, displayName As String)
    RankChecker.GetInstance().StoreColumn(columnID, displayName)
    Return True
End Function

Public Function StoreGroupLearnerMark(groupID As Integer, _
                                      learnerID As Integer, _
                                      value As String, _
                                      points As String, _
                                      columnID As Integer) As Boolean
    RankChecker.GetInstance().StoreGroupLearnerMark( _
        groupID, learnerID, value, points, columnID)
    Return True
End Function

Public Function GetGroupLearnerMark(groupID As Integer, _
                                learnerID As Integer, _
                                colName As String) As String
    Return RankChecker.GetInstance().GetGroupLearnerMark(groupID, _
                                                        learnerID, _
                                                        colName)
End Function

Public Function GetGroupLearnerProblems(groupID As Integer, _
                                    learnerID As Integer) As Integer
    Return _
        RankChecker.GetInstance().GetGroupLearnerProblems(groupID, learnerID)
End Function

Public Function GetGroupLearnerRankDelta(groupID As Integer, _
                                        learnerID As Integer) As Double
    Return _
        RankChecker.GetInstance().GetGroupLearnerRankDelta(groupID, learnerID)
End Function

Public Function ShowGrades() As String
    Return RankChecker.GetInstance().ShowGrades()
End Function

Public Function ShowColumns() As String
    Return RankChecker.GetInstance().ShowColumns()
End Function
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GRADES.VB
Public Class Grades
    Inherits MarkList
    Private effectiveGradePoints As Double = -1
    Private gradePoints As New System.Collections.Generic.List(Of Double)

    Public Overloads Sub Add(letter As String, points As Double)
        MyBase.Add(letter)
        gradePoints.Add(points)
    End Sub

    Public Function GetEffectiveGradePoints() As Double
        If effectiveGradePoints < 0 Then
            ' First time this is called, so effective grade hasn't been set
            SetEffectiveGradePoints()
        End If
        Return effectiveGradePoints
    End Function

    Private Sub SetEffectiveGradePoints()
        ' Return True if there are conflicting grades
        Dim sum As Double = 0
        Dim count As Integer = 0
        Dim prevPoints As Double = 0
        If gradePoints Is Nothing Then
            effectiveGradePoints = 0
            Return
        End If
        For Each points As Double In gradePoints
            sum += points
            count += 1
            If prevPoints AndAlso points <> prevPoints Then
                conflict = True
            End If
            prevPoints = points
        Next points
        If count = 0 Then
            effectiveGradePoints = 0
        Else
            effectiveGradePoints = sum / count
        End If
    End Sub
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GROUP.VB
Public Class Group
    Private learnerList As New System.Collections.Generic.List(Of GroupLearner)
    Private learnerDict As New _
        System.Collections.Generic.Dictionary(Of Integer, GroupLearner)
    Public sortedAndAnalysed As Boolean = False

    Public Sub StoreLearnerMark(learnerID As Integer, _
                                value As String, _
                                points As String, _
                                rankOrGrade As String)
        Dim learner As GroupLearner
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            learner = New GroupLearner()
            learnerList.Add(learner)
            learnerDict.Add(learnerID, learner)
        End If
        Select Case rankOrGrade
            Case "Rank"
                learner.AddRank(CDbl(points))
            Case "Grade"
                learner.AddGrade(value, CDbl(points))
            Case Else
                learner.AddOtherMark(rankOrGrade, value)
        End Select
    End Sub

    Public Function GetLearnerMark(learnerID As Integer, colName As String) _
                                   As String
        Dim learner As GroupLearner
        SortAndAnalyse()
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            Return ""
        End If
        Select Case colName
            Case "Rank"
                Return learner.GetAllRanks()
            Case "Grade"
                Return learner.GetAllGrades()
            Case Else
                Return learner.GetOtherMarks(colName)
        End Select
    End Function

    Public Function GetLearnerProblems(learnerID As Integer) As Integer
        Dim learner As GroupLearner
        SortAndAnalyse()
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            Return 0
        End If
        Return learner.problemCode
    End Function

    Public Function GetLearnerRankDelta(learnerID As Integer) As Double
        Dim learner As GroupLearner
        SortAndAnalyse()
        If Not learnerDict.TryGetValue(learnerID, learner) Then
            Return 0
        End If
        Return learner.rankDelta
    End Function

    Private Sub SortAndAnalyse()
        If sortedAndAnalysed Then
            Return
        End If
        learnerList.Sort(AddressOf CompareUsingRanks)
        FindSkippedRanks()
        ' Find grade/rank invompatibilites & store in learners' rankDelta var
        SetRankDeltas(RankRangesForGrades(CountEachGrade()))
        sortedAndAnalysed = True
    End Sub

    Private Shared Function CompareUsingRanks(learner1 As GroupLearner, _
                                              learner2 As GroupLearner) _
                                              As Integer
        ' Pass method to List#sort() to sort learners by their ranks
        ' Side-effect: If learners have the same rank, sets 2nd bit of problem
        '   code to 1 for each learner
        Dim comparison As Integer =
            learner1.GetEffectiveRank().CompareTo(learner2.GetEffectiveRank())
        If comparison = 0 Then
            learner1.problemCode = learner1.problemCode Or 2
            learner2.problemCode = learner2.problemCode Or 2
        End If
        Return comparison
    End Function

    Private Sub FindSkippedRanks()
        ' Find learners with ranks more than 1 greater than the learner before
        '   and set 3rd bit of their problem code to 1
        Dim lastEffectiveRank As Double
        For Each learner As GroupLearner In learnerList
            If learner.GetEffectiveRank() - lastEffectiveRank >= 2.0 Then
                ' 1 or more ranks have been skipped
                learner.problemCode = learner.problemCode Or 4
            End If
            lastEffectiveRank = learner.GetEffectiveRank()
        Next learner
    End Sub

    Private Function CountEachGrade() As _
            System.Collections.Generic.Dictionary(Of Double, Integer)
        Dim currentGrade As Double
        CountEachGrade = New _
            System.Collections.Generic.Dictionary(Of Double, Integer)
        For Each learner As GroupLearner In learnerList
            currentGrade = learner.GetEffectiveGradePoints()
            If CountEachGrade.ContainsKey(currentGrade) Then
                CountEachGrade.Item(currentGrade) += 1
            Else
                CountEachGrade.Add(currentGrade, 1)
            End If
        Next learner
    End Function

    Private Function RankRangesForGrades(gradeCounts As _
            System.Collections.Generic.Dictionary(Of Double, Integer)) As _
            System.Collections.Generic.Dictionary(Of Double, Integer())
        Dim highestRank As Integer = 0
        Dim newHighestRank As Integer
        Dim gradeArray(gradeCounts.Count - 1) As Double
        RankRangesForGrades = New _
            System.Collections.Generic.Dictionary(Of Double, Integer())
        gradeCounts.Keys.CopyTo(gradeArray, 0)
        Array.Sort(gradeArray)
        Array.Reverse(gradeArray)
        For Each grade As Double In gradeArray
            newHighestRank = highestRank + gradeCounts(grade)
            RankRangesForGrades.Add(grade, {highestRank + 1, newHighestRank})
            highestRank = newHighestRank
        Next grade
    End Function

    Private Sub SetRankDeltas(gradeRankRanges As _
            System.Collections.Generic.Dictionary(Of Double, Integer()))
        Dim learnerRank As Double
        Dim rankRange() As Integer
        For Each learner As GroupLearner In learnerList
            rankRange = gradeRankRanges(learner.GetEffectiveGradePoints())
            learnerRank = learner.GetEffectiveRank()
            If learnerRank < rankRange(0) Then
                learner.rankDelta = rankRange(0) - learnerRank
            ElseIf learnerRank > rankRange(1) Then
                learner.rankDelta = rankRange(1) - learnerRank
            Else
                learner.rankDelta = 0
            End If
        Next learner
    End Sub

    Public Function ShowLearners() As String
        For Each LearnerID As Integer In learnerDict.Keys
            ShowLearners = ShowLearners & "  " & CStr(learnerID) & ":" & _
                           vbCrLf & learnerDict(learnerID).ShowMarks() & vbCrLf
        Next LearnerID
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\GROUPLEARNER.VB
Public Class GroupLearner
    Public grades As New Grades
    Public ranks As New Ranks
    Public problemCode As Integer
    ' problem code is a pseudo-bitfield whose bits have these meanings:
    '   1st bit => has multiple, conflicting ranks
    '   2nd bit => has same rank as another learner
    '   3rd bit => has skipped a rank
    Public rankDelta As Double
    ' The minimum magnitude (signed) change in rank needed for this learner to
    '   have a rank appropriate for their grade. Note that this may place them
    '   in the same rank as another learner
    Private otherMarks As New OtherMarks()

    Public Sub AddRank(newRank As Double)
        ranks.Add(newRank)
    End Sub

    Public Sub AddGrade(newGradeLetter As String, newGradePoints As Double)
        grades.Add(newGradeLetter, newGradePoints)
    End Sub

    Public Sub AddOtherMark(colName As String, mark As String)
        otherMarks.AddMark(colName, mark)
    End Sub

    Public Function GetAllRanks() As String
        Return ranks.GetAllRanks()
    End Function

    Public Function GetEffectiveRank()
        ' Sets 1st bit of problem code to 1 if ranks conflict
        GetEffectiveRank = ranks.GetEffectiveRank()
        If ranks.conflict Then
            problemCode = problemCode Or 1
        End If
    End Function

    Public Function GetAllGrades() As String
        Return grades.GetAllMarks()
    End Function

    Public Function GetEffectiveGradePoints() As Double
        Return grades.GetEffectiveGradePoints()
    End Function

    Public Function GetOtherMarks(colName As String) As String
        Return otherMarks.GetMarks(colName)
    End Function

    Public Function ShowMarks() As String
        ShowMarks = "    Ranks: " & ranks.GetAllRanks() & vbCrLf & _
                  "    Grades: " & grades.GetAllMarks()
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\MARKLIST.VB
Public Class MarkList
    Private list As New System.Collections.Generic.List(Of String)
    Public conflict As Boolean = False

    Public Sub New(initMark As String)
        list.Add(initMark)
    End Sub

    Public Sub New()
        ' For sub-classes
    End Sub

    Public Sub Add(mark As String)
        list.Add(mark)
        If list.Count > 0 AndAlso list(0) <> mark Then
            conflict = True
        End If
    End Sub

    Public Function GetAllMarks() As String
        ' Result has the form `mark1, mark2` if they differ, else `mark1`
        If list.Count = 0 Then
            Return ""
        End If
        If conflict Then
            GetAllMarks = ""
            For Each mark As String In list
                GetAllMarks = GetAllMarks & mark & ", "
            Next mark
            GetAllMarks = Left(GetAllMarks, Len(GetAllMarks) - 2)
        Else
            Return list(0)
        End If
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\OTHERMARKS.VB
Public Class OtherMarks
    Private colMarkMap As New _
        System.Collections.Generic.Dictionary(Of String, MarkList)

    Public Sub AddMark(colName As String, mark As String)
        Dim markList As MarkList
        If colMarkMap.ContainsKey(colName) Then
            markList = colMarkMap(colName)
            markList.Add(mark)
        Else
            colMarkMap.Add(colName, New MarkList(mark))
        End If
    End Sub

    Public Function GetMarks(colName As String) As String
        ' Homework and Classwork cols have been stored as such. Any others are
        ' unreliable. But if there are multiple they'll be named PPE1, PPE2, etc
        Dim colCount As Integer = colMarkMap.Count
        Dim markList As MarkList
        Dim keys(colCount) As String
        Dim values(colCount) As MarkList
        Dim numberFinder As New System.Text.RegularExpressions.Regex("\d+")
        Dim numberInName As String

        If Not colMarkMap.TryGetValue(colName, markList) Then
            colMarkMap.Values.CopyTo(values, 0)
            If colCount = 4 Then
                markList = values(3)
            ElseIf colCount > 4 Then
                numberInName = numberFinder.Match(colName).Value
                colMarkMap.Keys.CopyTo(keys, 0)
                For i As Integer = 3 To colCount
                    If keys(i).Contains(numberInName) Then
                        markList = values(i)
                    End If
                Next
            End If
        End If
        If markList Is Nothing Then
            Return ""
        End If
        Return markList.GetAllMarks()
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\RANKCHECKER.VB
Public NotInheritable Class RankChecker
    Private Shared singleton_rank_checker As RankChecker
    Private groups As New _
        System.Collections.Generic.Dictionary(Of Integer, Group)
    Private columns As New _
        System.Collections.Generic.Dictionary(Of Integer, String)
    Private closedForMarkEntry As Boolean = False

    Public Shared Function GetInstance() As RankChecker
        If (singleton_rank_checker Is Nothing) Then
            singleton_rank_checker = New RankChecker()
        End If
        Return singleton_rank_checker
    End Function

    Public Sub StoreColumn(columnID As Integer, displayName As String)
        If columns.ContainsKey(columnID) Then
            Return
        End If
        columns.Add(columnID, CategoriseColumn(displayName))
    End Sub

    Public Sub StoreGroupLearnerMark(groupID As Integer, _
                                     learnerID As Integer, _
                                     value As String, _
                                     points As String, _
                                     columnID As Integer)
        Dim group As Group
        Dim rankOrGrade As String

        If closedForMarkEntry Then
            Return
        End If
        If columns.TryGetValue(columnID, rankOrGrade) Then
            If Not groups.TryGetValue(groupID, group) Then
                group = New Group
                groups.Add(groupID, group)
            End If
            group.StoreLearnerMark(learnerID, value, points, rankOrGrade)
        End If
    End Sub

    Public Function GetGroupLearnerMark(groupID As Integer, _
                                        learnerID As Integer, _
                                        colName As String) As String
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return ""
        End If
        Return group.GetLearnerMark(learnerID, colName)
    End Function

    Public Function GetGroupLearnerProblems(groupID As Integer, _
                                            learnerID As Integer) As Integer
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return 0
        End If
        Return group.GetLearnerProblems(learnerID)
    End Function

    Public Function GetGroupLearnerRankDelta(groupID As Integer, _
                                             learnerID As Integer) As Double
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return 0
        End If
        Return group.GetLearnerRankDelta(learnerID)
    End Function

    Private Function CategoriseColumn(displayName As String)
        If displayName.Contains("redict") Then
            Return "Grade"
        ElseIf displayName.Contains("ank") Then
            Return "Rank"
        ElseIf displayName.Contains("omework") Then
            Return "Homework"
        ElseIf displayName.Contains("lasswork") Then
            Return "Classwork"
        Else
            Return displayName
        End If
    End Function

    Public Function ShowGrades() As String
        ShowGrades = ""
        For Each groupID As Integer In groups.Keys
            ShowGrades = ShowGrades & CStr(groupID) & vbCrLf & _
                         groups(groupID).showLearners() & vbCrLf & vbCrLf
        next groupID 
    End Function

    Public Function ShowColumns() As String
        ShowColumns = ""
        For Each columnID As Integer In columns.Keys
            ShowColumns = ShowColumns & CStr(columnID) & ": " & _
                          columns(columnID) & vbCrLf
        next columnID 
    End Function
End Class

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\RANKS.VB
Public Class Ranks
    Private list As New System.Collections.Generic.List(Of Double)
    Public effectiveRank As Double
    Public conflict As Boolean = False

    Public Sub Add(rank As Double)
        list.Add(rank)
    End Sub

    Public Function GetEffectiveRank() As Double
        ' Side effect: Sets 1st `conflicts` var to True if ranks conflict
        Dim sum As Double = 0
        Dim count As Integer = 0
        Dim prevRank As Double = 0
        If list Is Nothing Then
            effectiveRank = 0
            Return effectiveRank
        End If
        For Each rank As Double In list
            sum += rank
            count += 1
            If prevRank AndAlso rank <> prevRank Then
                conflict = True
            End If
            prevRank = rank
        Next rank
        If count = 0 Then
            effectiveRank = 0
        Else
            effectiveRank = sum / count
        End If
        Return effectiveRank
    End Function

    Public Function GetAllRanks() As String
        ' Result has the form `rank1, rank2` if ranks conflict, else `rank1`
        If list.Count = 0 Then
            Return ""
        End If
        If conflict Then
            GetAllRanks = ""
            For Each rank As Double In list
                GetAllRanks = GetAllRanks & CStr(CInt(rank)) & ", "
            Next rank
            If list.Count > 0 Then
                GetAllRanks = Left(GetAllRanks, Len(GetAllRanks) - 2)
            End If
        Else
            Return CStr(CInt(list(0)))
        End If
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
