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
