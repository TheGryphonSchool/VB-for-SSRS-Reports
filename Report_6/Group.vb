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

