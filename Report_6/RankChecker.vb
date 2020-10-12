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
