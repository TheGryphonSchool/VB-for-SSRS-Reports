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