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
