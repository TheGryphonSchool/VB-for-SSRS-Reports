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
