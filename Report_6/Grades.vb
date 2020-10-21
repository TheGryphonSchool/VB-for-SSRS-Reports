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
