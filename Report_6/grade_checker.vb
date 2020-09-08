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
