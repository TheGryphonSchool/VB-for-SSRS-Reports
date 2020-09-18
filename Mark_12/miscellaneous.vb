Dim any_mark_6s  As Boolean

Public Function IsMark6(value_or_label As String, _
                           search_item As Object, _
                           param As Object) As Boolean
    IsMark6 = isInParam(value_or_label, search_item, param)
    If IsMark6 Then
        any_mark_6s = True
    End If
End Function

Public Function AnyMark6s() As Boolean
    Return any_mark_6s
End Function

Public NotInheritable Class GradeStore
    Private Shared singleton_grade_store As GradeStore
    Private grades_dict As new _
        System.Collections.Generic.Dictionary(Of String, Object())

    Public Shared Function getInstance() As GradeStore
        If (singleton_grade_store Is Nothing) Then
            singleton_grade_store = New GradeStore()
        End If
        Return singleton_grade_store
    End Function

    Public Sub add_grades(group_learner As String, grades() As Object)
        grades_dict.Add(group_learner, grades)
    End Sub

    Public Function get_grade(group_learner As String, index As Integer) _
                    As Object
        Dim grades() As Object
        If Not grades_dict.TryGetValue(group_learner, grades) OrElse _
           index < 0 OrElse grades.Length <= index Then
            Return ""
        End If
        Return grades(index)
    End Function
    
End Class

Public Function StoreAndCountGrades(group_learner As Object, _
                                    param As Object) As Integer
    Dim grades() As Object = _
        lookupAllMatchingParams("Label", group_learner, param, True)
    GradeStore.getInstance().add_grades(group_learner, grades)
    Return grades.Length
End Function

Public Function GetNthGrade(group_learner As String, n As Integer) As Object
    ' To conform with respective conventions, n will begin at 1 in the SSRS API,
    '  but at 0 in the private VB code.
    Return GradeStore.getInstance().get_grade(group_learner, n - 1)
End Function

