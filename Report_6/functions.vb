Public Function roundIfFloat(float As String) As String
    If float.Contains(".0") Then
        Return CStr(CInt(float))
    End If
    Return float
End Function

Public Function lookupCleanedJoinedGrades(value_or_label As String, _
                                          group_learner_column As Object, _
                                          grades_param As Object) As String
    Return cleanAndJoin(lookupAllMatchingParams(value_or_label, _
                                                group_learner_column, _
                                                grades_param))
End Function

Private Function cleanAndJoin(items As Object()) As String
    Dim output As String = ""
    Dim output_length As Integer
    Dim unique_items As New System.Collections.Generic.List(Of Object)
    If items Is Nothing Then
        Return output
    End If
    For Each item As Object In items
        If Not TypeOf item Is String Then
            item = TryCast(item, String)
            item = IIf(item, item, "can't convert to string")
        End If
        If item.Contains(".0") Then
            item = item.Substring(0, item.IndexOf(".0"))
        End If
        If Not unique_items.contains(item) Then
            unique_items.Add(item)
            output += item & ", "
        End If
    Next item
    output_length = output.Length()
    If output_length = 0 Then
        Return output
    End If
    Return output.Substring(0, output_length - 2)
End Function

Public Function storeAndRankGroupLearner(groupID As Integer, _
                                         learnerID As Integer, _
                                         gradePointOnScale As Integer, _
                                         ranks As String) As Double
    Return _
        RankChecker.getInstance().storeAndRankGroupLearner(groupID, _
                                                           learnerID, _
                                                           gradePointOnScale, _
                                                           ranks)
End Function

Public Function getGroupLearnerProblems(groupID As Integer, _
                                        learnerID As Integer) As Integer
    Return _
        RankChecker.getInstance().getGroupLearnerProblems(groupID, learnerID)
End Function

Public Function getGroupLearnerRankDelta(groupID As Integer, _
                                         learnerID As Integer) As Double
    Return _
        RankChecker.getInstance().getGroupLearnerRankDelta(groupID, learnerID)
End Function


' Public Function highlightRank(group_code As String, _
'                               rank As String, _
'                               empty As String, _
'                               bad As String, _
'                               ok As String, _
'                               differing As String) As String
'     'rank will be a string unless it's nothing
'     If IsNothing(rank) Then
'         Return empty
'     ElseIf rank.Contains(",") Then
'         'Two teachers have entered different grades
'         Return differing
'     End If
'     Return IIf(RankChecker.getInstance().isOk(group_code, rank), ok, bad)
' End Function
