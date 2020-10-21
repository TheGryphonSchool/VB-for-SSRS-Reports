Public NotInheritable Class RankChecker
    Private Shared singleton_rank_checker As RankChecker
    Private groups As New _
        System.Collections.Generic.Dictionary(Of Integer, Group)
    Private columns As New _
        System.Collections.Generic.Dictionary(Of Integer, String)
    Private closedForMarkEntry As Boolean = False

    Public Shared Function GetInstance() As RankChecker
        If (singleton_rank_checker Is Nothing) Then
            singleton_rank_checker = New RankChecker()
        End If
        Return singleton_rank_checker
    End Function

    Public Sub StoreColumn(columnID As Integer, displayName As String)
        If columns.ContainsKey(columnID) Then
            Return
        End If
        columns.Add(columnID, CategoriseColumn(displayName))
    End Sub

    Public Sub StoreGroupLearnerMark(groupID As Integer, _
                                     learnerID As Integer, _
                                     value As String, _
                                     points As String, _
                                     columnID As Integer)
        Dim group As Group
        Dim rankOrGrade As String

        If closedForMarkEntry Then
            Return
        End If
        If columns.TryGetValue(columnID, rankOrGrade) Then
            If Not groups.TryGetValue(groupID, group) Then
                group = New Group
                groups.Add(groupID, group)
            End If
            group.StoreLearnerMark(learnerID, value, points, rankOrGrade)
        End If
    End Sub

    Public Function GetGroupLearnerMark(groupID As Integer, _
                                        learnerID As Integer, _
                                        colName As String) As String
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return ""
        End If
        Return group.GetLearnerMark(learnerID, colName)
    End Function

    Public Function GetGroupLearnerProblems(groupID As Integer, _
                                            learnerID As Integer) As Integer
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return 0
        End If
        Return group.GetLearnerProblems(learnerID)
    End Function

    Public Function GetGroupLearnerRankDelta(groupID As Integer, _
                                             learnerID As Integer) As Double
        Dim group As Group
        closedForMarkEntry = True
        If Not groups.TryGetValue(groupID, group) Then
            Return 0
        End If
        Return group.GetLearnerRankDelta(learnerID)
    End Function

    Private Function CategoriseColumn(displayName As String)
        If displayName.Contains("redict") Then
            Return "Grade"
        ElseIf displayName.Contains("ank") Then
            Return "Rank"
        ElseIf displayName.Contains("omework") Then
            Return "Homework"
        ElseIf displayName.Contains("lasswork") Then
            Return "Classwork"
        Else
            Return displayName
        End If
    End Function

    Public Function ShowGrades() As String
        ShowGrades = ""
        For Each groupID As Integer In groups.Keys
            ShowGrades = ShowGrades & CStr(groupID) & vbCrLf & _
                         groups(groupID).showLearners() & vbCrLf & vbCrLf
        next groupID 
    End Function

    Public Function ShowColumns() As String
        ShowColumns = ""
        For Each columnID As Integer In columns.Keys
            ShowColumns = ShowColumns & CStr(columnID) & ": " & _
                          columns(columnID) & vbCrLf
        next columnID 
    End Function
End Class
