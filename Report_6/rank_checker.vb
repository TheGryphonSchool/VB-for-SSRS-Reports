Public Class RankChecker
    Public badRanks As New System.Collections.Generic.List(Of Integer)

    Public Sub addBad(badRank As Integer)
        badRanks.Add(badRank)
    End Sub

    Public Function isOk(rank As Integer) As Boolean
        Return Not badRanks.Contains(rank)
    End Function
End Class

Public NotInheritable Class RankCheckers
    Private Shared singleton_rank_checker As RankCheckers
    Private rank_checkers_dict As New _
        System.Collections.Generic.Dictionary(Of String, RankChecker)

    Public Shared Function getInstance() As RankCheckers
        If (singleton_rank_checker Is Nothing) Then
            singleton_rank_checker = New RankCheckers()
        End If
        Return singleton_rank_checker
    End Function

    Public Function countOkRanks(group_code As String, _
                                 ranks As Object()) As Integer
        Dim rank_checker As RankChecker = New RankChecker
        Dim prev_rank As Int64 = 0
        countOkRanks = 0
        Array.sort(ranks)
        For Each rank As Object In ranks
            If TypeOf rank Is Integer Then
                If rank - prev_rank = 1 Then
                    countOkRanks += 1
                Else
                    rank_checker.addBad(rank)
                End If
                prev_rank = rank
            End If
        Next rank
        rank_checkers_dict.Add(group_code, rank_checker)
    End Function

    Public Function isOk(group_code As String, rank As String) As Boolean
        Dim rank_checker As RankChecker = rank_checkers_dict(group_code)
        Return rank_checker Is Nothing OrElse rank_checker.isOk(CInt(rank))
    End Function
End Class

Public Function countOkRanks(group_code As String, ranks As Object()) As Integer
    Return RankCheckers.getInstance().countOkRanks(group_code, ranks)
End Function

Public Function highlightRank(group_code As String, _
                              rank As String, _
                              empty As String, _
                              bad As String, _
                              ok As String, _
                              differing As String) As String
    'rank will be a string unless it's nothing
    If IsNothing(rank) Then
        Return empty
    ElseIf rank.Contains(",") Then
        'Two teachers have entered different grades
        Return differing
    End If
    Return IIf(RankCheckers.getInstance().isOk(group_code, rank), ok, bad)
End Function
