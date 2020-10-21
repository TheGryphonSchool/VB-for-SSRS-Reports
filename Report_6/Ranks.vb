Public Class Ranks
    Private list As New System.Collections.Generic.List(Of Double)
    Public effectiveRank As Double
    Public conflict As Boolean = False

    Public Sub Add(rank As Double)
        list.Add(rank)
    End Sub

    Public Function GetEffectiveRank() As Double
        ' Side effect: Sets 1st `conflicts` var to True if ranks conflict
        Dim sum As Double = 0
        Dim count As Integer = 0
        Dim prevRank As Double = 0
        If list Is Nothing Then
            effectiveRank = 0
            Return effectiveRank
        End If
        For Each rank As Double In list
            sum += rank
            count += 1
            If prevRank AndAlso rank <> prevRank Then
                conflict = True
            End If
            prevRank = rank
        Next rank
        If count = 0 Then
            effectiveRank = 0
        Else
            effectiveRank = sum / count
        End If
        Return effectiveRank
    End Function

    Public Function GetAllRanks() As String
        ' Result has the form `rank1, rank2` if ranks conflict, else `rank1`
        If list.Count = 0 Then
            Return ""
        End If
        If conflict Then
            GetAllRanks = ""
            For Each rank As Double In list
                GetAllRanks = GetAllRanks & CStr(CInt(rank)) & ", "
            Next rank
            If list.Count > 0 Then
                GetAllRanks = Left(GetAllRanks, Len(GetAllRanks) - 2)
            End If
        Else
            Return CStr(CInt(list(0)))
        End If
    End Function
End Class
