Public Class MarkList
    Private list As New System.Collections.Generic.List(Of String)
    Public conflict As Boolean = False

    Public Sub New(initMark As String)
        list.Add(initMark)
    End Sub

    Public Sub New()
        ' For sub-classes
    End Sub

    Public Sub Add(mark As String)
        list.Add(mark)
        If list.Count > 0 AndAlso list(0) <> mark Then
            conflict = True
        End If
    End Sub

    Public Function GetAllMarks() As String
        ' Result has the form `mark1, mark2` if they differ, else `mark1`
        If list.Count = 0 Then
            Return ""
        End If
        If conflict Then
            GetAllMarks = ""
            For Each mark As String In list
                GetAllMarks = GetAllMarks & mark & ", "
            Next mark
            GetAllMarks = Left(GetAllMarks, Len(GetAllMarks) - 2)
        Else
            Return list(0)
        End If
    End Function
End Class
