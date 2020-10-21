Public Class OtherMarks
    Private colMarkMap As New _
        System.Collections.Generic.Dictionary(Of String, MarkList)

    Public Sub AddMark(colName As String, mark As String)
        Dim markList As MarkList
        If colMarkMap.ContainsKey(colName) Then
            markList = colMarkMap(colName)
            markList.Add(mark)
        Else
            colMarkMap.Add(colName, New MarkList(mark))
        End If
    End Sub

    Public Function GetMarks(colName As String) As String
        ' Homework and Classwork cols have been stored as such. Any others are
        ' unreliable. But if there are multiple they'll be named PPE1, PPE2, etc
        Dim colCount As Integer = colMarkMap.Count
        Dim markList As MarkList
        Dim keys(colCount) As String
        Dim values(colCount) As MarkList
        Dim numberFinder As New System.Text.RegularExpressions.Regex("\d+")
        Dim numberInName As String

        If Not colMarkMap.TryGetValue(colName, markList) Then
            colMarkMap.Values.CopyTo(values, 0)
            If colCount = 4 Then
                markList = values(3)
            ElseIf colCount > 4 Then
                numberInName = numberFinder.Match(colName).Value
                colMarkMap.Keys.CopyTo(keys, 0)
                For i As Integer = 3 To colCount
                    If keys(i).Contains(numberInName) Then
                        markList = values(i)
                    End If
                Next
            End If
        End If
        If markList Is Nothing Then
            Return ""
        End If
        Return markList.GetAllMarks()
    End Function
End Class
