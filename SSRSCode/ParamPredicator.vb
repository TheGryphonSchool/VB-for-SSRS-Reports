Namespace SSRSCode
    Friend Class ParamPredicator
        Public predicate As Predicate(Of Object)
        Protected ReadOnly searchItem As Object
        Protected ReadOnly searchString As String
        Protected reg As Text.RegularExpressions.Regex

        Public Shared Function Create(searchItem As Object, _
                                      matchStrategy As Char, _
                                      Optional nthMatch As Integer = 1) _
                                      As ParamPredicator
            If nthMatch = 1 Then Return New ParamPredicator(searchItem, matchStrategy)

            Return New TargetedPredicator(searchItem, matchStrategy, nthMatch)
        End Function

        Protected Sub New(searchItem As Object, _
                          matchStrategy As Char)
            Me.searchItem = searchItem
            Me.searchString = searchItem.ToString
            Select Case matchStrategy
                Case "C"C ' Contains
                    predicate = AddressOf Contains
                Case "R"C ' Regular Expression
                    predicate = AddressOf Regex
                Case "S"C ' Starts-with
                    predicate = AddressOf StartsWith
                Case Else ' Equals
                    predicate = AddressOf EqualTo
            End Select
        End Sub

        Public Function Contains(candidate As Object) As Boolean
            Return candidate.ToString().Contains(searchString)
        End Function

        Friend Function StartsWith(candidate As Object) As Boolean
            Return candidate.ToString().StartsWith(searchString)
        End Function

        Friend Function EqualTo(candidate As Object) As Boolean
            Return searchItem.Equals(candidate)
        End Function

        Friend Function Regex(candidate As Object) As Boolean
            If IsNothing(reg) Then
                reg = New Text.RegularExpressions.Regex(searchString)
            End If
            Return reg.IsMatch(candidate.ToString())
        End Function

        Private Class TargetedPredicator
            Inherits ParamPredicator
            Public ReadOnly subPredicate As Predicate(Of Object)
            Protected ReadOnly nthMatch As Integer
            Protected matchCount As Integer = 0

            Sub New(searchItem As Object, _
                              matchStrategy As Char, _
                              Optional nthMatch As Integer = 1)
                MyBase.New(searchItem, matchStrategy)
                Me.nthMatch = nthMatch
                subPredicate = predicate
                predicate = AddressOf AdjustForCount
            End Sub

            Private Function AdjustForCount(candidate As Object) As Boolean
                If Not subPredicate.Invoke(candidate) Then Return False

                matchCount += 1
                Return matchCount >= nthMatch
            End Function
        End Class
    End Class
End Namespace
