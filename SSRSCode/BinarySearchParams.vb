Namespace SSRSCode
    Module BinarySearchParams
        Friend Function BinSearchParam(searches() As Object, _
                                       results() As Object, _
                                       searchItem As Object, _
                                       matchStrategy As Char, _
                                       nthMatch As Integer) As Object
            Dim comp As Object = GetComparer(matchStrategy)
            Dim randMatch As Integer = _
                Array.BinarySearch(searches, searchItem, comp)
            Return NthMatchFromAMatch(searches, results, searchItem, nthMatch, _
                                      randMatch, comp)
        End Function

        Private Function NthMatchFromAMatch(searches() As Object, _
                                            results() As Object, _
                                            searchItem As Object, _
                                            nthMatch As Integer, _
                                            matchIndex As Integer, _
                                            comp As Object) As Object

            If matchIndex < 0 Then Return Nothing
            ' Return any match if the caller doesn't care:
            If nthMatch = 0 Then Return results(matchIndex)

            ' Index of leftmost match found so far
            Dim firstIndex As Integer = matchIndex
            While firstIndex > 0 AndAlso _
                    AreComparable(searchItem, searches(firstIndex - 1), comp)
                firstIndex -= 1
            End While
            Dim nthMatchingIndex As Integer = firstIndex + nthMatch - 1
            If nthMatchingIndex >= searches.Length OrElse _
                nthMatchingIndex > matchIndex AndAlso _
                Not AreComparable(searchItem, searches(nthMatchingIndex), comp) Then
                ' Out of bounds or-else not a match
                Return Nothing
            End If

            Return results(nthMatchingIndex)
        End Function

        Private Function AreComparable(searchItem As Object, _
                                       candidate As Object, _
                                       comp As Object) As Boolean
            If comp Is Nothing Then
                Return searchItem.CompareTo(candidate) = 0
            End If

            Return comp.PredicateCompare(candidate, searchItem)
        End Function

        Private Function GetComparer(matchStrategy As Char) As Object
            Select Case matchStrategy
                Case "C"C ' Contains
                    Throw New InvalidBinarySearchException("Contains")
                Case "R"C ' Regular Expression
                    Throw New InvalidBinarySearchException("Regular Expression")
                Case "S"C ' Starts-with
                    Return New StartsWithComparer()
                Case Else ' Equals (use default CompareTo implementation)
                    Return Nothing
            End Select
        End Function

        Private Class StartsWithComparer
            Implements System.Collections.IComparer

            Public Function Compare(candidate As Object, searchItem As Object) _
                    As Integer Implements System.Collections.IComparer.Compare
                If PredicateCompare(candidate, searchItem) Then Return 0

                Return candidate.CompareTo(searchItem)
            End Function

            Public Function PredicateCompare(candidate As Object, _
                                             searchItem As Object) As Boolean
                Return candidate.ToString().StartsWith(searchItem.ToString())
            End Function
        End Class

        Private Class InvalidBinarySearchException
            Inherits ArgumentException

            Public Sub New(strategyName As String)
                MyBase.New("You may not use binary-search with the '" _
                           & strategyName & "' match-strategy. Use " _
                           & "'Equals' or 'Starts-with' instead.")
            End Sub
        End Class

    End Module
End Namespace
