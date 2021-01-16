Namespace TestSSRSCode
    Friend Class LookupParamTestCase
        Public searchItem As Object
        Public matchStrategy As Char
        Public expected As Object
        Public binSearch As Boolean
        Public nthMatch As Integer

        Public Sub New(searchItem As Object, matchStrategy As Char, _
                        expected As Object, Optional nthMatch As Integer = 1, _
                        Optional binSearch As Boolean = False)
            Me.searchItem = searchItem
            Me.matchStrategy = matchStrategy
            Me.expected = expected
            Me.binSearch = binSearch
            Me.nthMatch = nthMatch
        End Sub
    End Class
End Namespace
