Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SSRSCode.SSRSCode.LookupParams

Namespace TestSSRSCode
    <TestClass>
    Public Class TestLookupParams
        Private param As MockParam

        <TestInitialize>
        Sub Initialize()
            param = New MockParam( _
                {1, 2, 3, 4, 5}, _
                {"1#Predicted", "2#Rank", "2#Schoolwork", "3#Homework", "3#PPE"} _
            )
        End Sub

        <DataRow("label", "2#Rank", 1, "E"C, 2, _
                 DisplayName:="Find 2nd item using 'equals'")>
        <DataRow("label", "^1#.*[Pp]redict", 1, "R"C, 1, _
                 DisplayName:="Find 1st item with regex")>
        <DataRow("label", "^3#.*", 2, "R"C, 5, _
                 DisplayName:="Find the 2nd item (value 5) that matches regex")>
        <DataRow("label", "2#Sch", 1, "S"C, 3, _
                 DisplayName:="Find 3rd item using 'starts-with'")>
        <DataRow("label", "Homework", 1, "C"C, 4, _
                 DisplayName:="Find 4th item using 'contains'")>
        <DataRow("label", "MISSING", 1, "R"C, Nothing, _
                 DisplayName:="Non-match returns nothing")>
        <DataRow("value", 2, 1, "E"C, "2#Rank", _
                 DisplayName:="Use int to find 2nd item using 'equals'")>
        <DataRow("value", 1, 1, "R"C, "1#Predicted", _
                 DisplayName:="Use int to find 1st item with regex")>
        <DataRow("value", 3, 1, "S"C, "2#Schoolwork", _
                 DisplayName:="Use int to find 3rd item using 'starts-with'")>
        <DataRow("value", 4, 1, "C"C, "3#Homework", _
                 DisplayName:="Use int to find 4th item using 'contains'")>
        <DataTestMethod>
        Sub TestLookupParam(valueOrLabel As String, _
                            searchTerm As Object, _
                            nthMatch As Integer, _
                            matchStrategy As Char, _
                            expected As Object)
            With New LookupParamTestCase(searchTerm, matchStrategy, expected)
                Assert.AreEqual( _
                    .expected, _
                    LookupParam(valueOrLabel, .searchItem, param, nthMatch, _
                                .matchStrategy)
                    )
            End With
        End Sub

        <DataRow("2#Rank", 1, "E"C, 2, _
                 DisplayName:="Find 2nd item using 'equals'")>
        <DataRow("2#", 1, "S"C, 2, _
                 DisplayName:="Find 2nd item using 'starts-with'")>
        <DataRow("2#", 2, "S"C, 3, _
                 DisplayName:="Find 2nd matching item (value 3) using 'starts-with'")>
        <DataTestMethod>
        Sub TestBinarySearchLookup(searchTerm As String, _
                                   nthMatch As Integer, _
                                   matchStrategy As Char, _
                                   expected As Object)
            With New LookupParamTestCase(searchTerm, matchStrategy, expected)
                Assert.AreEqual(.expected, LookupParam("Label", .searchItem, param, _
                                                       nthMatch, .matchStrategy, True)
                    )
            End With
        End Sub

        <TestMethod>
        Sub TestArbitraryMatchBinaryLookup()
            ' If the caller omits the nthMatch param, the first match found
            ' should be returned. (This will be the middle of the array if it
            ' matches)
            Dim orderlessParam As MockParam = _
                New MockParam({1, 1, 1}, {"1", "2", "3"})
            Assert.AreEqual("2", LookupParam("value", 1, orderlessParam, True))
        End Sub

        <DataRow(True, "Boolean")>
        <DataRow({"ar", "ray"}, "String[]")>
        <DataTestMethod>
        Sub TestBadSearchItemTypeThrowsError(searchItem As Object, _
                                             messageFragment As String)
            Try
                LookupParam("Value", searchItem, param, "S"C)
            Catch e As Exception
                Assert.IsInstanceOfType(e, GetType(ArgumentException))
                Assert.IsTrue(e.Message.Contains(messageFragment))
            End Try
        End Sub

        <DataRow("2#Rank", "E"C, {2}, _
                 DisplayName:="Find 2nd item using 'equals'")>
        <DataRow("^2#.*k", "R"C, {2, 3}, _
                 DisplayName:="Find 'rank' and 'schoolwork' with regex")>
        <DataRow("3#", "S"C, {4, 5}, _
                 DisplayName:="Find last 2 items using 'starts-with'")>
        <DataRow("work", "C"C, {3, 4}, _
                 DisplayName:="Find 'school'- & 'home'-work using 'contains'")>
        <DataRow("MISSING", "R"C, New Integer() {}, _
                 DisplayName:="Non-match returns empty array")>
        <DataTestMethod>
        Sub TestLookupAllMatchingParams(searchItem As Object, _
                                        matchStrategy As Char, _
                                        expected As Integer())
            Dim result As Object() = _
                LookupAllMatchingParams("label", searchItem, param, matchStrategy)
            Assert.AreEqual(expected.Length, result.Length)
            For i As Integer = 0 To expected.GetUpperBound(0)
                Assert.AreEqual(expected(i), result(i))
            Next
        End Sub
    End Class
End Namespace
