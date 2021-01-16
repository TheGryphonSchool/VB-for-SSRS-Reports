Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SSRSCode.SSRSCode.SearchParams

Namespace TestSSRSCode
    <TestClass>
    Public Class TestSearchArrays
        Private arry As Object()

        <TestInitialize>
        Sub Initialize()
            arry = {"a", "by", "bz", "c", "dz", "c", "c"}
        End Sub

        <DataRow("X", False)>
        <DataRow("c", True)>
        <DataTestMethod>
        Sub TestCanDetectInArray(searchItem As String, result As Boolean)
            Assert.AreEqual(result, IsInArray(searchItem, arry))
        End Sub

        <DataRow("X", "E"C, 0, DisplayName:="Absent")>
        <DataRow("c", "E"C, 3, DisplayName:="Present")>
        <DataRow("b", "S"C, 2, DisplayName:="'Starts-with' match strategy")>
        <DataRow("z", "C"C, 2, DisplayName:="'Contains' match strategy")>
        <DataTestMethod>
        Sub TestCountsInArray(searchItem As String, _
                              matchStrategy As Char, _
                              count As Integer)
            Assert.AreEqual(count, CountInArray(searchItem, matchStrategy, arry))
        End Sub

        <TestMethod>
        Sub TestCountsPresentItemCorrectlyUsingDefault()
            Assert.AreEqual(3, CountInArray("c", arry))
        End Sub
    End Class

    <TestClass>
    Public Class TestSearchParamsForPairs
        Private param As Object

        <TestInitialize>
        Sub Initialize()
            param = New MockParam({"a", "b", "c", "c"}, {"A", "B", "D", "C"})
        End Sub

        <DataRow("a", False, DisplayName:="Absent")>
        <DataRow("c", True, DisplayName:="Present")>
        <DataTestMethod>
        Sub TestMatchSearchInParam(searchItem As String, result As Boolean)
            Assert.AreEqual(result, VLPairIsInParam(searchItem, "C", param))
        End Sub
    End Class

    <TestClass>
    Public Class TestSearchSingleValueParamForPairs
        Private param As Object

        <TestInitialize>
        Sub Initialize()
            param = New SingleValueMockParam("a", "A")
        End Sub

        <DataRow("B", False, DisplayName:="Absent")>
        <DataRow("A", True, DisplayName:="Present")>
        <DataTestMethod>
        Sub TestMatchInSingleValueParam(label As String, result As Boolean)
            Assert.AreEqual(result, VLPairIsInParam("a", label, param))
        End Sub
    End Class
End Namespace
