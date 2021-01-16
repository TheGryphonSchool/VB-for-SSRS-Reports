Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SSRSCode.SSRSCode.Mark12

Namespace TestSSRSCode
    <TestClass>
    Public Class TestMark12
        Private param As Object

        <TestInitialize>
        Sub Initialize()
            param = New MockParam( _
                {"1000Sc", "1000Sc", "5555Fr", "9999ScDouble", "9999ScDouble"}, _
                {"11Bi/Darwin", "11Sc/Darwin", "9Fr/deGualle", "10Ch/Curie", "10Ph/Curie"} _
            )
        End Sub

        <DataRow("1000", "Sc", "7Ph/Gauss", Nothing, _
                 DisplayName:="No matches should return blank")>
        <DataRow("5555", "Fr", "9Fr/Hugo", "9Fr/deGualle", _
                 DisplayName:="Non generic-science should match as normal")>
        <DataRow("1000", "Sc", "11Bi/Fisher", "11Bi/Darwin", _
                 DisplayName:="Finds the matching ScDouble group")>
        <DataRow("9999", "ScDouble", "10Ph/Newton", "10Ph/Curie", _
                 DisplayName:="Finds the matching Sc group even if it's not first")>
        <DataTestMethod>
        Sub TestLookupGroupScientifically(learnerID As String, _
                                          subjectCode As String, _
                                          groupCode As String, _
                                          result As String)
            Assert.AreEqual(result, LookupGroupScientifically( _
                            learnerID, subjectCode, groupCode, param))
        End Sub
    End Class
End Namespace
