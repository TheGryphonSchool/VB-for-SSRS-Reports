Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports DebugAssemblies.SSRSCode.Arrays

Namespace TestSSRSCode
    <TestClass>
    Public Class TestArrayUtils

        <DataRow(New Object() {1, 2}, 3, DisplayName:="Sums Integers")>
        <DataRow(New Object() {1.0, 2.0}, 3, DisplayName:="Sums Floats")>
        <DataRow(New Object() {1.0, 2.5}, 1, _
                 DisplayName:="Ignores non-integer Floats")>
        <DataRow(New Object() {"1", "2"}, 3, _
                 DisplayName:="Sums integers in Strings")>
        <DataTestMethod>
        Public Sub TestSumsArrays(arry As Object(), _
                                  expected As Integer)
            Assert.AreEqual(expected, SumArray(arry))
        End Sub

        <DataRow(New Object() {1, 2, 1}, _
                 DisplayName:="Removes duplicate Integers")>
        <DataTestMethod>
        Public Sub TestRemoveDuplicates(arry As Object())
            Dim expected As Object() =  New Object() {1, 2}
            CollectionAssert.AreEqual(expected, RemoveDuplicates(arry))
        End Sub

        <DataRow( _
            New Object() {0, 1, 2}, _
            New Object() {0, 1}, _
            New Object() {2} _
            )>
        <DataRow( _
            New Object() {0, 1, 2}, _
            New Object() {"0", "1"}, _
            New Object() {2} _
            )>
        <DataRow( _
            New Object() {0, 1, 2}, _
            New Object() {0}, _
            New Object() {"1", "2"} _
            )>
        <DataRow( _
            New Object() {"1", "2KG"}, _
            New Object() {1}, _
            New Object() {"2KG"} _
            )>
        <DataRow( _
            New Object() {"1", "0.2"}, _
            New Object() {1}, _
            New Object() {0.2} _
            )>
        <DataTestMethod>
        Sub TestMergesArrays(expected As Object(), _
                             left As Object(), right As Object())
            CollectionAssert.AreEqual(expected, ArrayMerge(left, right))
        End Sub

        Sub TestAppendsItemToArray()
            CollectionAssert.AreEqual(New Object() {1, 2}, _
                                      ArrayMerge(New Object() {1}, 2))
        End Sub

        Sub TestPrependsItemBeforeArray()
            CollectionAssert.AreEqual(New Object() {1, 2}, _
                                      ArrayMerge(1, New Object() {2}))
        End Sub

        <TestMethod>
        Sub TestParsesStringToDateAndMerges()
            Dim expectedDates As Object() = { _
                New DateTime(DateTime.MinValue.Ticks), _
                New DateTime(DateTime.MaxValue.Ticks) _
            }
            CollectionAssert.AreEqual(expectedDates, _
                                      ArrayMerge(New Object() {"01/01/0001"}, _
                                                 New Object() {expectedDates(1)}))
        End Sub
    End Class
End Namespace
