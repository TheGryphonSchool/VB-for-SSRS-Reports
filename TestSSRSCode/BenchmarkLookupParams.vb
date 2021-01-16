Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SSRSCode.SSRSCode.LookupParams

Namespace TestSSRSCode
    <TestClass>
    Public Class BenchmarkLookupParams
        Private ReadOnly param As MockParam

        ' Runs some informal benchmarks with its assertions. Will slow down test
        ' suite unless ignored.
        <Ignore>
        <TestMethod>
        Sub TestBinarySearchStringParamLookup()
            Dim values() As IComparable
            Dim labels() As String
            Dim listValues As New List(Of IComparable)
            Dim listLabels As New List(Of String)
            Dim caseCount As Integer = 10
            For i As Integer = 0 To caseCount
                listValues.Add(chr(asc("a") + i \ 4).ToString())
                listLabels.Add(CStr(i))
            Next i
            values = ListValues.ToArray()
            labels = ListLabels.ToArray()
            Dim strParam As MockParam = New MockParam(values, labels)
            Dim testCases(caseCount) As LookupParamTestCase
            For i As Integer = 0 To caseCount
                testCases(i) = _
                    New LookupParamTestCase(chr(asc("a") + i \ 4).ToString(), _
                                            "E"C, CStr(i), i Mod 4 + 1)
            Next
            Debug.Print("Binary Search (strings): " & TimeTestCases(testCases, strParam, true))
            Debug.Print("Linear Search (strings): " & TimeTestCases(testCases, strParam, false))
        End Sub

        ' Runs some informal benchmarks with its assertions. Will slow down test
        ' suite unless ignored.
        <Ignore>
        <TestMethod>
        Sub TestBinarySearchStartsWithParamLookup()
            Dim values() As IComparable
            Dim labels() As String
            Dim listValues As New List(Of IComparable)
            Dim listLabels As New List(Of String)
            Dim caseCount As Integer = 10
            For i As Integer = 0 To caseCount
                Dim letter As String = chr(asc("a"C) + i \ 4).ToString()
                listValues.Add(letter & letter)
                listLabels.Add(CStr(i))
            Next i
            values = ListValues.ToArray()
            labels = ListLabels.ToArray()
            Dim strParam As MockParam = New MockParam(values, labels)
            For i As Integer = 0 To caseCount
                Assert.AreEqual( _
                    CStr(i), _
                    LookupParam("value", chr(asc("a"C) + i \ 4), _
                                strParam, nthMatch:=i Mod 4 + 1, "S"C, true))
            Next
        End Sub

        ' Runs some informal benchmarks with its assertions. Will slow down test
        ' suite unless ignored.
        <Ignore>
        <TestMethod>
        Sub TestBinarySearchNumericalParamLookup()
            Dim values() As IComparable
            Dim labels() As String
            Dim listValues As New List(Of IComparable)
            Dim listLabels As New List(Of String)
            Dim caseCount As Integer = 10
            For i As Integer = 0 To caseCount
                listValues.Add(i \ 4)
                listLabels.Add(CStr(i))
            Next i
            values = ListValues.ToArray()
            labels = ListLabels.ToArray()
            Dim numParam As MockParam = New MockParam(values, labels)
            Dim testCases(caseCount) As LookupParamTestCase
            For i As Integer = 0 To caseCount
                testCases(i) = New LookupParamTestCase(i \ 4, "E"C, CStr(i), i Mod 4 + 1)
            Next
            Debug.Print("Binary Search: " & TimeTestCases(testCases, numParam, true))
            Debug.Print("Linear Search: " & TimeTestCases(testCases, numParam, false))
        End Sub

        ' Utility for timing test cases
        Private Function TimeTestCases(testCases() As LookupParamTestCase, _
                                       param As Object, binSearch As Boolean)
            Dim timer = New Stopwatch
            timer.Start()
            For Each testCase As LookupParamTestCase In testCases
                With testCase
                    Assert.AreEqual(.expected, _
                                    LookupParam("value", .searchItem, param, _
                                         .nthMatch, .matchStrategy, binSearch))
                End With
            Next
            timer.Stop()
            Return timer.ElapsedMilliseconds
        End Function
    End Class
End Namespace
