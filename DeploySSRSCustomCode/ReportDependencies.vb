Imports System.IO
Imports System.Text.RegularExpressions

Friend Class ReportDependencies
    Public Property Name As String
    Public Property Dependencies As List(Of String)
    Public Property Writer As StreamWriter
    Private Property Initialised As Boolean = False
    Private Property DependencyRegexes As List(Of Regex) = New List(Of Regex)

    Public Sub WriteLine(line As String)
        Writer.WriteLine(line)
    End Sub

    Public Function NeedModule(moduleName As String) As Boolean
        If Not Initialised Then Initialise()
        Return DependencyRegexes.Any(Function(regex) regex.IsMatch(moduleName))
    End Function

    Private Sub Initialise()
        Dependencies.ForEach(Sub(dependency)
                                 DependencyRegexes.Add(New Regex(dependency))
                             End Sub)
        Initialised = True
    End Sub
End Class
