Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions

Friend Class SSRSDeployer
    Private ReadOnly outputPath As String
    Private ReadOnly modulePath As String
    Private ReadOnly allReportDependencies As List(Of ReportDependencies)

    Public Sub New(deployPath As String,
                   dependencyConfigJson As String)
        outputPath = Path.Combine(deployPath, "output")
        modulePath = GetSSRScodePath(deployPath)
        Dim jsonOptions = New JsonSerializerOptions With {
            .AllowTrailingCommas = True,
            .ReadCommentHandling = JsonCommentHandling.Skip,
            .PropertyNameCaseInsensitive = True
        }
        allReportDependencies =
            JsonSerializer.Deserialize(Of List(Of ReportDependencies)) _
                                      (dependencyConfigJson, jsonOptions)
        InitialiseReportsForWriting()
    End Sub

    Public Sub Deploy()
        ' For each .vb file in the moduleDir, if there are reportFiles
        '   requesting it, transcribe the contents of the ModuleFile to each:
        Dim moduleName As String
        Dim requestingReports As List(Of ReportDependencies)
        For Each modulePath As String In GetModules()
            moduleName = Path.GetFileNameWithoutExtension(modulePath)
            requestingReports = allReportDependencies.Where(
                Function(report) report.NeedModule(moduleName)).ToList()
            If Not requestingReports.Any Then Continue For

            ' Transcribe the module's contents to each requesting report file
            Try
                Using moduleReader As StreamReader =
                        New StreamReader(modulePath)
                    FlattenAndCopyFileContents(moduleReader, requestingReports)
                End Using
            Catch Ex As Exception
                Console.Out.WriteLine("Could not copy contents of " & modulePath)
            End Try
        Next
        allReportDependencies.ForEach(Sub(report) report.Writer.Close())
    End Sub

    Private Sub FlattenAndCopyFileContents(
                    moduleReader As StreamReader,
                    requestingReports As List(Of ReportDependencies))
        Dim flatContents As StringBuilder = New StringBuilder()
        Dim nsIndent As Integer = 0
        Dim modIndent As Integer = 0
        Dim commentRegex As Regex = New Regex("^\s*'")
        Dim indentRegexes As Regex() = {
           New Regex("^"),
           New Regex("^\s{0,4}"),
           New Regex("^\s{0,8}")
           }
        Dim line As String
        While moduleReader.Peek() >= 0
            line = moduleReader.ReadLine()
            If commentRegex.IsMatch(line) Then Continue While
            If nsIndent = 0 Then
                If line.Contains("Namespace") Then
                    nsIndent = 1
                    Continue While
                End If
            ElseIf line.Contains("End Namespace") Then
                nsIndent = 0
                Continue While
            End If
            If modIndent = 0 Then
                If line.Contains("Module") Then
                    modIndent = 1
                    Continue While
                End If
            ElseIf line.Contains("End Module") Then
                modIndent = 0
                Continue While
            End If

            flatContents.AppendLine(
                indentRegexes(nsIndent + modIndent).Replace(line, ""))
        End While
        For Each requestingReport In requestingReports
            Try
                requestingReport.Writer.Write(flatContents)
            Catch Ex As Exception
                Console.Out.WriteLine(
                        "Could not copy contents of " & requestingReport.Name)
            End Try
        Next
    End Sub

    Private Sub InitialiseReportsForWriting()
        ' For each reportFile, initialise a StreamWriter and empty the contents
        Dim reportPath As String
        For Each report As ReportDependencies In allReportDependencies
            reportPath = Path.Combine(outputPath, report.Name & ".vb")
            Try
                File.WriteAllText(reportPath, "")
                report.Writer = New StreamWriter(reportPath)
            Catch ex As Exception
                Console.Out.WriteLine("Could not open the file " & reportPath)
                allReportDependencies.Remove(report)
            End Try
        Next
    End Sub

    Private Function GetModules() As IEnumerable(Of String)
        Return Directory.EnumerateFiles(modulePath, "*.vb")
    End Function

    Private Function GetSSRScodePath(deployPath As String) As String
        Return Path.Join(Path.GetDirectoryName(deployPath), "SSRSCode")
    End Function
End Class
