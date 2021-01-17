Imports System.IO

Module Deploy
    Sub Main()
        Dim projectDirPath As String =
            New Text.RegularExpressions.Regex("(?<=DeploySSRSCustomCode).*") _
            .Replace(Directory.GetCurrentDirectory(), "")
        Dim dependenciesPath As String = Path.Join(projectDirPath,
                                                   "ReportDependencies.json")
        Dim ssrsDeployer As SSRSDeployer

        Try
            Dim dependenciesFile As StreamReader =
                New StreamReader(dependenciesPath)
            ssrsDeployer = New SSRSDeployer(projectDirPath,
                                            dependenciesFile.ReadToEnd())
        Catch Ex As Exception
            Throw New IOException("Tried to open 'ReportDependencies.json', " &
                "but got this error: " & Ex.Message)
        End Try
        ssrsDeployer.Deploy()
    End Sub
End Module
