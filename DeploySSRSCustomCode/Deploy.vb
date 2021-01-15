Imports System.IO

Module Deploy
    Sub Main()
        Dim originalDirPath As String = Directory.GetParent(
            Directory.GetCurrentDirectory()).Parent.Parent.FullName
        Dim dependenciesPath As String = Path.Join(originalDirPath,
                                                   "ReportDependencies.json")
        Dim ssrsDeployer As SSRSDeployer

        Try
            Dim dependenciesFile As StreamReader = _
                New StreamReader(dependenciesPath)
            ssrsDeployer = New SSRSDeployer(originalDirPath, _
                                            dependenciesFile.ReadToEnd())
        Catch Ex As Exception
            Throw New IOException("Tried to open 'ReportDependencies.json', " &
                "but got this error: " & Ex.Message)
        End Try
        ssrsDeployer.Deploy()
    End Sub
End Module
