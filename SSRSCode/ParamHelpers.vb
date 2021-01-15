Namespace SSRSCode
    Module ParamHelpers
        Friend Sub ThrowUnlessStringifiable(searches As Object(), _
                                            searchItem As Object, _
                                            matchStrategy As Char)
            Dim varMessages = {{searchItem, "The search item"}, _
                {searches(0), "The items in the searched side of the parameter"}}
            For i As Integer = 0 To 1
                If IsNotStringifiable(varMessages(i, 0)) Then
                    Throw New ArgumentException(MatchStrategyExceptionMessage( _
                        varMessages(i, 1), matchStrategy, _
                        varMessages(i, 0).GetType().Name))
                End If
            Next i
        End Sub

        Private Function IsNotStringifiable(ssrsObject As Object) As Boolean
            ' Arrays and bools are the only SSRS data types that cannot sensibly
            ' have their string representations compared
            Return IsArray(ssrsObject) OrElse TypeOf ssrsObject Is Boolean
        End Function

        Private Function MatchStrategyExceptionMessage(problematicParam As String, _
                                                       matchStrategy As Char, _
                                                       badType As String) As String
            Dim strategyDescription As String
            Select Case matchStrategy
                Case "C"C
                    strategyDescription = "'Contains' ('C')"
                Case "S"C
                    strategyDescription = "'Starts-with' ('S')"
                Case Else
                    strategyDescription = "'Regular Expression' ('R')"
            End Select
            Return problematicParam & " cannot be of type " & badType & _
                " to use the match strategy " & strategyDescription & _
                ". Omit the matchStrategy argument to use exact matching."
        End Function
    End Module
End Namespace
