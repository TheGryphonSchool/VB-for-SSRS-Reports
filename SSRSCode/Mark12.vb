Namespace SSRSCode
    ''' <summary>
    '''     This module contains functions that are only expected to be useful
    '''     in SSRS reports related to the yearly Mark-12 reporting cycle.
    ''' </summary>
    Public Module Mark12

        ''' <summary>
        '''     If the subject is science or double science, the relevant group
        '''     is found carefully by ensuring that it is the same discipline
        '''     (e.g. discipline `Ph` matches between `11Ph/Gauss` and
        '''     `10Ph/Euler`). If the subject is not a generic science, a normal
        '''     lookup will be used.
        ''' </summary>
        ''' <param name="learnerID">
        '''     A learner ID, with a delimeter appended, if there one is used in
        '''     the parameter
        ''' </param>
        ''' <param name="param">
        '''     An SSRS parameter in this form:
        '''     Value: [learnerID][delimeter]?[subjectCode]
        '''     Label: [groupCode]
        '''     The parameter MUST be ordered by its values.
        ''' </param>
        ''' <param name="matchStrategy">
        '''      A character denoting the match-strategy; one of:
        '''     <list type="bullet">
        '''         <item><term>E</term><description>Equals</description></item>
        '''         <item><term>S</term><description>Starts with</description></item>
        '''         <item><term>C</term><description>Contains</description></item>
        '''     </list>
        ''' </param>
        ''' <returns>
        '''     The group code that a learner is currently in for a subject
        ''' </returns>
        Public Function LookupGroupScientifically(learnerID As String, _
                                                  subjectCode As String, _
                                                  groupCode As String, _
                                                  param As Object, _
                                                  Optional matchStrategy As Char = "E"C _
                                                 ) As Object
            If subjectCode <> "Sc" AndAlso subjectCode <> "ScDouble" Then
                Return LookupParam("value", learnerID & subjectCode, param, _
                                    matchStrategy, True)
            End If

            Dim subSubject As String = New Text.RegularExpressions. _
                Regex("^\d{1,2}(\w+)(?=\/)").Match(groupCode).Groups(1).Value
            Dim groupCodeRegex As Text.RegularExpressions.Regex = New _
                Text.RegularExpressions.Regex("^\d{1,2}" & subSubject & "\/")
            For Each potentialGroup As Object In _
                LookupAllMatchingParams("value", learnerID & subjectCode, _
                                        param, matchStrategy)
                If groupCodeRegex.IsMatch(potentialGroup.ToString()) Then Return potentialGroup
            Next potentialGroup
            Return Nothing
        End Function
    End Module
End Namespace
