''' <summary>
    ''' If the subject is science or double science, the relevant group
    ''' is found carefully by ensuring that it is the same discipline
    ''' (e.g. discipline `Ph` matches between `11Ph/Gauss` and
    ''' `10Ph/Euler`). If the subject is not a generic science, a normal
    ''' lookup will be used.
''' </summary>
''' <param name="param">
    ''' An SSRS parameter in this form:
    ''' Value: [learnerID][delimeter][subjectCode]
    ''' Label: [groupCode]
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