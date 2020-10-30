    ''' <summary>
    '''     Sanatises a mark. Trims unnecessary decimal points and zeros from
    '''     floats in Strings, or else, if there are points, appends them to
    '''     the mark, with a `#` seperator
    ''' </summary>
    ''' <param name="mark">
    '''     A mark, stored as a String
    ''' </param>
    ''' <param name="points">
    '''     Points corresponding to a mark's position on a gradescale. If there
    '''     is no relevant gradescale, <c>points</c> will be blank
    ''' </param>
    ''' <returns>
    '''     If the mark is a float, e.g. "12.00000" => "12"
    '''     Else if points aren't blank, e.g. "A*", "20" => "A*#20"
    '''     ELse just the mark, unchanged
    ''' </returns>
    Public Function CleanValue(mark As String, points As String) As String
        If mark.Contains(".0") Then
            Return CStr(CInt(mark))
        End If
        If Not points.Equals("") Then
            Return mark & "#" & points
        End If
        Return mark
    End Function

    ''' <summary>
    '''     Finds all Values/Labels (as sepecified) in a param that start with
    '''     <c>searchStart + searchEnd</c>, and joins the corresponding
    '''     Labels/Values.
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisitions are returned.
    '''     If "label" passed, searches the Labels and returns the Values.
    ''' </param>
    ''' <param name="searchStart">
    '''     The 1st part of the string to search for in the param.
    ''' </param>
    ''' <param name="searchEnd">
    '''     The 2nd part of the string to search for in the param. If this
    '''     string is empty, an empty String is returned. If the caller doesn't
    '''     want this option, they should use the other overloaded.
    '''     <see cref="Miscellaneous.LookupAndJoinMarksFromParam(String, String, Object)"/>
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels. A single-value
    '''     param is acceptable, but it must have Strings in the side being
    '''     searched in.
    ''' </param>
    ''' <returns>
    '''     The Labels/Values in the same positions in the param as the
    '''     Values/Labels that matched, but joined into a ", " delimited String.
    '''     (If none matched, the string is empty.)
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupAndJoinMarksFromParam(valueOrLabel As String, _
                                                searchStart As String, _
                                                searchEnd As String, _
                                                param As Object) As String
        If searchEnd = "" Then
            Return ""
        End If
        Return LookupAndJoinMarksFromParam(valueOrLabel, _
                                           searchStart & searchEnd, _
                                           param)
    End Function

    ''' <summary>
    '''     Finding all Values/Labels (as sepecified) in a param that start with
    '''     the searchItem, and joins the corresponding Labels/Values.
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisitions are returned.
    '''     If "label" passed, searches the Labels and returns the Values.
    ''' </param>
    ''' <param name="searchItem">The string to search for in the param.</param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels. A single-value
    '''     param is acceptable, but it must have Strings in the side being
    '''     searched in.
    ''' </param>
    ''' <returns>
    '''     The Labels/Values in the same positions in the param as the
    '''     Values/Labels that matched, but joined into a ", " delimited String.
    '''     (If none matched, the string is empty.)
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupAndJoinMarksFromParam(valueOrLabel As String, _
                                                searchItem As String, _
                                                param As Object) As String
        Dim results() As Object = _
            LookupAllMatchingParams(valueOrLabel, searchItem, param, "S")
        Select Case results.Length
            Case 0
                Return ""
            Case 1
                Return results(0)
            Case Else
                LookupAndJoinMarksFromParam = results(0) & ", "
                For i As Integer = 1 To results.Length - 1
                    If results(i) <> results(0) Then
                        LookupAndJoinMarksFromParam += results(i) & ", "
                    End If
                Next
                Return Strings.Left(LookupAndJoinMarksFromParam, _
                                    LookupAndJoinMarksFromParam.Length - 2)
        End Select
    End Function
    
    ''' <summary>
    '''     Retrieves all grades from a column, joining the grades in a comma
    '''     -delimeted list
    ''' </summary>
    ''' <param name="groupLearnerColumn">
    '''     IDs for a group, a learner (in the group) and a column, joined in
    '''     this format: group|learner#column
    ''' </param>
    ''' <param name="param">
    '''     A parameter containing group|learner#column[~anythingUnique] in its
    '''     values, and `grades#points` (or just `grades`) in its values
    ''' </param>
    ''' <returns>
    '''     A comma-delimited String of a strudent's grades in that column, in
    '''     that group. e.g. "A*, A, A"
    ''' </returns>
    Public Function LookupGradesFromParam(groupLearnerColumn As String, _
                                          param As Object) As String
        Return LookupGradesFromParam(groupLearnerColumn, param, False)
    End Function
    
    ''' <summary>
    '''     Use this version for early return if the column param is empty
    ''' </summary>
    ''' <param name="groupLearner">
    '''     IDs for a group, a learner (in the group) and a column, joined in
    '''     this format: group|learner#
    ''' </param>
    ''' <param name="column">
    '''     ID of a column. If this is empty, the method will return blank.
    ''' </param>
    ''' <param name="param">
    '''     A parameter containing group|learner#column[~anythingUnique] in its
    '''     values, and `grades#points` (or just `grades`) in its values
    ''' </param>
    ''' <returns>
    '''     A comma-delimited String of a strudent's grades in that column, in
    '''     that group. e.g. "A*, A, A"
    ''' </returns>
    Public Function LookupGradesFromParam(groupLearner As String, _
                                          column As String, _
                                          param As Object) As String
        If column Is Nothing OrElse column = "" Then Return ""
        Return LookupGradesFromParam(groupLearner & Column, param, False)
    End Function

    ''' <summary>
    '''     Retrieves all grades from a column, joining the grades in a comma
    '''     -delimeted list, and, if appendPoints is True, appends a comma
    '''     -delimeted list of the corresponding points, after a `#`
    ''' </summary>
    ''' <param name="groupLearnerColumn">
    '''     IDs for a group, a learner (in the group) and a column, joined in
    '''     this format: group|learner#column
    ''' </param>
    ''' <param name="param">
    '''     A parameter containing group|learner#column(~anythingUnique) in its
    '''     values, and grades#points in its values
    ''' </param>
    ''' <param name="appendPoints">
    '''     If True, appends a comma-delimited string of the points for each
    '''     looked-up grade
    ''' </param>
    ''' <returns>
    '''     A comma-delimited String of a strudent's grades in that column, in
    '''     that group. Possibly with the mean points afterward, seperated by
    '''     a `#`. e.g. "A*, B, B", or "A*, B, B#9" with appendPoints
    ''' </returns>
    Public Function LookupGradesFromParam(groupLearnerColumn As String, _
                                          param As Object, _
                                          appendPoints As Boolean) As String
        Dim results() As Object = _
            LookupAllMatchingParams("value", groupLearnerColumn, param, "S")
        Dim gradePointPair() As String
        Dim grades As String = ""
        Dim points As String = ""
        Dim uniqueGradeList As new System.Collections.Generic.List(Of String)
'       Concatenate only unique grades and points
        For Each result As String In results
            Dim include As Boolean = True
            gradePointPair = result.Split("#")
            For Each uniqueGrade As String In uniqueGradeList
                If uniqueGrade = gradePointPair(0) Then
                    include = False
                    Exit For
                End If
            Next uniqueGrade
            If include Then
                uniqueGradeList.Add(gradePointPair(0))
                grades += gradePointPair(0) & ", "
                If gradePointPair.Length > 1 And appendPoints Then
                    points += gradePointPair(1) & ", "
                End If
            End If
        Next
'       Trim dangling delimeters
        If grades.Length > 2 Then
            grades = Left(grades, grades.Length - 2)
        End If
        If appendPoints And points.Length > 2 Then
            Return grades & "#" & _
                EffectiveMark(Left(points, points.Length - 2), 0)
        End If
        Return grades
    End Function

    ''' <summary>
    '''     Calculate the average value of a series of 0 or more values in a string
    ''' </summary>
    ''' <param name="vals">
    '''     A string containing 0 or more numeric values delimited by `, ` 
    ''' </param>
    ''' <returns>
    '''     The average of <c>vals</c> as a double, or 40.0 if <c>vals</c> is empty
    ''' </returns>
    Public Function EffectiveMark(vals As String, _
                                  Optional valIfBlank As Double = 40) As Double
        Dim current As Double
        Dim sum As Double = 0
        Dim count As Integer = 0

        If vals = "" Then
            Return valIfBlank
        End If
        For Each val As String In Split(vals, ", ")
            If Not Double.TryParse(val, current) Then
                Return valIfBlank
            End If
            sum += current
            count += 1
        Next val
        Return sum / count
    End Function
