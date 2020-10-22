    Public Function roundIfFloat(float As String) As String
        If Not float.Contains(".0") Then
            Return float
        End If
        Return CStr(CInt(float))
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
            Return 0
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
