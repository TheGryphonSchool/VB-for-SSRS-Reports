    ' Dependent on utilities/param_helpers.vb
    ' It must be combined if this file is used

    ''' <summary>
    '''     Use param like a dict, finding the 1st item equalling the searchItem
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisition is returned.
    '''     If "label" passed, searches the Labels and returns one of the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <returns>
    '''     <para>
    '''         If a match is found: The label/value in the same position in the
    '''         param as the value/label that matched.
    '''     </para>
    '''     <para>If a match is not found: <c>Nothing</c></para>
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 1, "E"C)
    End Function

    ''' <summary>
    '''     Use param like a dict, finding the Nth item equalling the searchItem
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the last matching posisition is returned.
    '''     If "label" passed, searches the Labels and returns one of the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <param name="nthMatch">
    '''     When <c>nthMatch</c> matches are found, the value/label (as appropriate
    '''     ) at the <c>nthMatch</c> mathing position is returned. If there are
    '''     fewer than <c>nthMatch</c> matches, <c>Nothing</c> is returned.
    ''' </param>
    ''' <returns>
    '''     <para>
    '''         If a match is found: The label/value in the same position in the
    '''         param as the value/label that matched.
    '''     </para>
    '''     <para>If a match is not found: <c>Nothing</c></para>
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                nthMatch As Integer) As Object
        Return LookupParam(valueOrLabel, searchItem, param, nthMatch, "E"C)
    End Function

    ''' <summary>
    '''     Use param like a dict, finding the 1st item that matches the searchItem
    '''     using the specified matchStrategy
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisition is returned.
    '''     If "label" passed, searches the Labels and returns one of the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <param name="matchStrategy">
    '''      A character denoting the match-strategy; one of:
    '''     <list type="bullet">
    '''         <item><term>E</term><description>Equals</description></item>
    '''         <item><term>S</term><description>Starts with</description></item>
    '''         <item><term>C</term><description>Contains</description></item>
    '''         <item><term>R</term><description>
    '''             String interpretable as a Regular Expression
    '''         </description></item>
    '''     </list>
    ''' </param>
    ''' <returns>
    '''     <para>
    '''         If a match is found: The label/value in the same position in the
    '''         param as the value/label that matched.
    '''     </para>
    '''     <para>If a match is not found: Nothing</para>
    ''' </returns>
    ''' </summary>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                matchStrategy As Char) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 1, matchStrategy)
    End Function

    ''' <summary>
    '''     Use param like a dict, finding the Nth item that matches the searchItem
    '''     using the specified matchStrategy
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the last matching posisition is returned.
    '''     If "label" passed, searches the Labels and returns one of the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <param name="nthMatch">
    '''     When <c>nthMatch</c> matches are found, the value/label (as appropriate
    '''     ) at the <c>nthMatch</c> mathing position is returned. If there are
    '''     fewer than <c>nthMatch</c> matches, <c>Nothing</c> is returned.
    ''' </param>
    ''' <param name="matchStrategy">
    '''      A character denoting the match-strategy; one of:
    '''     <list type="bullet">
    '''         <item><term>E</term><description>Equals</description></item>
    '''         <item><term>S</term><description>Starts with</description></item>
    '''         <item><term>C</term><description>Contains</description></item>
    '''         <item><term>R</term><description>
    '''             String interpretable as a Regular Expression
    '''         </description></item>
    '''     </list>
    ''' </param>
    ''' <returns>
    '''     <para>
    '''         If a match is found: The label/value in the same position in the
    '''         param as the value/label that matched.
    '''     </para>
    '''     <para>If a match is not found: <c>Nothing</c></para>
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Private Function LookupParam(valueOrLabel As String, _
                                 searchItem As Object, _
                                 param As Object, _
                                 nthMatch As Integer, _
                                 matchStrategy As Char) As Object
        Dim searches As Object()
        Dim results As Object()
        Dim foundCount As Integer = 0

        valueOrLabel = valueOrLabel.ToLower()

        If Not param.IsMultiValue Then
            If valueOrLabel = "label" Then
                searches = {param.Label}
                results = {param.Value}
            Else
                searches = {param.Value}
                results = {param.Label}
            End If
        Else
            searches = IIf(valueOrLabel = "value", param.Value, param.Label)
            results = IIf(valueOrLabel = "value", param.Label, param.Value)
        End If

        If searches.Length = 0 Then
            ' This is impossible for multivalue params in the current SSRS version
            Return Nothing
        End If

        Select Case matchStrategy
            Case "C"C ' Contains
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                For i As Integer = 0 To param.Count - 1
                    If searches(i).Contains(searchItem) Then
                        foundCount += 1
                        If foundCount.Equals(nthMatch) Then
                            Return results(i)
                        End If
                    End If
                Next i
            Case "R"C ' Regular Expression
                ThrowUnlessSearchIsString(searchItem, matchStrategy)
                Return SearchUsingRegex( _
                    New System.Text.RegularExpressions.Regex(searchItem), _
                    searches, results, nthMatch, matchStrategy)
            Case "S"C ' Starts-with
                ThrowUnlessSearchIsString(searchItem, matchStrategy)
                Return SearchUsingRegex(StartsWithRegex(searchItem), searches, _
                                        results, nthMatch, matchStrategy)
            Case Else ' Equals
                For i As Integer = 0 To param.Count - 1
                    If searchItem.Equals(searches(i)) Then
                        foundCount += 1
                        If foundCount.Equals(nthMatch) Then
                            Return results(i)
                        End If
                    End If
                Next i
        End Select

        Return Nothing ' searchItem was not found in parameter
    End Function

    ''' <summary>
    '''     Use a regular expression to look through an array of <C>searches</C>
    '''     until <C>nthMatch</C> matches are found. The element of the
    '''     <C>results</C> array at the same index is returned.
    ''' </summary>
    ''' <param name="regex">A regular expression, Regex object</param>
    ''' <param name="searches">
    '''     An array of Strings to search in. Must be the same length as the
    '''     <C>results</C> array
    ''' </param>
    ''' <param name="results">
    '''     An array of objects, one of which will be returned. Must be the same
    '''     length as the <C>results</C> array.
    ''' </param>
    ''' <param name="nthMatch">
    '''     The number of matches that must be found to return a result
    ''' </param>
    ''' <returns>
    '''     If <C>nthMatch</C> matches are found, returns the object in
    '''     <c>results</c> at the same position as the last match in
    '''     <C>searches</C>. Else, returns Nothing.
    ''' </returns>
    Private Function SearchUsingRegex( _
            regex As System.Text.RegularExpressions.Regex, _
            searches As Object(), _
            results As Object(), _
            nthMatch As Integer, _
            matchStrategy As Char) As Object
        Dim foundCount As Integer
        ThrowUnlessSearchesAreStrings(searches, matchStrategy)
        For i As Integer = 0 To searches.Length - 1
            If regex.IsMatch(searches(i)) Then
                foundCount += 1
                If foundCount.Equals(nthMatch) Then
                    Return results(i)
                End If
            End If
        Next i
        Return Nothing
    End Function

    ''' <summary>
    '''     Get the Nth Value from an SSRS parameter
    ''' </summary>
    ''' <param name="number">
    '''     The position in the param from which the Value/Label will be returned
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter. It may not be a single-value param, but it may have
    '''     any type, and it may have only Values (not Labels).
    ''' </param>
    Public Function LookupNthParam(number As Integer, param As Object) As Object
        Return LookupNthParam("value", number, param)
    End Function

    ''' <summary>
    '''     Get the Nth Value or Label from an SSRS parameter
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Nth Value is returned.
    '''     If "label" is passed, the param's Nth Label is returned.
    ''' </param>
    ''' <param name="number">
    '''     The position in the param from which the Value/Label will be returned
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter. It may not be a single-value param, but it may have
    '''     any type, and it may have only Values (not Labels).
    ''' </param>
    Public Function LookupNthParam(valueOrLabel As String, _
                                   number As Integer, _
                                   param As Object) As Object
        Dim results As Object() = _
            IIf(valueOrLabel.ToLower() = "value", param.Value, param.Label)
        If number <= param.Count Then
            Return results(number - 1)
        End If
        Return Nothing 'if parameter doesn't have that number of items
    End Function

    ''' <summary>
    '''     Use param like a dict, finding all items that equals the searchItem
    '''     using the specified matchStrategy. Return them all in an array
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisitions are returned.
    '''     If "label" passed, searches the Labels and returns the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <returns>
    '''     An array of the labels/values in the same positions in the param as the
    '''     values/labels that matched. (If none matched, the array is empty.)
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupAllMatchingParams(valueOrLabel As String, _
                                            searchItem As Object, _
                                            param As Object) As Object()
        Return LookupAllMatchingParams(valueOrLabel, searchItem, param, "E"C)
    End Function

    ''' <summary>
    '''     Use param like a dict, finding all items that match the searchItem.
    '''     Return them all in an array
    ''' </summary>
    ''' <param name="valueOrLabel">
    '''     Either the word 'value' or 'label' as a string (using any case).
    '''     If "value" is passed, the param's Values are searched for matches and
    '''     the its Label at the matching posisitions are returned.
    '''     If "label" passed, searches the Labels and returns one of the Values.
    ''' </param>
    ''' <param name="searchItem">
    '''     The thing being searched for in the param. This is expected to be the
    '''     same type as the type as the param's values.
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <returns>
    '''     An array of the labels/values in the same positions in the param as the
    '''     values/labels that matched. (If none matched, the array is empty.)
    ''' </returns>
    ''' <exception cref="System.ArgumentException">
    '''     Thrown if a 'contains' or 'starts-with' match-strategy is selected, but
    '''     either the searchItem or the param's values/labels (whichever is being
    '''     searched) is not a String.
    ''' </exception> 
    Public Function LookupAllMatchingParams(valueOrLabel As String, _
                                            searchItem As Object, _
                                            param As Object, _
                                            matchStrategy As Char) As Object()
        Dim searches As Object()
        Dim results As Object()
        Dim finds As New System.Collections.Generic.List(Of Object)

        valueOrLabel = valueOrLabel.ToLower()

        If Not param.IsMultiValue Then
            If valueOrLabel = "label" Then
                searches = {param.Label}
                results = {param.Value}
            Else
                searches = {param.Value}
                results = {param.Label}
            End If
        Else
            searches = IIf(valueOrLabel = "value", param.Value, param.Label)
            results = IIf(valueOrLabel = "value", param.Label, param.Value)
        End If

        If searches.Length = 0 Then
            ' This is impossible for multivalue params in the current SSRS version
            Return {}
        End If

        Select Case matchStrategy
            Case "C"C ' Contains
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                For i As Integer = 0 To param.Count - 1
                    If searches(i).Contains(searchItem) Then
                        finds.Add(results(i))
                    End If
                Next i
            Case "S"C ' Starts-with
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                Dim regexForStartsWith As System.Text.RegularExpressions.Regex = _
                StartsWithRegex(searchItem)
                For i As Integer = 0 To param.Count - 1
                    If regexForStartsWith.IsMatch(searches(i)) Then
                        finds.Add(results(i))
                    End If
                Next i
            Case Else ' Equals
                For i As Integer = 0 To param.Count - 1
                    If searchItem.Equals(searches(i)) Then
                        finds.Add(results(i))
                    End If
                Next i
        End Select

        Return finds.ToArray()
    End Function
