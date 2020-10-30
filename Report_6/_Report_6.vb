'File produced by combining files using the Combine Files VScode extension
'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\REPORT_6\MISCELLANEOUS.VB
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

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\ARRAYS.VB
Public Function removeDuplicates(items As Object()) As Object()
    'maps original indicies to new indicies
    Dim index_map As New _
        System.Collections.Generic.Dictionary(Of Integer, Integer)
    Dim shift As Integer = 0
    Dim unique_items As New System.Collections.Generic.List(Of Object)
    Dim index As Integer = 0
    For Each item As Object In items
        If unique_items.contains(item) Then
            shift += 1
        Else
            unique_items.Add(item)
            index_map.Add(index, index - shift)
        End If
        index += 1
    Next item
    For Each old_index As Integer In index_map.Keys
        items(index_map(old_index)) = items(old_index)
    Next old_index
    ReDim Preserve items(index - 1 - shift)
    Return items
End Function

Public Function SumArray(nums() As Object) As Integer
    Dim num As Integer
	SumArray = 0
	For Each o As Object In nums
        If Integer.TryParse(o, num)
            SumArray += num
        End If
	Next o
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\COLOUR_SCALE.VB
Public Class ColourScale
    Private scale As New System.Collections.Generic.List(Of Integer())

    Public Sub New(first As String, _
                   second As String, _
                   Optional third As String = "", _
                   Optional fourth As String = "", _
                   Optional fifth As String = "")
        For Each nth As String In New String(4) {first, second, third, fourth, fifth}  
            If nth Is "" Then
                Exit For
            End If
            addToScale(nth)
        Next nth
    End Sub

    Public Function getColour(fraction As Double)
        Dim last_index As Integer = scale.Count - 1
        Dim start As Integer
        If fraction >= 1.0 Then
            Return mixTwoColours(1.0, last_index - 1)
        End If
        start = CInt(Math.Floor(fraction * last_index))
        Return mixTwoColours(fraction * last_index - start, start)
    End Function

    Private Sub addToScale(hexColour As String)
        Dim rgb(2) As Integer
        hexColour = hexColour.Replace("#", "")
        For i As Integer = 0 To 2
            rgb(i) = Convert.ToInt32(hexColour.Substring(i * 2, 2), 16)
        Next i
        scale.Add(rgb)
    End Sub

    Private Function mixTwoColours(fraction As Double, _
                                   start_index As Integer) As String
        Dim starts As Integer
        Dim ends As Integer
        mixTwoColours = "#"
        For i As Integer = 0 To 2
            starts = scale.Item(start_index)(i)
            ends = scale.Item(start_index + 1)(i)
            mixTwoColours += _
                Hex(CInt(starts + fraction * (ends - starts))).PadLeft(2,  "0")
        Next i
    End Function
End Class

Dim header_colour_scale As ColourScale

Public Function colourFromScale(fraction As Double, _
                                first As String, _
                                second As String, _
                                third As String) As String
    If header_colour_scale Is Nothing Then
        header_colour_scale = New ColourScale(first, second, third)
    End If
    Return header_colour_scale.getColour(fraction)
End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\LOOKUP_PARAMS.VB
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
        Return LookupParam(valueOrLabel, searchItem, param, 1, "E")
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
        Return LookupParam(valueOrLabel, searchItem, param, nthMatch, "E")
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
        Return LookupAllMatchingParams(valueOrLabel, searchItem, param, "E")
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

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\SEARCH_PARAMS.VB
    ' Dependent on utilities/param_helpers.vb
    ' It must be combined if this file is

    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object) As Integer
        Return CountMatchingParams(valueOrLabel, searchItem, param)
    End Function

    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object, _
                                        matchStrategy As Char) As Integer
        Dim searches As Object()
        Dim foundCount As Integer = 0
        Dim regexForStartsWith As System.Text.RegularExpressions.Regex

        valueOrLabel = valueOrLabel.ToLower()
        If Not param.IsMultiValue Then
            Return CountInSingleValueParam(valueOrLabel, searchItem, _
                                       param, matchStrategy)
        End If
        searches = IIf(valueOrLabel = "value", param.Value, param.Label)
        Select Case matchStrategy
            Case "C"C ' Contains
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                For i As Integer = 0 To param.Count - 1
                    If searches(i).Contains(searchItem) Then
                        foundCount += 1
                    End If
                Next i
            Case "S"C ' Starts-with
                ThrowIfMatchStrategyTypeConflict(searches, searchItem, matchStrategy)
                regexForStartsWith = StartsWithRegex(searchItem)
                For i As Integer = 0 To param.Count - 1
                    If regexForStartsWith.IsMatch(searches(i)) Then
                        foundCount += 1
                    End If
                Next i
            Case Else ' Equals
                For i As Integer = 0 To param.Count - 1
                    If searchItem.Equals(searches(i)) Then
                        foundCount += 1
                    End If
                Next i
        End Select
        Return foundCount
    End Function

    Private Function CountInSingleValueParam(valueOrLabel As String, _
                                             searchItem As String, _
                                             param As Object, _
                                             matchStrategy As Char) As Integer
        Dim search As Object

        search = IIf(valueOrLabel = "value", param.Value, param.Label)
        If Not TypeOf search Is String Then
            Throw New ArgumentException("The parameter must be a string")
        End If
        Select Case matchStrategy
            Case "C" ' Contains
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(search.Contains(searchItem), 1, 0)
            Case "S" ' Starts-with
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(StartsWithRegex(searchItem).IsMatch(search), 1, 0)
            Case Else ' Equals
                Return IIf(searchItem.Equals(search), 1, 0)
        End Select
    End Function

    Public Function IsInParam(valueOrLabel As String, _
                              searchItem As Object, _
                              param As Object) As Boolean
        Dim lookups As Object() = _
            IIf(valueOrLabel.ToLower() = "value", param.Value, param.Label)
        Return Array.IndexOf(lookups, searchItem) >= 0
    End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\PARAM_HELPERS.VB
    Private Function StartsWithRegex(start As String) As _
                                 System.Text.RegularExpressions.Regex
        Return New _
    System.Text.RegularExpressions.Regex("^" & EscapeRegexString(start))
    End Function

    Private Function EscapeRegexString(unescaped As String) As String
        ' Escape regex meta-characters in user-supplied string so that a regex can
        ' be built from the string that matches the supplied characters literally
        Dim escRgx As System.Text.RegularExpressions.Regex
        escRgx = New System.Text.RegularExpressions.Regex("[|^$.()?+*\[\]\\]")
        Return escRgx.Replace(unescaped, "\$&")
    End Function

    Private Sub ThrowIfMatchStrategyTypeConflict(searches As Object(), _
                                                 searchItem As Object, _
                                                 matchStrategy As Char)
        ThrowUnlessSearchIsString(searchItem, matchStrategy)
        ThrowUnlessSearchesAreStrings(searches, matchStrategy)
    End Sub

    Private Sub ThrowUnlessSearchIsString(searchItem As Object, _
                                          matchStrategy As Char)
        If TypeOf searchItem Is String Then Exit Sub
        Throw New ArgumentException(MatchStrategyExceptionMessage( _
            "The search item must be a string", matchStrategy))
    End Sub

    Private Sub ThrowUnlessSearchesAreStrings(searches As Object(), _
                                              matchStrategy As Char)
        If TypeOf searches(0) Is String Then Exit Sub
        Throw New ArgumentException(MatchStrategyExceptionMessage( _
            "The parameter must have string values", matchStrategy))
    End Sub

    Private Function MatchStrategyExceptionMessage(problemStatement As String, _
                                                   matchStrategy As Char) As String
        Dim strategyDescription As String
        Select Case matchStrategy
            Case "C"C
                strategyDescription = "'Contains' ('C')"
            Case "S"C
                strategyDescription = "'Starts-with' ('S')"
            Case Else
                strategyDescription = "'Regular Expression' ('R')"
        End Select
        Return problemStatement & " to use the match strategy " & _
            strategyDescription & _
            ". Omit the matchStrategy argument to use exact matching."
    End Function
