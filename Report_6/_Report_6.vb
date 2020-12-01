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
    ''' <param name="searchStart">
    '''     The 1st part of the string to search for in the param.
    ''' </param>
    ''' <param name="searchEnd">
    '''     The 2nd part of the string to search for in the param. If this
    '''     string is empty, an empty String is returned. If the caller doesn't
    '''     want this option, they should use the other overload.
    '''     <see cref="Miscellaneous.LookupAndJoinMarksFromParam(String, String, Object)"/>
    ''' </param>
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
    '''     If "value" is passed, the param's Values are searched for matches
    '''     and its Label at the matching posisitions are returned.
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
            LookupAllMatchingParams(valueOrLabel, searchItem, param, "S"C)
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
            LookupAllMatchingParams("value", groupLearnerColumn, param, "S"C)
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
    '''     Calculate the average value of any values in a string
    ''' </summary>
    ''' <param name="vals">
    '''     A string containing 0 or more numeric values delimited by `, ` 
    ''' </param>
    ''' <returns>
    '''     Average of <c>vals</c> as a double, or <c>valIfBlank</c> if
    '''     <c>vals</c> is empty
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

''' <summary>
'''     Appends an object onto an array (without altering it) and returns
'''     the new array.
''' </summary>
''' <returns>
'''      An array resulting form appending <c>appendage</c> to
'''      <c>leftArry</c>.
''' </returns>
Public Function ArrayMerge(leftArray As Object(), _
        appendage As Object) As Object()
    Return ArrayMerge(leftArray, New Object() { appendage })
End Function

''' <summary>
'''     Prepends an object at the start of an array (without altering
'''     it) and returns the new array.
''' </summary>
''' <returns>
'''      An array resulting form appending <c>appendage</c> to
'''      <c>leftArry</c>.
''' </returns>
Public Function ArrayMerge(prependage As Object, _
        leftArray As Object()) As Object()
    Return ArrayMerge(New Object() { prependage }, leftArray)
End Function

''' <summary>
'''     Merges 2 arrays (without altering either) and returns the new
'''     array. Will attempt a narrowing conversion (e.g. String to
'''     Integer) if only one array is String(). If this converions
'''     fails for any item, the opposing, widening converion will be
'''     used. If neither arrays are of differing, non-string types,
'''     both will be converted to String.
''' </summary>
''' <param name="leftArray"> The first array </param>
''' <param name="rightArray"> The 2nd array </param>
''' <returns>
'''      An array resulting form appending <c>rightArray</c> to
'''      <c>leftArry</c>.
''' </returns>
Public Function ArrayMerge(leftArray As Object(), _
        rightArray As Object()) As Object()

    Dim leftType, rightType As Type

    If leftArray.Length = 0 Then Return rightArray
    leftType = leftArray(0).GetType()
    If rightArray.Length = 0 Then Return leftArray
    rightType = rightArray(0).GetType()

    leftArray = leftArray.Clone
    rightArray = rightArray.Clone

    If leftType.Equals(rightType) Then
        ' pass
    ElseIf leftType Is GetType(String) Then
        If Not TryCastArray(leftArray, rightType) _
            Then StringifyArray(rightArray)
    ElseIf rightType Is GetType(String) Then
        If Not TryCastArray(rightArray, leftType) _
            Then StringifyArray(leftArray)
    Else
        StringifyArray(leftArray)
        StringifyArray(rightArray)
    End If
    Return MergeSameTypeArrays(leftArray, rightArray)
End Function

Private Sub StringifyArray(inArray() As Object)
    For i As Long = 0 To UBound(inArray)
        inArray(i) = inArray(i).ToString()
    Next
End Sub

Private Function TryCastArray(array() As Object, _
                                destType As Type) As Boolean
    Dim parseMethod As Reflection.MethodInfo = destType.GetMethod( _
        "TryParse", New Type() { GetType(String), _
        destType.MakeByRefType } _
        )
    Dim sourceDestTuple(1) As Object

    If parseMethod Is Nothing Then Return False

    For i As Long = 0 To UBound(array)
        sourceDestTuple(0) = array(i)
        If Not parseMethod.Invoke(Nothing, sourceDestTuple) Then _
            Return False
        array(i) = sourceDestTuple(1)
    Next
    Return True
End Function

Private Function MergeSameTypeArrays(leftArray As Object(), _
                                        rightArray As Object()) As Object()

    Dim leftLength As Long = leftArray.Length
    Dim outArray(leftLength + UBound(rightArray)) As Object

    Array.Copy(leftArray, outArray, leftLength)
    Array.Copy(rightArray, 0, outArray, leftLength, rightArray.Length)
    Return outArray
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

    ''' <summary>
    '''     Find the 1st item equalling the searchItem
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 1, "E"C, False)
    End Function

    ''' <summary>
    '''     Find the Nth item equalling the searchItem
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                nthMatch As Integer) As Object
        Return LookupParam(valueOrLabel, searchItem, param, nthMatch, "E"C, False)
    End Function

    ''' <summary>
    '''     Find the 1st item that matches the searchItem using the specified
    '''     matchStrategy
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                matchStrategy As Char) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 1, matchStrategy, False)
    End Function

    ''' <summary>
    '''     Find the 1st item that equals the searchItem using Binary Search
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                useBinarySearch As Boolean) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 0, "E"C, useBinarySearch)
    End Function

    ''' <summary>
    '''     Find the nth item that matches the searchItem using the specified matchStrategy
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                nthMatch As Integer, _
                                matchStrategy As Char) As Object
        Return LookupParam(valueOrLabel, searchItem, param, nthMatch, matchStrategy, False)
    End Function

    ''' <summary>
    '''     Find the 1st item that matches the searchItem using the specified
    '''     match strategy and Binary Search
    ''' </summary>
    Public Function LookupParam(valueOrLabel As String, _
                                searchItem As Object, _
                                param As Object, _
                                matchStrategy As Char, _
                                useBinarySearch As Boolean) As Object
        Return LookupParam(valueOrLabel, searchItem, param, 0, matchStrategy, useBinarySearch)
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
    ''' <param name="useBinarySearch">
    '''     Pass true if the <c>param</c> is sorted by its <c>valueOrLabel</c>.
    '''     If so, binary-search (O(log(n))) will be used. (Otherwise O(n).)
    '''     Note that an SSRS query is gauranteed to be sorted by its first
    '''     field iff it's in its own query group (i.e. the field name is
    '''     repeated above the field.
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
                                    nthMatch As Integer, _
                                    matchStrategy As Char, _
                                    useBinarySearch As Boolean) _
                                    As Object
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
                If useBinarySearch Then ThrowInvalidBinarySearch("Contains")
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
                If useBinarySearch Then ThrowInvalidBinarySearch("Regular Expression")
                Return SearchUsingRegex( _
                    New Text.RegularExpressions.Regex(searchItem), _
                    searches, results, nthMatch, matchStrategy)
            Case "S"C ' Starts-with
                ThrowUnlessSearchIsString(searchItem, matchStrategy)
                If useBinarySearch Then Return BinSearchParam(searches,  _
                    results, searchItem, nthMatch, New StartsWithComparer)
                For i As Integer = 0 To param.Count - 1
                    If searches(i).StartsWith(searchItem) Then
                        foundCount += 1
                        If foundCount.Equals(nthMatch) Then
                            Return results(i)
                        End If
                    End If
                Next i
            Case Else ' Equals
                If useBinarySearch Then Return BinSearchParam( _
                    searches, results, searchItem, nthMatch, _
                    StringComparer.Create( _
                        New Globalization.CultureInfo("en-EN"), False _
                    ) _
                )
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

    Private Function BinSearchParam(searches As Object, _
                                    results As Object, _
                                    searchItem As Object, _
                                    nthMatch As Integer, _
                                    comparer As Collections.IComparer) As Object
        Dim randMatch As Integer = _
            Array.BinarySearch(searches, searchItem, comparer)
        Return NthMatchFromAMatch(searches, results, searchItem, nthMatch, _
                                  randMatch, comparer)
    End Function

    Private Function NthMatchFromAMatch(searches() As Object, _
                                        results() As Object, _
                                        searchItem As Object, _
                                        nthMatch As Integer, _
                                        matchIndex As Integer, _
                                        comparer As Collections.IComparer) As Object
    
        If matchIndex < 0 Then Return Nothing
        ' Return any match if the caller doesn't care:
        If nthMatch < 1 Then Return results(matchIndex)

        Dim firstIndex As Integer = matchIndex ' set as index of leftmost match:
        While firstIndex > 0 AndAlso _
                comparer.Compare(searches(firstIndex - 1), searchItem) = 0
            firstIndex -= 1
        End While
        Dim nthMatchingIndex As Integer = firstIndex + nthMatch - 1
        If nthMatchingIndex >= searches.Length OrElse _
            nthMatchingIndex > matchIndex AndAlso _
            comparer.Compare(searches(nthMatchingIndex), searchItem) <> 0 Then
            ' Out of bounds or-else not a match
            Return Nothing
        End If
        
        Return results(nthMatchingIndex)
    End Function

    Private Sub ThrowInvalidBinarySearch(strategyName As String)
        Throw New ArgumentException("You may not use binary-search with the '" _
                                    & strategyName & "' match-strategy. Use " _
                                    & "'Equals' or 'Starts-with' instead.")
    End Sub

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
            regex As Text.RegularExpressions.Regex, _
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

    Private Class StartsWithComparer
        Implements Collections.IComparer

        Public Function Compare(longString As Object, shortString As Object) _
                As Integer Implements Collections.IComparer.Compare
            If longString.StartsWith(shortString) Then Return 0

            Return longString.CompareTo(shortString)
        End Function
    End Class

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
                For i As Integer = 0 To param.Count - 1
                    If searches(i).StartsWith(searchItem) Then
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

    ''' <summary>
    '''     Counts all Values/Labels (as sepecified) in a param that equal the
    '''     <c>searchItem</c>. Beware that if the <c>searchItem</c> is an
    '''     integer from a query, it will be a long (int64), meaning that params
    '''     of type Integer (Int32) will not #equals them unless they are cast
    '''     to an Integer.
    ''' </summary>
    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object) As Integer
        Return CountMatchingParams(valueOrLabel, searchItem, param, "E"C)
    End Function

    Public Function CountMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object, _
                                        matchStrategy As Char) As Integer
        Dim searches As Object()
        Dim foundCount As Integer = 0

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
                For i As Integer = 0 To param.Count - 1
                    If searches(i).StartsWith(searchItem) Then
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
            Case "C"C ' Contains
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(search.Contains(searchItem), 1, 0)
            Case "S"C ' Starts-with
                ThrowIfMatchStrategyTypeConflict({search}, searchItem, matchStrategy)
                Return IIf(search.StartsWith(searchItem), 1, 0)
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

    ''' <summary>
    '''     Tests whether the parameter has a position that matches both the
    '''     value and the label supplied.
    ''' </summary>
    ''' <param name="value">
    '''     The object being searched for in the parameter's Values
    ''' </param>
    ''' <param name="label">
    '''     The String being searched for in the parameter's Labels
    ''' </param>
    ''' <param name="param">
    '''     An SSRS parameter containing both Values and Labels.
    '''     A single-value param is acceptable, and any type is fine.
    ''' </param>
    ''' <returns>
    '''     True if matching value/label pair is found; false otherwise
    ''' </returns>
    Public Function VLPairIsInParam(value As Object, label As String, _
                                    param As Object) As Boolean
        If Not param.IsMultiValue Then Return param.Value.Equals(value) _
            AndAlso param.Label.Equals(label)

        Dim values As Object() = param.Value
        ' Search Values first because they're more likely to be unique
        For i As Integer = 0 To param.Count - 1
            If values(i).Equals(value) AndAlso _
                param.Label(i).Equals(label) Then Return True
        Next
        Return False
    End Function

'C:\USERS\ZAC\DOCUMENTS\PROJECTS\SSRS CODE\UTILITIES\PARAM_HELPERS.VB
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
