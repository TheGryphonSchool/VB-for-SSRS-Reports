Namespace SSRSCode
    ''' <summary>
    '''     A collection of pure functions that find and return items within
    '''     SSRS parameters, using various matching strategies.
    ''' </summary>
    ''' <remarks>
    '''     Where a function accepts a 'match-strategy' as a parameter, it
    '''     expects a character, which must be one of:
    '''     <list type="bullet">
    '''         <item>
    '''             <term><c>"E"C</c></term>
    '''             <description>
    '''                 'Equals': A match is found when the <c>searchItem</c>
    '''                 equals an item in (the searched side of) the SSRS
    '''                 parameter.
    '''             </description>
    '''         </item>
    '''         <item>
    '''             <term><c>"S"C</c></term>
    '''             <description>
    '''                 'Starts with': The <c>searchItem</c> parameter should be
    '''                 an object with a meaningful String representation\*. A
    '''                 match will be an object in (the searched side of) the
    '''                 SSRS parameter, with a meaningful String
    '''                 representation\*, that starts with** the
    '''                 <c>searchItem</c>.
    '''             </description>
    '''         </item>
    '''         <item>
    '''             <term><c>"C"C</c></term>
    '''             <description>
    '''                 'Contains': The <c>searchItem</c> parameter should be
    '''                 an object with a meaningful String representation\*. A
    '''                 match will be an object in (the searched side of) the
    '''                 SSRS parameter, with a meaningful String
    '''                 representation\*, that contains** the <c>searchItem</c>.
    '''             </description>
    '''         </item>
    '''         <item>
    '''             <term><c>"R"C</c></term>
    '''             <description>
    '''                 'Regular Expression': The <c>searchItem</c> parameter
    '''                 should be an object with a meaningful String
    '''                 representation\*, interpretable as a Regular Expression.
    '''                 A match will be an object in (the searched side of) the
    '''                 SSRS parameter, with a meaningful String
    '''                 representation\*, that the Regular Expression
    '''                 accepts. For guidance on Regular Expressions in VB,
    '''                 <see href="https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference">
    '''                     use this quick reference guide
    '''                 </see>.
    '''             </description>
    '''         </item>
    '''     </list>
    '''     <para>
    '''         \*The 'String representation' is the result of calling its
    '''         #ToString() method. This will not be meaningful if
    '''         <c>searchItem</c> is an array. Obviously if the
    '''         <c>searchItem</c> or the elements in (the searched side of) the
    '''         SSRS parameter are Strings, they are their own String
    '''         representations.
    '''     </para>
    '''     <para>
    '''         **Wildcards are not recognised by the 'starts-with' and
    '''         'contains' strategies; use the 'regex' match-strategy if you
    '''         need wildcards.
    '''     </para>
    ''' </remarks>
    Public Module LookupParams
        ' Dependent on utilities/param_helpers.vb

        ''' <summary>
        '''     Get the 'partner' of the 1st item equalling the
        '''     <c>searchItem</c>. For full parameter documentation, see the
        '''     delegate function 
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "b34", "c67"]
        ''' }
        ''' >>> LookupParam("label", "b34", param)
        ''' 3456
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object) As Object
            Return LookupParam(valueOrLabel, searchItem, param, 1, "E"C, False)
        End Function

        ''' <summary>
        '''     Get the 'partner' of the Nth item equalling the
        '''     <c>searchItem</c>. For full parameter documentation, see the
        '''     delegate function
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <param name="nthMatch"/>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "b34", "b34"]
        ''' }
        ''' >>> LookupParam("label", "b34", param, 2)
        ''' 6789
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object, _
                                    nthMatch As Integer) As Object
            Return LookupParam(valueOrLabel, searchItem, param, nthMatch, "E"C, _
                               False)
        End Function

        ''' <summary>
        '''     Get the 'partner' of the 1st item that matches the
        '''     <c>searchItem</c> using the specified <c>matchStrategy</c>. For
        '''     full parameter documentation, see the delegate function
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <param name="matchStrategy"/>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "b34", "c67"]
        ''' }
        ''' >>> LookupParam("label", "c67", param, "E"C)
        ''' 6789
        ''' >>> LookupParam("label", "b", param, "S"C)
        ''' 3456
        ''' >>> LookupParam("label", "6", param, "C"C)
        ''' 6789
        ''' >>> LookupParam("label", "1$", param, "R"C)
        ''' 01234
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object, _
                                    matchStrategy As Char) As Object
            Return LookupParam(valueOrLabel, searchItem, param, 1, _
                               matchStrategy, False)
        End Function

        ''' <summary>
        '''     Get the 'partner' of the 1st item that equals the
        '''     <c>searchItem</c> using Binary Search. For full parameter
        '''     documentation, see the delegate function
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <param name="useBinarySearch"/>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "b34", "c67"]
        ''' }
        ''' >>> LookupParam("label", "b34", param, True)
        ''' 3456
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object, _
                                    useBinarySearch As Boolean) As Object
            Return LookupParam(valueOrLabel, searchItem, param, 0, "E"C, _
                               useBinarySearch)
        End Function

        ''' <summary>
        '''     Get the 'partner' of the nth item that matches the
        '''     <c>searchItem</c> using the specified <c>matchStrategy</c>. For
        '''     full parameter documentation, see the delegate function
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <param name="nthMatch"/>
        ''' <param name="matchStrategy"/>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "a34", "c14"]
        ''' }
        ''' >>> LookupParam("label", "a", param, 2, "S"C)
        ''' 3456
        ''' >>> LookupParam("label", "1", param, 1, "C"C)
        ''' 0123
        ''' >>> LookupParam("label", "...", param, 3, "R"C)
        ''' 6789
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object, _
                                    nthMatch As Integer, _
                                    matchStrategy As Char) As Object
            Return LookupParam(valueOrLabel, searchItem, param, nthMatch, _
                               matchStrategy, False)
        End Function

        ''' <summary>
        '''     Get the 'partner' of the 1st item that matches the
        '''     <c>searchItem</c> using the specified match strategy and Binary
        '''     Search. For full parameter documentation, see the delegate
        '''     function
        '''     <see href="#lookupparam-valueorlabel-searchitem-param-nthmatch-matchstrategy-usebinarysearch-"/>
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        ''' <param name="matchStrategy"/>
        ''' <param name="useBinarySearch"/>
        ''' <remarks>
        '''     It is not possible to pass <c>useBinarySearch:=True</c> with
        '''     <c>matchStrategy:="C"C</c> or <c>matchStrategy:="R"C</c>,
        '''     because there is no ordering of the parameter that would allow
        '''     a binary search to run without skipping elements.
        ''' </remarks>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "a34", "c14"]
        ''' }
        ''' >>> LookupParam("label", "a34", param, "E"C, True)
        ''' 3456
        ''' >>> LookupParam("label", "a", param, "S"C, True)
        ''' 0123
        ''' >>> LookupParam("label", "1", param, "C"C, True)
        ''' InvalidBinarySearchException
        ''' >>> LookupParam("label", "...", param, "R"C, True)
        ''' InvalidBinarySearchException
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                    searchItem As Object, _
                                    param As Object, _
                                    matchStrategy As Char, _
                                    useBinarySearch As Boolean) As Object
            Return LookupParam(valueOrLabel, searchItem, param, 0, _
                               matchStrategy, useBinarySearch)
        End Function

        ''' <summary>
        '''     Search along one side (either the Values or Labels) of an SSRS
        '''     parameter (<c>param</c>), finding the Nth item (<c>nthMatch</c>)
        '''     that matches the <c>searchItem</c> using the specified
        '''     <c>matchStrategy</c>. Then return the object in the array on the
        '''     other side, but at the same position as the match.
        ''' </summary>
        ''' <param name="valueOrLabel">
        '''     Either the word <c>"value"</c> or <c>"label"</c> as a string
        '''     (using any case). If <c>"value"</c> is passed, the
        '''     <c>param</c>'s Values are searched for matches and the its Label
        '''     at the last matching position is returned. If <c>"label"</c> is
        '''     passed, searches the Labels and returns a Value.
        ''' </param>
        ''' <param name="searchItem">
        '''     The thing being searched for in the param. This is usually the
        '''     same type as the elements being searched in the <c>param</c>,
        '''     but this isn't necessary, provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="param">
        '''     An SSRS parameter containing both Values and Labels. A
        '''     single-value param is acceptable, and any type is fine,
        '''     provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="nthMatch">
        '''     When <c>nthMatch</c> matches are found, the value/label (as
        '''     appropriate) at the same position as the <c>nthMatch</c> is
        '''     returned. If there are fewer than <c>nthMatch</c> matches,
        '''     <c>Nothing</c> is returned.
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
        '''     For more information on match-strategies, see the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="useBinarySearch">
        '''     Pass <c>True</c> if the <c>param</c> is sorted by its
        '''     <c>valueOrLabel</c>. If so, binary-search (O(log(n))) will be
        '''     used. (Otherwise O(n).) Note that an SSRS query is gauranteed to
        '''     be sorted by its first field iff it's in its own query group
        '''     (typically this will mean the field name is repeated above
        '''     the field).
        ''' </param>
        ''' <returns>
        '''     <para>
        '''         If a match is found: The Label/Value in the same position in
        '''         the <c>param</c> as the Value/Label that matched.
        '''     </para>
        '''     <para>If a match is not found: <c>Nothing</c></para>
        ''' </returns>
        ''' <exception cref="ArgumentException">
        '''     Thrown if a 'contains', 'regex' or 'starts-with' match-strategy
        '''     is selected, but either the <c>searchItem</c> or the param's
        '''     Values/Labels (whichever is being searched) is not meaningfully
        '''     representable as a String.
        ''' </exception>
        ''' <exception cref="InvalidBinarySearchException">
        '''     Thrown if <c>useBinarySearch:=True</c> is passed with
        '''     <c>matchStrategy:="C"C</c> or <c>matchStrategy:="R"C</c>,
        '''     because there is no ordering of the parameter that would allow
        '''     a binary search to run without skipping elements.
        ''' </exception>
        ''' <example> Find the ID of the student with a certain Learner Code
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "a34", "c14"]
        ''' }
        ''' >>> LookupParam("label", "a", param, 2, "S"C, True)
        ''' 3456
        ''' >>> LookupParam("label", {"not", "stringifiable"}, param, 1, "S"C, False)
        ''' ArgumentException
        ''' >>> LookupParam("label", "1", param, 1, "C"C)
        ''' InvalidBinarySearchException
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupParam(valueOrLabel As String, _
                                     searchItem As Object, _
                                     param As Object, _
                                     nthMatch As Integer, _
                                     matchStrategy As Char, _
                                     useBinarySearch As Boolean) _
                                     As Object
            Dim searches As Object()
            Dim results As Object()

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
                ' Multivalue params mustn't be empty in the current SSRS version
                Return Nothing
            End If

            ' Throw an error if the caller has used incompatible options
            If Not matchStrategy.Equals("E"C) Then
                ThrowUnlessStringifiable(searches, searchItem, matchStrategy)
            End If

            If useBinarySearch Then
                Return BinSearchParam(searches, results, searchItem, _
                                      matchStrategy, nthMatch)
            End If

            Dim i As Integer = Array.FindIndex(searches, _
                    ParamPredicator.Create(searchItem, matchStrategy, nthMatch) _
                              .predicate)
            If i < 0 Then Return Nothing

            Return results(i)
        End Function

        ''' <summary>
        '''     Get the Nth item from an SSRS parameter's Values (as opposed to
        '''     its Labels). For full parameter documentation, see the delegate
        '''     function
        '''     <see href="#lookupnthparam-valueorlabel-number-param-"/>.
        ''' </summary>
        ''' <param name="number"/>
        ''' <param name="param"/>
        ''' <example> Find the ID of the 2nd student in the parameter
        '''     <code>
        ''' >>> param = { Value: [0123, 3456, 6789] }
        ''' >>> LookupNthParam(2, param)
        ''' 3456
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupNthParam(number As Integer, param As Object) _
                                       As Object
            Return LookupNthParam("value", number, param)
        End Function

        ''' <summary>
        '''     Get the Nth Value or Label from an SSRS parameter
        ''' </summary>
        ''' <param name="valueOrLabel">
        '''     Either the word <c>"value"</c> or <c>"label"</c> as a string
        '''     (using any case). If <c>"value"</c> is passed, the
        '''     <c>param</c>'s <c>number</c> Value is returned. If
        '''     <c>"label"</c> is passed, the <c>param</c>'s <c>number</c>
        '''     Label is returned.
        ''' </param>
        ''' <param name="number">
        '''     The position in the param from which the Value/Label will be
        '''     returned
        ''' </param>
        ''' <param name="param">
        '''     An SSRS parameter. It may not be a single-value param, but it
        '''     may have any type. It is acceptable for the <c>param</c> to not
        '''     have Labels, but if so, the value of <c>valueOrLabel</c> must be
        '''     <c>"value"</c>.
        ''' </param>
        ''' <example> Find the learnerCode of the 2nd student in a parameter
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123, 3456, 6789],
        '''     Label: ["a01", "b34", "c67"
        ''' }
        ''' >>> LookupNthParam("label", 2, param)
        '''         3456
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
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
        '''     Get the 'partners' of all items in either the Values or Labels
        '''     of an SSRS parameter that equal the <c>searchItem</c>. For full
        '''     parameter documentation, see the delegate function
        '''     <see href="#lookupallmatchingparams-valueorlabel-searchitem-param-matchstrategy-"/>.
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="searchItem"/>
        ''' <param name="param"/>
        Public Function LookupAllMatchingParams(valueOrLabel As String, _
                                                searchItem As Object, _
                                                param As Object) As Object()
            Return LookupAllMatchingParams(valueOrLabel, searchItem, param, "E"C)
        End Function

        ''' <summary>
        '''     Get the 'partners' of all items in either the Values or Labels
        '''     of an SSRS parameter that equal the <c>searchItem</c> using the
        '''     specified <c>matchStrategy</c>. Returns the results in an array.
        ''' </summary>
        ''' <param name="valueOrLabel">
        '''     Either the word <c>"value"</c> or <c>"label"</c> as a string
        '''     (using any case). If <c>"value"</c> is passed, the
        '''     <c>param</c>'s Values are searched for matches and the its
        '''     Labels at the matching positions are returned. If
        '''     <c>"label"</c> is passed, searches the Labels and returns the
        '''     corresponding Values.
        ''' </param>
        ''' <param name="searchItem">
        '''     The thing being searched for in the param. This is usually the
        '''     same type as the elements being searched in the <c>param</c>,
        '''     but this isn't necessary, provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="param">
        '''     An SSRS parameter containing both Values and Labels. A
        '''     single-value param is acceptable, and any type is fine,
        '''     provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
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
        '''     For more information on match-strategies, see the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <returns>
        '''     An array of the labels/values in the same positions in the param
        '''     as the values/labels that matched. (If none matched, the array
        '''     is empty.)
        ''' </returns>
        ''' <exception cref="ArgumentException">
        '''     Thrown if a 'contains' or 'starts-with' match-strategy is
        '''     selected, but either the <c>searchItem</c> or the param's
        '''     values/labels (whichever is being searched) is not meaningfully
        '''     representable as a String.
        ''' </exception>
        ''' <example> Find the IDs of all the students whose last name ends with
        '''     "Bain"
        '''     <code>
        ''' >>> param = {
        '''     Value: [0123,      3456,    6789,     8901],
        '''     Label: ["MacBain", "Smith", "McBain", "Jones"]
        ''' }
        ''' >>> LookupAllMatchingParams("label", "Bain$", param, "R"C)
        '''     {"MacBain", "McBain"}
        ''' >>> LookupAllMatchingParams("label", {"not", "stringifiable"}, param, "S"C, False)
        ''' ArgumentException
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function LookupAllMatchingParams(valueOrLabel As String, _
                                                searchItem As Object, _
                                                param As Object, _
                                                matchStrategy As Char) _
                                                As Object()
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
                ' Multivalue params mustn't be empty in the current SSRS version
                Return {}
            End If

            ' Throw an error if the caller has used incompatible options
            If Not matchStrategy.Equals("E"C) Then
                ThrowUnlessStringifiable(searches, searchItem, matchStrategy)
            End If

            Dim predicate As Predicate(Of Object) = _
                ParamPredicator.Create(searchItem, matchStrategy).predicate

            For i As Integer = 0 To searches.GetUpperBound(0)
                If predicate.Invoke(searches(i)) Then
                    finds.Add(results(i))
                End If
            Next i

            Return finds.ToArray()
        End Function
    End Module
End Namespace
