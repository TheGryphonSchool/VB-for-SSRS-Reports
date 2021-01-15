Namespace SSRSCode
    ''' <summary>
    '''     A collection of pure functions that test whether SSRS parameters
    '''     have elements that match supplied criteria, often using various
    '''     matching strategies.
    ''' </summary>
    ''' <remarks><para>
    '''     Some functions accepts a 'match-strategy' as a parameter. In these
    '''     cases, read the remarks provided here:
    '''     <see href="#lookupparams"/>.</para><para>
    '''     These functions are grouped (and distinguished from those in the
    '''     <see href="#lookupparams"/> module) they do not retrieve items from
    '''     parameters, they just test for matches, returning either a boolean,
    '''     or an Integer count of the matches.
    ''' </para></remarks>
    Public Module SearchParams
        ' Dependent on utilities/param_helpers.vb

        ''' <summary>
        '''     Tests whether an <c>searchItem</c> is in an array, using the
        '''     specified <c>matchStrategy</c>.
        ''' </summary>
        ''' <param name="searchItem">
        '''     The thing being searched for in the param. This is usually the
        '''     same type as the elements being searched in the <c>param</c>,
        '''     but this isn't necessary, provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="arry">
        '''     The array being searched in. Any type is fine, provided you
        '''     heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="matchStrategy">
        '''     A character denoting the match-strategy; one of:
        '''     <list type="bullet">
        '''         <item><term>E</term><description>Equals (the default)</description></item>
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
        '''     <c>True</c>, if the item is found in the array; <c>False</c>
        '''     otherwise
        ''' </returns>
        ''' <exception cref="ArgumentException">
        '''     Thrown if a 'contains' or 'starts-with' match-strategy is
        '''     selected, but either the <c>searchItem</c> or the param's
        '''     values/labels (whichever is being searched) is not meaningfully
        '''     representable as a String.
        ''' </exception>
        ''' <example> Check whether a param contains certain subjects:
        '''     <code>
        ''' >>> param = {
        '''     Value: ["Sc",      "Ar",  "Ar3D"],
        '''     Label: ["Science", "Art", "Art 3D"]
        ''' }
        ''' >>> IsInArray("Sience", param.Label)
        ''' True
        ''' >>> IsInArray("English", param.Label, "E"C)
        ''' False
        ''' >>> IsInArray("Ar", param.Label, "S"C)
        ''' True
        ''' >>> IsInArray("3", param.Value, "C"C)
        ''' True
        ''' >>> IsInArray(".{3,}", param.Value, "R"C)
        ''' True
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function IsInArray(searchItem As Object, _
                                  arry As Object(), _
                                  Optional matchStrategy As Char = "E"C) _
                                  As Boolean
            Return Array.Exists( _
                arry, _
                ParamPredicator.Create(searchItem, matchStrategy).predicate _
            )
        End Function

        ''' <summary>
        '''     Delegates to <see href="#countinarray-searchitem-matchstrategy-arry-"/>,
        '''     using `equals` as a match strategy.
        ''' </summary>
        ''' <param name="searchItem"/>
        ''' <param name="arry"/>
        Public Function CountInArray(searchItem As Object, _
                                     arry As Object()) As Integer
            Return CountInArray(searchItem, "E"C, arry)
        End Function

        ''' <summary>
        '''     Counts all items in an array that match the <c>searchItem</c>
        '''     (using the specified <c>searchItem</c>).
        ''' </summary>
        ''' <param name="searchItem">
        '''     The thing being searched for in the param. This is usually the
        '''     same type as the elements being searched in the <c>param</c>,
        '''     but this isn't necessary, provided you heed the remarks here:
        '''     <see href="#lookupparams"/>.
        ''' </param>
        ''' <param name="arry">
        '''     The array being searched in. Typically either the Values or
        '''     Labels of a param, but if so, note that the param must be
        '''     multivalue. Any type is fine, provided you heed the remarks
        '''     here:
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
        ''' <returns>The number of matches</returns>
        ''' <exception cref="ArgumentException">
        '''     Thrown if a 'contains', 'regex' or 'starts-with' match-strategy
        '''     is selected, but either the <c>searchItem</c> or the param's
        '''     values/labels (whichever is being searched) is not meaningfully
        '''     representable as a String.
        ''' </exception>
        ''' <remarks>
        '''     Beware that if the <c>searchItem</c> is an Integer from a query,
        '''     it will be of type Long (int64), meaning that params of type
        '''     Integer (Int32) will not match them unless they are cast to
        '''     Integer.
        ''' </remarks>
        ''' <example>
        '''     Check whether how many of certain subjects a param contains:
        '''     <code>
        ''' >>> param = {
        '''     Value: ["Sc",      "Ar",  "Ar3D"],
        '''     Label: ["Science", "Art", "Art 3D"]
        ''' }
        ''' >>> CountInArray("English", param.Label, "E"C)
        ''' 0
        ''' >>> CountInArray("Ar", param.Label, "S"C)
        ''' 2
        ''' >>> CountInArray("3", param.Value, "C"C)
        ''' 1
        ''' >>> CountInArray(".{3,}", param.Label, "R"C)
        ''' 3
        ''' >>> CountInArray({"not", "stringifiable"}, param.Value, "C"C)
        ''' ArgumentException
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function CountInArray(searchItem As Object, _
                                     matchStrategy As Char, _
                                     arry As Object()) As Integer
            Dim foundCount As Integer = 0

            ' Throw an error if the caller has used incompatible options
            If Not matchStrategy.Equals("E"C) Then
                ThrowUnlessStringifiable(arry, searchItem, matchStrategy)
            End If

            Dim predicate As Predicate(Of Object) = _
                ParamPredicator.Create(searchItem, matchStrategy).predicate

            For i As Integer = 0 To arry.GetUpperBound(0)
                If predicate.Invoke(arry(i)) Then
                    foundCount += 1
                End If
            Next i
            Return foundCount
        End Function

        ''' <summary>
        '''     Tests whether the parameter has a position where its value and
        '''     label respectively match the <c>value</c> and <c>label</c>
        '''     supplied.
        ''' </summary>
        ''' <param name="value">
        '''     The object being searched for in the <c>param</c>'s Values
        ''' </param>
        ''' <param name="label">
        '''     The object being searched for in the <c>param</c>'s Labels.
        '''     Although this can be any object, SSRS parameter Labels are
        '''     always stored as Strings, so #ToString() will be called on this
        '''     value, and the result tested for equality.
        ''' </param>
        ''' <param name="param">
        '''     An SSRS parameter containing both Values and Labels.
        '''     A single-value param is acceptable, and any type is fine.
        ''' </param>
        ''' <returns>
        '''     True if matching value/label pair is found; false otherwise
        ''' </returns>
        ''' <example>
        '''     Testing whether a learner/group are in a param:
        '''     <code>
        ''' >>> param = {
        '''     Value: ["12Sc/A", "12Sc/B", "12Sc/A"],
        '''     Label: ["Jones",  "Smith",  "Smith"]
        ''' }
        ''' >>> VLPairIsInParam("12Sc/A", "Smith", param)
        ''' True
        ''' >>> VLPairIsInParam("12Sc/B", "Jones", param)
        ''' False
        '''     </code>
        '''     The param is defined in JavaScript object notation for
        '''     illustrative purposes - this code isn't valid VB.
        ''' </example>
        Public Function VLPairIsInParam(value As Object, label As Object, _
                                        param As Object) As Boolean
            Dim stringLabel As String = label.ToString
            If Not param.IsMultiValue Then Return param.Value.Equals(value) _
                AndAlso param.Label.Equals(stringLabel)

            Dim values = param.Value
            ' Search Values first because they're more likely to be unique
            For i As Integer = 0 To param.Count - 1
                If values(i).Equals(value) AndAlso _
                    param.Label(i).Equals(stringLabel) Then Return True
            Next
            Return False
        End Function
    End Module
End Namespace
