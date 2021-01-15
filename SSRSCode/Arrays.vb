Namespace SSRSCode
    ''' <summary>
    '''     A collection of purely functional utilities for working with arrays
    ''' </summary>
    Public Module Arrays
        ''' <summary>
        '''     Create a version of an array without duplicate elements
        ''' </summary>
        ''' <param name="items">An array of any type and size</param>
        ''' <returns>
        '''     An array of type <c>Object</c>, containing the non-duplicate
        '''     elements from <c>items</c> in the same order, but without any
        '''     elements that already appeared at lower indices.
        ''' </returns>
        ''' <example> Removing duplicate 2s from an array of Integers
        '''     <code> RemoveDuplicates({1, 2, 3, 2, 2}) => {1, 2, 3} </code>
        ''' </example>
        Public Function RemoveDuplicates(items As Object()) As Object()
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

        ''' <summary>
        '''     Attempts to sum the items in <c>nums</c> by parsing them as
        '''     integers.
        ''' </summary>
        ''' <param name="nums">
        '''     An array of any type and size. Will attempt to parse Strings and
        '''     Dates to Integers. Doubles will be treated as 0 unless they are
        '''     integer.
        ''' </param>
        ''' <returns>
        '''     An integer sum of the elements in <c>nums</c>. If no parses were
        '''     possible, 0 is returned.
        ''' </returns>
        ''' <example> Summing an array of Integers
        '''     <code> SumArray({1, 2}) => 3</code>
        ''' </example>
        ''' <example> Summing an array of integers stored as doubles
        '''     <code> SumArray({1.0, 2.0}) => 3</code>
        ''' </example>
        ''' <example> Summing an array of Doubles counts non-integers as 0
        '''     <code> SumArray({1.0, 2.5}) => 1</code>
        ''' </example>
        ''' <example> Parsing an array of Strings containing integers
        '''     <code> SumArray({"1", "2"}) => 3</code>
        ''' </example>
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
        '''     Appends an object onto an array (without altering it) and
        '''     returns the new array.
        ''' </summary>
        ''' <param name="leftArray">
        '''     An array, at the end of which, <c>appendage</c> will be
        '''     appended.
        ''' </param>
        ''' <param name="appendage">
        '''     An object to append to the array <c>leftArray</c>
        ''' </param>
        ''' <returns>
        '''      An array resulting form appending <c>appendage</c> to
        '''      <c>leftArry</c>.
        ''' </returns>
        Public Function ArrayMerge(leftArray As Object(), _
                appendage As Object) As Object()
            Return ArrayMerge(leftArray, New Object() {appendage})
        End Function

        ''' <summary>
        '''     Prepends an object at the start of an array (without altering
        '''     it) and returns the new array.
        ''' </summary>
        ''' <param name="rightArray">
        '''     An array, at the start of which, <c>prependage</c> will be
        '''     prepended.
        ''' </param>
        ''' <param name="prependage">
        '''     An object to prepend to the array <c>rightArray</c>
        ''' </param>
        ''' <returns>
        '''      An array resulting form appending <c>prependage</c> to
        '''      <c>leftArry</c>.
        ''' </returns>
        Public Function ArrayMerge(prependage As Object, _
                rightArray As Object()) As Object()
            Return ArrayMerge(New Object() {prependage}, rightArray)
        End Function

        ''' <summary>
        '''     Merges 2 arrays (without altering either) and returns the new
        '''     array. Will attempt a narrowing conversion (e.g. <c>String</c>
        '''     to <c>Integer</c>) if only one array is String(). If this
        '''     converion fails for any item, the opposing, widening converion
        '''     will be used. If the arrays are of differing, non-string types,
        '''     both will be converted to <c>String</c> before merging.
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
                "TryParse", New Type() {GetType(String), _
                destType.MakeByRefType} _
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
    End Module
End Namespace
