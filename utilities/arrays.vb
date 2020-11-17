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
