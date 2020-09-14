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

Public Function averageArray(items As Object()) As Double
	Dim sum as Double = 0
	Dim count as Integer = 0
	For Each item As Double In items
		If Not IsNumeric(item) Then
            Try
                item = CDbl(item) 
            Catch _ex As Exception
                item = 0
            End Try
        End If
        sum += item
		count += 1
	Next item
	If count = 0 Then
        Return 0
    End If
    Return sum / count
End Function
