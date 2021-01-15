Public Function RemoveDuplicates(items As Object()) As Object()
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

Public Function ArrayMerge(leftArray As Object(), _
        appendage As Object) As Object()
    Return ArrayMerge(leftArray, New Object() {appendage})
End Function

Public Function ArrayMerge(prependage As Object, _
        rightArray As Object()) As Object()
    Return ArrayMerge(New Object() {prependage}, rightArray)
End Function

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
Friend Function BinSearchParam(searches() As Object, _
                               results() As Object, _
                               searchItem As Object, _
                               matchStrategy As Char, _
                               nthMatch As Integer) As Object
    Dim comp As Object = GetComparer(matchStrategy)
    Dim randMatch As Integer = _
        Array.BinarySearch(searches, searchItem, comp)
    Return NthMatchFromAMatch(searches, results, searchItem, nthMatch, _
                              randMatch, comp)
End Function

Private Function NthMatchFromAMatch(searches() As Object, _
                                    results() As Object, _
                                    searchItem As Object, _
                                    nthMatch As Integer, _
                                    matchIndex As Integer, _
                                    comp As Object) As Object

    If matchIndex < 0 Then Return Nothing
    If nthMatch = 0 Then Return results(matchIndex)

    Dim firstIndex As Integer = matchIndex
    While firstIndex > 0 AndAlso _
            AreComparable(searchItem, searches(firstIndex - 1), comp)
        firstIndex -= 1
    End While
    Dim nthMatchingIndex As Integer = firstIndex + nthMatch - 1
    If nthMatchingIndex >= searches.Length OrElse _
        nthMatchingIndex > matchIndex AndAlso _
        Not AreComparable(searchItem, searches(nthMatchingIndex), comp) Then
        Return Nothing
    End If

    Return results(nthMatchingIndex)
End Function

Private Function AreComparable(searchItem As Object, _
                               candidate As Object, _
                               comp As Object) As Boolean
    If comp Is Nothing Then
        Return searchItem.CompareTo(candidate) = 0
    End If

    Return comp.PredicateCompare(candidate, searchItem)
End Function

Private Function GetComparer(matchStrategy As Char) As Object
    Select Case matchStrategy
        Case "C"C ' Contains
            Throw New InvalidBinarySearchException("Contains")
        Case "R"C ' Regular Expression
            Throw New InvalidBinarySearchException("Regular Expression")
        Case "S"C ' Starts-with
            Return New StartsWithComparer()
        Case Else ' Equals (use default CompareTo implementation)
            Return Nothing
    End Select
End Function

Private Class StartsWithComparer
    Implements System.Collections.IComparer

    Public Function Compare(candidate As Object, searchItem As Object) _
            As Integer Implements System.Collections.IComparer.Compare
        If PredicateCompare(candidate, searchItem) Then Return 0

        Return candidate.CompareTo(searchItem)
    End Function

    Public Function PredicateCompare(candidate As Object, _
                                     searchItem As Object) As Boolean
        Return candidate.ToString().StartsWith(searchItem.ToString())
    End Function
End Class

Private Class InvalidBinarySearchException
    Inherits ArgumentException

    Public Sub New(strategyName As String)
        MyBase.New("You may not use binary-search with the '" _
                   & strategyName & "' match-strategy. Use " _
                   & "'Equals' or 'Starts-with' instead.")
    End Sub
End Class


Dim header_colour_scale As ColourScale

Public Function ColourFromScale(fraction As Double, _
                                first As String, _
                                second As String, _
                                third As String) As String
    If header_colour_scale Is Nothing Then
        header_colour_scale = New ColourScale(first, second, third)
    End If
    Return header_colour_scale.GetColour(fraction)
End Function

Public Class ColourScale
    Private ReadOnly scale As New _
        System.Collections.Generic.List(Of Integer())

    Public Sub New(first As String, _
               second As String, _
               Optional third As String = "", _
               Optional fourth As String = "", _
               Optional fifth As String = "")
        For Each nth As String In _
                New String(4) {first, second, third, fourth, fifth}
            If nth Is "" Then Exit For

            AddToScale(nth)
        Next nth
    End Sub

    Public Function GetColour(fraction As Double)
        Dim last_index As Integer = scale.Count - 1
        Dim start As Integer
        If fraction >= 1.0 Then
            Return MixTwoColours(1.0, last_index - 1)
        End If
        start = CInt(Math.Floor(fraction * last_index))
        Return MixTwoColours(fraction * last_index - start, start)
    End Function

    Private Sub AddToScale(hexColour As String)
        Dim rgb(2) As Integer
        hexColour = hexColour.Replace("#", "")
        For i As Integer = 0 To 2
            rgb(i) = Convert.ToInt32(hexColour.Substring(i * 2, 2), 16)
        Next i
        scale.Add(rgb)
    End Sub

    Private Function MixTwoColours(fraction As Double, _
                                   start_index As Integer) As String
        Dim starts As Integer
        Dim ends As Integer
        MixTwoColours = "#"
        For i As Integer = 0 To 2
            starts = scale.Item(start_index)(i)
            ends = scale.Item(start_index + 1)(i)
            MixTwoColours += _
                Hex(CInt(starts + fraction * (ends - starts))) _
                .PadLeft(2, "0")
        Next i
    End Function
End Class

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object) As Object
    Return LookupParam(valueOrLabel, searchItem, param, 1, "E"C, False)
End Function

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object, _
                            nthMatch As Integer) As Object
    Return LookupParam(valueOrLabel, searchItem, param, nthMatch, "E"C, _
                       False)
End Function

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object, _
                            matchStrategy As Char) As Object
    Return LookupParam(valueOrLabel, searchItem, param, 1, _
                       matchStrategy, False)
End Function

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object, _
                            useBinarySearch As Boolean) As Object
    Return LookupParam(valueOrLabel, searchItem, param, 0, "E"C, _
                       useBinarySearch)
End Function

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object, _
                            nthMatch As Integer, _
                            matchStrategy As Char) As Object
    Return LookupParam(valueOrLabel, searchItem, param, nthMatch, _
                       matchStrategy, False)
End Function

Public Function LookupParam(valueOrLabel As String, _
                            searchItem As Object, _
                            param As Object, _
                            matchStrategy As Char, _
                            useBinarySearch As Boolean) As Object
    Return LookupParam(valueOrLabel, searchItem, param, 0, _
                       matchStrategy, useBinarySearch)
End Function

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
        Return Nothing
    End If

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

Public Function LookupNthParam(number As Integer, param As Object) _
                               As Object
    Return LookupNthParam("value", number, param)
End Function

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

Public Function LookupAllMatchingParams(valueOrLabel As String, _
                                        searchItem As Object, _
                                        param As Object) As Object()
    Return LookupAllMatchingParams(valueOrLabel, searchItem, param, "E"C)
End Function

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
        Return {}
    End If

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
Friend Sub ThrowUnlessStringifiable(searches As Object(), _
                                    searchItem As Object, _
                                    matchStrategy As Char)
    Dim varMessages = {{searchItem, "The search item"}, _
        {searches(0), "The items in the searched side of the parameter"}}
    For i As Integer = 0 To 1
        If IsNotStringifiable(varMessages(i, 0)) Then
            Throw New ArgumentException(MatchStrategyExceptionMessage( _
                varMessages(i, 1), matchStrategy, _
                varMessages(i, 0).GetType().Name))
        End If
    Next i
End Sub

Private Function IsNotStringifiable(ssrsObject As Object) As Boolean
    Return IsArray(ssrsObject) OrElse TypeOf ssrsObject Is Boolean
End Function

Private Function MatchStrategyExceptionMessage(problematicParam As String, _
                                               matchStrategy As Char, _
                                               badType As String) As String
    Dim strategyDescription As String
    Select Case matchStrategy
        Case "C"C
            strategyDescription = "'Contains' ('C')"
        Case "S"C
            strategyDescription = "'Starts-with' ('S')"
        Case Else
            strategyDescription = "'Regular Expression' ('R')"
    End Select
    Return problematicParam & " cannot be of type " & badType & _
        " to use the match strategy " & strategyDescription & _
        ". Omit the matchStrategy argument to use exact matching."
End Function
Friend Class ParamPredicator
    Public predicate As Predicate(Of Object)
    Protected ReadOnly searchItem As Object
    Protected ReadOnly searchString As String
    Protected reg As Text.RegularExpressions.Regex

    Public Shared Function Create(searchItem As Object, _
                                  matchStrategy As Char, _
                                  Optional nthMatch As Integer = 1) _
                                  As ParamPredicator
        If nthMatch = 1 Then Return New ParamPredicator(searchItem, matchStrategy)

        Return New TargetedPredicator(searchItem, matchStrategy, nthMatch)
    End Function

    Protected Sub New(searchItem As Object, _
                      matchStrategy As Char)
        Me.searchItem = searchItem
        Me.searchString = searchItem.ToString
        Select Case matchStrategy
            Case "C"C ' Contains
                predicate = AddressOf Contains
            Case "R"C ' Regular Expression
                predicate = AddressOf Regex
            Case "S"C ' Starts-with
                predicate = AddressOf StartsWith
            Case Else ' Equals
                predicate = AddressOf EqualTo
        End Select
    End Sub

    Public Function Contains(candidate As Object) As Boolean
        Return candidate.ToString().Contains(searchString)
    End Function

    Friend Function StartsWith(candidate As Object) As Boolean
        Return candidate.ToString().StartsWith(searchString)
    End Function

    Friend Function EqualTo(candidate As Object) As Boolean
        Return searchItem.Equals(candidate)
    End Function

    Friend Function Regex(candidate As Object) As Boolean
        If IsNothing(reg) Then
            reg = New Text.RegularExpressions.Regex(searchString)
        End If
        Return reg.IsMatch(candidate.ToString())
    End Function

    Private Class TargetedPredicator
        Inherits ParamPredicator
        Public ReadOnly subPredicate As Predicate(Of Object)
        Protected ReadOnly nthMatch As Integer
        Protected matchCount As Integer = 0

        Sub New(searchItem As Object, _
                          matchStrategy As Char, _
                          Optional nthMatch As Integer = 1)
            MyBase.New(searchItem, matchStrategy)
            Me.nthMatch = nthMatch
            subPredicate = predicate
            predicate = AddressOf AdjustForCount
        End Sub

        Private Function AdjustForCount(candidate As Object) As Boolean
            If Not subPredicate.Invoke(candidate) Then Return False

            matchCount += 1
            Return matchCount >= nthMatch
        End Function
    End Class
End Class

Public Function CleanValue(mark As String, points As String, _
                           templateColName As String) As String
    If mark.Contains(".0") Then
        Return CStr(CInt(mark))
    End If
    If templateColName.Contains("_PredGrade") Then
        Return mark & "#" & points
    End If
    Return mark
End Function

Public Function LookupParamIfPresent(valueOrLabel As String, _
                                     groupLearner As String, _
                                     column As String, _
                                     param As Object, _
                                     Optional matchStrategy As Char = "E"C) _
                                     As String
    If column Is Nothing OrElse column = "" Then Return ""
    Return LookupParam(valueOrLabel, groupLearner & Column, _
                       param, matchStrategy)
End Function

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

Public Function IsInArray(searchItem As Object, _
                          arry As Object(), _
                          Optional matchStrategy As Char = "E"C) _
                          As Boolean
    Return Array.Exists( _
        arry, _
        ParamPredicator.Create(searchItem, matchStrategy).predicate _
    )
End Function

Public Function CountInArray(searchItem As Object, _
                             arry As Object()) As Integer
    Return CountInArray(searchItem, "E"C, arry)
End Function

Public Function CountInArray(searchItem As Object, _
                             matchStrategy As Char, _
                             arry As Object()) As Integer
    Dim foundCount As Integer = 0

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

Public Function VLPairIsInParam(value As Object, label As Object, _
                                param As Object) As Boolean
    Dim stringLabel As String = label.ToString
    If Not param.IsMultiValue Then Return param.Value.Equals(value) _
        AndAlso param.Label.Equals(stringLabel)

    Dim values = param.Value
    For i As Integer = 0 To param.Count - 1
        If values(i).Equals(value) AndAlso _
            param.Label(i).Equals(stringLabel) Then Return True
    Next
    Return False
End Function
