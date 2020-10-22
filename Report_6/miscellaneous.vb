Public Function roundIfFloat(float As String) As String
    If Not float.Contains(".0") Then
        Return float
    End If
    Return CStr(CInt(float))
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
        If not Double.TryParse(val, current) Then
            return valIfBlank
        End If
        sum += current
        count += 1
    Next val
    Return sum / count
End Function
