Public Function roundIfFloat(float As String) As String
    If Not float.Contains(".0") Then
        Return float
    End If
    Return CStr(CInt(float))
End Function
