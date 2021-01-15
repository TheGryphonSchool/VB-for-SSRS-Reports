Public Class MockParam
    Public Value() As IComparable
    Public Label() As IComparable
    Public IsMultiValue As Boolean = True
    Public Count As Integer

    Public Sub New(vals() As IComparable, labs() As IComparable)
        Count = vals.Length

        If labs.Length <> Count Then
            Throw New ArgumentException
        End If

        Value = vals
        Label = labs
    End Sub
End Class

Public Class SingleValueMockParam
    Shadows Public Value As IComparable
    Shadows Public Label As IComparable
    Shadows Public IsMultiValue As Boolean = False
    Shadows Public Count As Integer = 1

    Public Sub New(value As IComparable, label As String)
        Me.Value = value
        Me.Label = label
    End Sub
End Class
