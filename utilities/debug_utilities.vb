'Debugging Utilities------------------------------------------------------------

Public Function typeAsString(obj As Object) As String
    Return obj.GetType().FullName
End Function

Public Function listProperties(obj As Object) As String
    listProperties = ""
    Dim pdc As System.ComponentModel.PropertyDescriptorCollection
    pdc = System.ComponentModel.TypeDescriptor.GetProperties(obj.GetType)
    For Each pd As System.ComponentModel.PropertyDescriptor In pdc
        listProperties = listProperties & pd.Name & ", "
    Next
End Function

Public Function paramToS(param As Object) As String
    Dim labels(param.Count - 1) As Object
    Array.copy(param.Label, labels, param.Count)
    Array.sort(labels)
    For Each label As Object In labels
        paramToS += label  & ": " & lookupParam("label", label, param) & ", "
    Next label
End Function
