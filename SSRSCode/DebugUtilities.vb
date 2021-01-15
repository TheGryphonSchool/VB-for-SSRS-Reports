Namespace SSRSCode
    ''' <summary>
    '''     These ulitiles are useful in debugging an SSRS report, but should
    '''     not be included in the shipped version of any report.
    ''' </summary>
    Module DebugUtilities
        Public Function TypeAsString(obj As Object) As String
            Return obj.GetType().FullName
        End Function

        Public Function ListProperties(obj As Object) As String
            ListProperties = ""
            Dim pdc As ComponentModel.PropertyDescriptorCollection
            pdc = ComponentModel.TypeDescriptor.GetProperties(obj.GetType)
            For Each pd As ComponentModel.PropertyDescriptor In pdc
                ListProperties = ListProperties & pd.Name & ", "
            Next
        End Function

        Public Function ParamToS(param As Object, _
                                 Optional sortByLabel As Boolean = True) _
                                 As String
            Dim labels(param.Count - 1) As Object
            ParamToS = ""
            If sortByLabel Then
                Array.Copy(param.Label, labels, param.Count)
                Array.Sort(labels)
            Else
                labels = param.Label
            End If
            For Each label As Object In labels
                ParamToS += label & ": " & _
                            LookupParam("label", label, param) & ", "
            Next label
        End Function
    End Module
End Namespace
