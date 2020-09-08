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

Public Function debugRankChecker(group_code As String) As String
    debugRankChecker = "BADDIES: "
    For Each bad_rank As Integer In rank_checkers(group_code).badRanks
        debugRankChecker += CStr(bad_rank) + ", "
    Next bad_rank 
End Function
