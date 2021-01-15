Namespace SSRSCode
    ''' <summary>
    '''     This module contains functions that are only expected to be useful
    '''     in the bi-termly cycle of reports for parents.
    ''' </summary>
    Public Module Report6

        ''' <summary>
        '''     Sanatises a mark. Trims unnecessary decimal points and zeros from
        '''     floats in Strings, or else, if the columns contains predicted
        '''     grades, the points are appended to the mark, with a `#` seperator
        ''' </summary>
        ''' <param name="mark">
        '''     A mark, stored as a String
        ''' </param>
        ''' <param name="points">
        '''     Points corresponding to a mark's position on a gradescale. If there
        '''     is no relevant gradescale, <c>points</c> will be blank
        ''' </param>
        ''' <param name="templateColName">
        '''     The name of the column in the template. Assumes that predicted
        '''     grade columns contain "_PredGrade"
        ''' </param>
        ''' <returns>
        '''     If the mark is a float, e.g. "12.00000" => "12"
        '''     Else if column name containes "_PredGrade", e.g. "A*", "20" => "A*#20"
        '''     Else just the mark, unchanged
        ''' </returns>
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

        ''' <summary>
        '''     Exits early if <c>column</c> is empty, but otherwise delegates
        '''     to <see href="#lookupparam-valueorlabel-searchitem-param-matchstrategy-"/>,
        '''     passing <c>grouplearner &amp; param</c>.
        '''     This function exists to prevent logical errors in cases where
        '''     either the 'starts-with' or 'contains' match-strategy is used,
        '''     or <c>column</c> is empty. (If not for the early return, false
        '''     positive matches would be found from any column.)
        ''' </summary>
        ''' <param name="valueOrLabel"/>
        ''' <param name="groupLearner">
        '''     The concatenation of a group code and a learner code.
        ''' </param>
        ''' <param name="column">
        '''     The name of the Screen Column. If empty, <c>Nothing</c> is returned
        ''' </param>
        ''' <param name="param"/>
        ''' <param name="matchStrategy"/>
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

        ''' <summary>
        '''     Calculate the average value of a series of 0 or more values in a string
        ''' </summary>
        ''' <param name="vals">
        '''     A string containing 0 or more numeric values delimited by `, ` 
        ''' </param>
        ''' <param name="valIfBlank">
        '''     The value to return if <c>vals</c> is empty. Optional; the
        '''     default is 40.
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
    End Module
End Namespace
