Namespace SSRSCode
    ''' <summary>
    '''     Create a colour scale - e.g. red-orange-green - in an SSRS report
    '''     (one scale per report) and retrieve colours X% of the way along it. 
    ''' </summary>
    ''' <remarks>
    '''     Usually it's unwise to expect an SSRS report to correctly maintain
    '''     the state of custom code objects - they should be used as pure
    '''     functions. In this case though, the global variable
    '''     <c>header_colour_scale</c> is used to store state. This is safe
    '''     for the reasons explained in the remarks for the function
    '''     <see href="#colourfromscale-fraction-first-second-third-"/>.
    ''' </remarks>
    Module ColourScale

        Dim header_colour_scale As ColourScale

        ''' <summary>
        '''     Retreives a colour <c>fraction</c> of the way along a colour
        '''     scale. If the scale doesn't exist, it creates it using 2 or 3
        '''     hexadecimal colours specified by the <c>first</c>, <c>second</c>
        '''     (and optionally <c>third</c>) arguments.
        ''' </summary>
        ''' <param name="fraction">
        '''     A double, f where 0.0 &lt;= f &lt;= 1.0, representing the
        '''     fraction along the scale of the desired colour. See examples.
        ''' </param>
        ''' <param name="first">
        '''     The colour at 0% on the colour scale.
        '''     A hexadecimal colour stored in a String. The hex value should
        '''     have all 6 characters, not just 3. e.g. <c>"#00ff00"</c>, not
        '''     <c>"#0f0"</c>, for the colour green.
        ''' </param>
        ''' <param name="second">
        '''     A hexadecimal colour stored in a String. Will be the colour at
        '''     100% on the colour scale if the <c>third</c> parameter is
        '''     omitted. Otherwise it will be at 50%.
        ''' </param>
        ''' <param name="third">
        '''     Optional. A hexadecimal colour stored in a String. Will be the
        '''     colour at 100% on the colour scale if supplied.
        '''     </param>
        ''' <returns></returns>
        ''' <remarks>
        '''     Checks whether the colour scale has been lost (or never existed)
        '''     before it uses it, and recreates it if it necessary. This is why
        '''     the data to recreate the scale must always be passed when the
        '''     caller attempts to retrieve a colour.
        ''' </remarks>
        ''' <example> Using an red-orange-green colour scale
        '''     <code>
        '''         >>> ColourFromScale(1, "#ff0000", "#ffbf00", "#00ff00")
        '''         "#00ff00"
        '''         >>> ColourFromScale(0.5, "#ff0000", "#ffbf00", "#00ff00")
        '''         "#ffbf00"
        '''         >>> ColourFromScale(0.2, "#ff0000", "#ffbf00", "#00ff00")
        '''         "#ff4d00"
        '''     </code>
        ''' </example>
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
    End Module
End Namespace
