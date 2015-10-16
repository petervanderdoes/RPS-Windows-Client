Public Class Helper
    ''' <summary>Trims the trailing slash off a string if it exists</summary>
    ''' <param name="input_string">String to trim</param>
    ''' <remarks>Trims the trailing slash off a string if it exists.  The primary use for this is validating file folder paths, i.e. changing "C:\Temp\" to "C:\Temp"
    ''' </remarks>
    ''' <returns>String</returns>
    ''' <example>
    ''' MsgBox(TrimTrailingSlash("C:\Windows\Temp\"))
    ''' </example>
    Public Shared Function trimTrailingSlash(ByVal input_string As String) As String
        'trim any blanks
        input_string = input_string.Trim
        Dim output_string As String 'the string to return
        'see if the last character is a \ or not
        If input_string.Trim.Length > 0 And input_string.Substring(input_string.Length - 1, 1) = "\" Then
            output_string = input_string.Substring(0, input_string.Length - 1)
        Else
            output_string = input_string
        End If
        Return output_string
    End Function
End Class
