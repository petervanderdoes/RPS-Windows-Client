Public Class Helper
    ''' <summary>Trims the trailing slash off a string if it exists</summary>
    ''' <param name="InputString">String to trim</param>
    ''' <remarks>Trims the trailing slash off a string if it exists.  The primary use for this is validating file folder paths, i.e. changing "C:\Temp\" to "C:\Temp"
    ''' </remarks>
    ''' <returns>String</returns>
    ''' <example>
    ''' MsgBox(TrimTrailingSlash("C:\Windows\Temp\"))
    ''' </example>
    Public Shared Function TrimTrailingSlash(ByVal InputString As String) As String
        'trim any blanks
        InputString = InputString.Trim
        Dim OutputString As String = "" 'the string to return
        'see if the last character is a \ or not
        If InputString.Trim.Length > 0 And InputString.Substring(InputString.Length - 1, 1) = "\" Then
            OutputString = InputString.Substring(0, InputString.Length - 1)
        Else
            OutputString = InputString
        End If
        Return OutputString
    End Function
End Class
